VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRecordChange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Record Number Changer"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   Icon            =   "frmRecordChange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   9975
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   7980
      TabIndex        =   40
      Top             =   4560
      Width           =   1695
   End
   Begin MSComctlLib.StatusBar stsStatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5505
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14975
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   5475
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   9915
      Begin VB.Frame fraChangeMap 
         Caption         =   "Change Map Number"
         Height          =   3675
         Left            =   4980
         TabIndex        =   41
         Top             =   0
         Width           =   4755
         Begin VB.CheckBox chkMapChange 
            Caption         =   "Change Teleport abilities in spells to new map."
            Height          =   375
            Index           =   3
            Left            =   180
            TabIndex        =   49
            Top             =   3000
            Value           =   1  'Checked
            Width           =   4455
         End
         Begin VB.CheckBox chkMapChange 
            Caption         =   "Change Teleport commands in textblocks to new map."
            Height          =   375
            Index           =   2
            Left            =   180
            TabIndex        =   48
            Top             =   2580
            Value           =   1  'Checked
            Width           =   4455
         End
         Begin VB.CheckBox chkMapChange 
            Caption         =   "If the new map on a map change exit now matches the rooms' map, change the exit type to normal."
            Height          =   375
            Index           =   1
            Left            =   420
            TabIndex        =   47
            Top             =   2160
            Value           =   1  'Checked
            Width           =   4215
         End
         Begin VB.CheckBox chkMapChange 
            Caption         =   "Change 'Map Change' exit types on rooms to new map."
            Height          =   375
            Index           =   0
            Left            =   180
            TabIndex        =   46
            Top             =   1740
            Value           =   1  'Checked
            Width           =   4455
         End
         Begin VB.TextBox txtChangeMap 
            Height          =   315
            Index           =   1
            Left            =   2460
            TabIndex        =   44
            Text            =   "0"
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txtChangeMap 
            Height          =   315
            Index           =   0
            Left            =   2460
            TabIndex        =   42
            Text            =   "0"
            Top             =   540
            Width           =   975
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "To Map:"
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
            Left            =   1380
            TabIndex        =   45
            Top             =   1080
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Change Map:"
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
            Left            =   960
            TabIndex        =   43
            Top             =   540
            Width           =   1140
         End
      End
      Begin VB.CommandButton cmdLogQ 
         Cancel          =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6960
         TabIndex        =   39
         Top             =   3960
         Width           =   315
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "&Start"
         Height          =   375
         Left            =   7920
         TabIndex        =   38
         Top             =   3960
         Width           =   1695
      End
      Begin VB.CommandButton cmdQ 
         BackColor       =   &H00C0C0FF&
         Caption         =   "?"
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
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   3900
         Width           =   315
      End
      Begin VB.CommandButton cmdLog 
         Caption         =   "View Log File *"
         Height          =   375
         Left            =   5100
         TabIndex        =   34
         Top             =   3960
         Width           =   1755
      End
      Begin VB.Frame fraRL 
         Caption         =   "List of Records From Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3675
         Left            =   60
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   4695
         Begin VB.CommandButton cmdListRemoveLine 
            Caption         =   "Remove Selected Line"
            Height          =   315
            Left            =   180
            TabIndex        =   36
            Top             =   300
            Width           =   2415
         End
         Begin VB.CommandButton cmdAddAll 
            Caption         =   "Add All"
            Height          =   375
            Left            =   1500
            TabIndex        =   35
            Top             =   3120
            Width           =   1095
         End
         Begin VB.ListBox lstRecordList 
            Height          =   3180
            Left            =   2760
            TabIndex        =   4
            Top             =   300
            Width           =   1755
         End
         Begin VB.CommandButton cmdHideRL 
            Caption         =   "Hide"
            Height          =   375
            Left            =   180
            TabIndex        =   3
            Top             =   3120
            Width           =   1035
         End
         Begin VB.Label Label1 
            Caption         =   $"frmRecordChange.frx":08CA
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2115
            Index           =   1
            Left            =   180
            TabIndex        =   5
            Top             =   840
            Width           =   2475
         End
      End
      Begin VB.ListBox lstChange 
         Height          =   3570
         Left            =   7980
         TabIndex        =   23
         Top             =   60
         Width           =   1755
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   315
         Left            =   6840
         TabIndex        =   22
         Top             =   2820
         Width           =   975
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "C&lear"
         Height          =   315
         Left            =   5040
         TabIndex        =   21
         Top             =   3360
         Width           =   1575
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add -->"
         Height          =   315
         Left            =   6840
         TabIndex        =   19
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtToEnd 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   2820
         Width           =   735
      End
      Begin VB.TextBox txtToStart 
         Height          =   315
         Left            =   5040
         TabIndex        =   17
         Text            =   "1"
         Top             =   2820
         Width           =   735
      End
      Begin VB.TextBox txtFromEnd 
         Height          =   315
         Left            =   5880
         TabIndex        =   16
         Text            =   "0"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtFromStart 
         Height          =   315
         Left            =   5040
         TabIndex        =   15
         Text            =   "0"
         Top             =   1920
         Width           =   735
      End
      Begin VB.ComboBox cmbDB 
         Height          =   315
         ItemData        =   "frmRecordChange.frx":0968
         Left            =   120
         List            =   "frmRecordChange.frx":098A
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   4560
         Width           =   1770
      End
      Begin VB.CommandButton cmdCopyClip 
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6840
         TabIndex        =   13
         Top             =   3360
         Width           =   975
      End
      Begin VB.Frame framFile 
         Height          =   3375
         Left            =   60
         TabIndex        =   9
         Top             =   300
         Width           =   4695
         Begin VB.FileListBox filFileList 
            Height          =   3015
            Left            =   2460
            Pattern         =   "*.mdb"
            TabIndex        =   12
            ToolTipText     =   "Double Click to Open"
            Top             =   240
            Width           =   2115
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   2295
         End
         Begin VB.DirListBox Dir1 
            Height          =   2565
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   2295
         End
      End
      Begin VB.CommandButton cmdWhat 
         BackColor       =   &H0080FF80&
         Caption         =   "What is this?"
         Height          =   375
         Left            =   5100
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4500
         Width           =   1755
      End
      Begin VB.TextBox txtMap 
         Height          =   315
         Left            =   2640
         TabIndex        =   7
         Text            =   "0"
         Top             =   4560
         Width           =   375
      End
      Begin VB.CommandButton cmdListRecords 
         Caption         =   "List Records"
         Height          =   315
         Left            =   3240
         TabIndex        =   6
         Top             =   4560
         Width           =   1515
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   315
         Left            =   60
         TabIndex        =   20
         Top             =   5100
         Visible         =   0   'False
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3780
      End
      Begin VB.Line Line4 
         Index           =   3
         X1              =   9840
         X2              =   9840
         Y1              =   3780
         Y2              =   5040
      End
      Begin VB.Line Line5 
         Index           =   1
         X1              =   4860
         X2              =   9840
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Line4 
         Index           =   2
         X1              =   0
         X2              =   0
         Y1              =   3780
         Y2              =   5040
      End
      Begin VB.Line Line4 
         Index           =   1
         X1              =   4860
         X2              =   4860
         Y1              =   3780
         Y2              =   5040
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
         TabIndex        =   33
         Top             =   2460
         Width           =   900
      End
      Begin VB.Label Label8 
         Caption         =   "End"
         Height          =   195
         Left            =   5880
         TabIndex        =   32
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Start"
         Height          =   195
         Left            =   5040
         TabIndex        =   31
         Top             =   2640
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
         TabIndex        =   30
         Top             =   1560
         Width           =   1320
      End
      Begin VB.Label Label5 
         Caption         =   "End"
         Height          =   195
         Left            =   5880
         TabIndex        =   29
         Top             =   1740
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Start"
         Height          =   195
         Left            =   5040
         TabIndex        =   28
         Top             =   1740
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Step 1: Select an NMR export file."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   60
         Width           =   3615
      End
      Begin VB.Label Label10 
         Caption         =   "Step 2: Select the database type of the record numbers you wish to change."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         TabIndex        =   26
         Top             =   3900
         Width           =   4515
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   4860
         X2              =   4860
         Y1              =   0
         Y2              =   3780
      End
      Begin VB.Line Line3 
         X1              =   9840
         X2              =   9840
         Y1              =   0
         Y2              =   3780
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   0
         X2              =   4860
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Label Label1 
         Caption         =   "Step 3: Choose the record numbers you wish to change, and to what."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Index           =   0
         Left            =   5040
         TabIndex        =   25
         Top             =   120
         Width           =   2535
      End
      Begin VB.Line Line5 
         Index           =   0
         X1              =   0
         X2              =   9840
         Y1              =   3780
         Y2              =   3780
      End
      Begin VB.Label lblMap 
         AutoSize        =   -1  'True
         Caption         =   "Map:"
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
         Left            =   2100
         TabIndex        =   24
         Top             =   4620
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmRecordChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Dim DB As Database
'Dim tabActions As Recordset
Dim tabMessages As Recordset
Dim tabTextblocks As Recordset
Dim tabItems As Recordset
Dim tabClasses As Recordset
Dim tabRaces As Recordset
Dim tabSpells As Recordset
Dim tabInfo As Recordset
Dim tabMonsters As Recordset
Dim tabShops As Recordset
Dim tabRooms As Recordset

Dim nScale As Integer
Dim nScaleCount As Long

Dim nChangeMap(1) As Long
Dim bCancelProcess As Boolean
Dim DataSource As String
Dim nMapChange As Long
Dim fso As FileSystemObject
Dim ts As TextStream
Dim ChangeList() As Long
Dim nNextRange As Long


Private Sub chkMapChange_Click(Index As Integer)

If Index = 0 Then
    If chkMapChange(0).Value = 0 Then
        chkMapChange(1).Value = 0
        chkMapChange(1).Enabled = False
    Else
        chkMapChange(1).Enabled = True
    End If
End If
    
End Sub

Private Sub cmdAddAll_Click()
Dim x As Long
On Error GoTo error:

If lstRecordList.ListCount < 0 Then Exit Sub
Me.Enabled = False
Me.MousePointer = vbHourglass
bCancelProcess = False
For x = 0 To lstRecordList.ListCount - 1
    lstRecordList.ListIndex = x
    DoEvents
    Call cmdAdd_Click
    DoEvents
    If bCancelProcess Then Exit For
Next x

out:
Me.Enabled = True
Me.MousePointer = vbDefault
Exit Sub
error:
Call HandleError("cmdAddAll_Click")
Resume out:
End Sub

Private Sub cmdHideRL_Click()
fraRL.Visible = False
End Sub

Private Sub cmdListRecords_Click()
Dim nInt As Integer, tabWorkingTable As Recordset, sText As String
Dim sNumberField As String, nStartRange As Long, nEndRange As Long, nMap As Long
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

'0-Classes
'1-Items
'2-Messages
'3-Monsters
'4-Shops
'5-Spells
'6-Races
'7-Rooms
'8-Textblocks
Select Case cmbDB.ListIndex
    Case 1: 'Items
        Set tabWorkingTable = tabItems
        If tabWorkingTable.RecordCount < 1 Then GoTo no_records:
        tabWorkingTable.Index = "pkItems"
    Case 5: 'Spells
        Set tabWorkingTable = tabSpells
        If tabWorkingTable.RecordCount < 1 Then GoTo no_records:
        tabWorkingTable.Index = "pkSpells"
    Case 7, 9: 'Rooms
        Set tabWorkingTable = tabRooms
        If tabWorkingTable.RecordCount < 1 Then GoTo no_records:
        tabWorkingTable.Index = "idxRooms"
    Case 4: 'Shops
        Set tabWorkingTable = tabShops
        If tabWorkingTable.RecordCount < 1 Then GoTo no_records:
        tabWorkingTable.Index = "pkShops"
    Case 3: 'Monsters
        Set tabWorkingTable = tabMonsters
        If tabWorkingTable.RecordCount < 1 Then GoTo no_records:
        tabWorkingTable.Index = "pkMonsters"
    Case 2: 'Messages
        Set tabWorkingTable = tabMessages
        If tabWorkingTable.RecordCount < 1 Then GoTo no_records:
        tabWorkingTable.Index = "pkMessages"
    Case 8: 'Textblocks
        Set tabWorkingTable = tabTextblocks
        If tabWorkingTable.RecordCount < 1 Then GoTo no_records:
        tabWorkingTable.Index = "idxTextblocks"
    Case 0: 'Classes
        Set tabWorkingTable = tabClasses
        If tabWorkingTable.RecordCount < 1 Then GoTo no_records:
        tabWorkingTable.Index = "pkClasses"
    Case 6: 'Races
        Set tabWorkingTable = tabRaces
        If tabWorkingTable.RecordCount < 1 Then GoTo no_records:
        tabWorkingTable.Index = "pkRaces"
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

If cmbDB.ListIndex = 7 Or cmbDB.ListIndex = 9 Then  'rooms
    sNumberField = "Room Number"
Else
    sNumberField = "Number"
End If

If cmbDB.ListIndex = 7 Or cmbDB.ListIndex = 9 Then nMap = tabWorkingTable.Fields("Map Number")
nStartRange = tabWorkingTable.Fields(sNumberField)
nEndRange = nStartRange - 1

stsStatusBar.Panels(1).Text = tabWorkingTable.Name
bCancelProcess = False
Do While tabWorkingTable.EOF = False And bCancelProcess = False
    
    If nEndRange + 1 = tabWorkingTable.Fields(sNumberField) Then
        If cmbDB.ListIndex = 7 Then
            If Not nMap = tabWorkingTable.Fields("Map Number") Then GoTo no_range:
        End If
        nEndRange = tabWorkingTable.Fields(sNumberField)
        GoTo nextrec:
    End If
    
no_range:
    If nEndRange > nStartRange Then
        sText = "-" & nEndRange
    Else
        sText = ""
    End If
    
    If cmbDB.ListIndex = 7 Or cmbDB.ListIndex = 9 Then 'rooms
        lstRecordList.AddItem nMap & "/" & nStartRange & sText
    ElseIf cmbDB.ListIndex = 8 Then 'tb
        If tabWorkingTable.Fields("Part #") > 0 Then GoTo nextrec:
        lstRecordList.AddItem nStartRange & sText
    Else
        lstRecordList.AddItem nStartRange & sText
    End If
    
    If cmbDB.ListIndex = 7 Or cmbDB.ListIndex = 9 Then nMap = tabWorkingTable.Fields("Map Number")
    nStartRange = tabWorkingTable.Fields(sNumberField)
    nEndRange = nStartRange
    
nextrec:
    tabWorkingTable.MoveNext
    Call IncreaseProgressBar
    stsStatusBar.Panels(2).Text = Fix(tabWorkingTable.PercentPosition) & "%"
    If Not bUseCPU Then DoEvents
Loop

If bCancelProcess Then GoTo out:

If nEndRange > nStartRange Then
    sText = "-" & nEndRange
Else
    sText = ""
End If

If cmbDB.ListIndex = 7 Or cmbDB.ListIndex = 9 Then  'rooms
    lstRecordList.AddItem nMap & "/" & nStartRange & sText
ElseIf cmbDB.ListIndex = 8 Then 'tb
    lstRecordList.AddItem nStartRange & sText
Else
    lstRecordList.AddItem nStartRange & sText
End If

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

Private Sub cmdListRemoveLine_Click()
Dim nTemp As Long
If lstRecordList.ListIndex < 0 Then Exit Sub
nTemp = lstRecordList.ListIndex
lstRecordList.RemoveItem (nTemp)
If nTemp > 0 Then lstRecordList.ListIndex = nTemp - 1
End Sub

Private Sub cmdLog_Click()
Dim sFile As String
On Error GoTo error:

If Right(Dir1.Path, 1) = "\" Then
    sFile = Dir1.Path & "NMR-Log_RecChange_" & cmbDB.Text & ".txt"
Else
    sFile = Dir1.Path & "\NMR-Log_RecChange_" & cmbDB.Text & ".txt"
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

Private Sub cmdLogQ_Click()
MsgBox "The record changer creates a log for each database type that you change.  " _
    & "Select the database type and then click view log to view the log for that database type.", vbInformation
End Sub

Private Sub cmdQ_Click()
MsgBox "To unlock the database type dropdown, clear the change list box.", vbInformation
End Sub

Private Sub Form_Load()

If Not ReadINI("Options", "RecCHGWarn") = "1" Then
    MsgBox "WARNING: This tool was the most cumbersome and overly complex piece of NMR " _
        & vbCrLf & "that I've coded for it.  I've done testing with it, but I have a severe lack of beta " _
        & vbCrLf & "testers.  So... " & vbCrLf & "#1) BACKUP BEFORE USING THIS. " _
        & vbCrLf & "#2) CHECK THE LOGS AND VERIFY THE CHANGES. " _
        & vbCrLf & "#3) REPORT ANY PROBLEMS!!", vbExclamation
    Call WriteINI("Options", "RecCHGWarn", 1)
End If

Me.Top = ReadINI("Windows", "RecChgTop")
Me.Left = ReadINI("Windows", "RecChgLeft")

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FolderExists(ReadINI("Options", "ImportPath")) = True Then
    Dir1.Path = ReadINI("Options", "ImportPath")
Else
    Dir1.Path = App.Path
End If

Erase ChangeList()
Call AutoSizeDropDownWidth(cmbDB)
Call ExpandCombo(cmbDB, HeightOnly, TripleWidth, fraMain.hwnd)
Me.Show
cmbDB.ListIndex = 0

End Sub

Private Sub cmbDB_Click()

Select Case cmbDB.ListIndex
    Case 7: 'Rooms
        lblMap.Visible = True
        txtMap.Visible = True
        fraChangeMap.Visible = False
        cmdAddAll.Enabled = True
    Case 9:
        lblMap.Visible = False
        txtMap.Visible = False
        cmdAddAll.Enabled = False
        fraChangeMap.Visible = True
        cmdAddAll.Enabled = False
    Case Else:
        lblMap.Visible = False
        txtMap.Visible = False
        fraChangeMap.Visible = False
        cmdAddAll.Enabled = True
End Select

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
tabItems.Close
tabSpells.Close
tabRaces.Close
tabClasses.Close
tabInfo.Close
tabMonsters.Close
tabShops.Close
tabMessages.Close
tabTextblocks.Close
'tabActions.Close

DB.Close

Set tabRooms = Nothing
Set tabMonsters = Nothing
Set tabShops = Nothing
Set tabItems = Nothing
Set tabSpells = Nothing
Set tabRaces = Nothing
Set tabClasses = Nothing
Set tabInfo = Nothing
Set tabMessages = Nothing
Set tabTextblocks = Nothing
'Set tabActions = Nothing

Set DB = Nothing

End Sub

Private Function OpenTables() As Integer
On Error GoTo error:

OpenTables = -1

Set tabRooms = DB.OpenRecordset("Rooms")
Set tabItems = DB.OpenRecordset("Items")
Set tabClasses = DB.OpenRecordset("Classes")
Set tabRaces = DB.OpenRecordset("Races")
Set tabSpells = DB.OpenRecordset("Spells")
'Set tabActions = DB.OpenRecordset("Actions")
Set tabMonsters = DB.OpenRecordset("Monsters")
Set tabShops = DB.OpenRecordset("Shops")
Set tabMessages = DB.OpenRecordset("Messages")
Set tabTextblocks = DB.OpenRecordset("Textblocks")
Set tabInfo = DB.OpenRecordset("Info")

tabItems.Index = "pkItems"
tabSpells.Index = "pkSpells"
tabRooms.Index = "idxRooms"
tabShops.Index = "pkShops"
tabMonsters.Index = "pkMonsters"
tabMessages.Index = "pkMessages"
tabTextblocks.Index = "idxTextblocks"
tabClasses.Index = "pkClasses"
tabRaces.Index = "pkRaces"
        
OpenTables = 1

Exit Function
error:
Call HandleError
End Function

Private Sub cmdCopyClip_Click()
Dim sClip As String, x As Long

sClip = "DB: " & cmbDB.Text & vbCrLf
sClip = sClip & String(Len(sClip), "-") & vbCrLf
For x = 0 To lstChange.ListCount - 1
    sClip = sClip & lstChange.List(x) & vbCrLf
Next x

If Not sClip = "" Then
    Clipboard.clear
    Clipboard.SetText sClip
End If

End Sub

Private Sub cmdWhat_Click()
MsgBox "This will go through all the records in an export file and change every reference of one record number " _
    & "to another.  So if you were using record numbers 1000-2000 and needed to import something that used " _
    & "those numbers, you would use this to change them to something else.  You need to use this because references " _
    & "to those numbers will most likely be used in other databases (e.g.: monsters drop items listed by their item " _
    & "number so if you change the item numbers the monster drops will need to be changed too)." & vbCrLf & vbCrLf _
    & "This is intended to be used when importing new records that might conflict with your existing databases.  " _
    & "It could also be used to change record numbers of current records, but you will need to export basically " _
    & "your entire database collection and then delete the changed records manually (or through the database deleter).  " _
    & "If changing room numbers, items, or spells you would also need to change the references in users manually (HAH!).", vbInformation
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
If Right(Dir1.Path, 1) = "\" Then
    DataSource = Dir1.Path & DataSource
Else
    DataSource = Dir1.Path & "\" & DataSource
End If


If fso.FileExists(DataSource) = True Then
    Call ShellExecute(0&, "open", DataSource, vbNullString, vbNullString, vbNormalFocus)
Else
    MsgBox DataSource & " was not found.", vbInformation
End If

Set fso = Nothing

End Sub

Private Sub cmdAdd_Click()
On Error GoTo error:
Dim x As Long, y As Long, nTest1 As Long, nTest2 As Long

If txtMap.Visible = True And txtMap.Locked = False And Val(txtMap.Text) <= 0 Then
    MsgBox "Please enter a map number first.", vbInformation
    Exit Sub
End If

If txtToEnd.Text = "INVALID" Then Exit Sub
If txtToStart.Text < 0 Then Exit Sub
If txtFromStart.Text < 0 Then Exit Sub
'If txtToStart.Text = txtFromStart.Text Then Exit Sub

y = txtToStart.Text

For x = 0 To lstChange.ListCount - 1
    nTest1 = Val(Mid(lstChange.List(x), 1, InStr(1, lstChange.List(x), " ") - 1))
    nTest2 = Val(Mid(lstChange.List(x), InStr(1, lstChange.List(x), " -> ") + Len(" -> ")))
    If (nTest1 >= Val(txtFromStart.Text) And nTest1 <= Val(txtFromEnd.Text)) Then
        MsgBox "Adding this range would include 'change from' record numbers you've already added." _
            & vbCrLf & "(Record #" & nTest1 & " is set to be changed to " & nTest2 & ")", vbExclamation
        bCancelProcess = True
        Exit Sub
    ElseIf (nTest2 >= Val(txtFromStart.Text) And nTest2 <= Val(txtFromEnd.Text)) Then
        MsgBox "Adding this range would include 'change to' record numbers you've already added." _
            & vbCrLf & "(Record #" & nTest1 & " is set to be changed to " & nTest2 & ")", vbExclamation
        bCancelProcess = True
        Exit Sub
    End If
    
    If (nTest1 >= Val(txtToStart.Text) And nTest1 <= Val(txtToEnd.Text)) Then
        MsgBox "Adding this range would include 'change from' record numbers you've already added." _
            & vbCrLf & "(Record #" & nTest1 & " is set to be changed to " & nTest2 & ")", vbExclamation
        bCancelProcess = True
        Exit Sub
    ElseIf (nTest2 >= Val(txtToStart.Text) And nTest2 <= Val(txtToEnd.Text)) Then
        MsgBox "Adding this range would include 'change to' record numbers you've already added." _
            & vbCrLf & "(Record #" & nTest1 & " is set to be changed to " & nTest2 & ")", vbExclamation
        bCancelProcess = True
        Exit Sub
    End If
Next

For x = txtFromStart.Text To txtFromEnd.Text
    lstChange.AddItem x & " -> " & y
    y = y + 1
Next

cmbDB.Enabled = False
txtMap.Locked = True
txtMap.BackColor = &H8000000B
            
nNextRange = Val(txtToEnd.Text) + 1
txtToStart.Text = nNextRange

Call CalcRange

Exit Sub

error:
Call HandleError
Me.Enabled = True
End Sub

Private Sub cmdClear_Click()
Dim x As Integer
x = MsgBox("Are you sure?", vbYesNo + vbDefaultButton2 + vbQuestion)
If x = vbYes Then
    lstChange.clear
    cmbDB.Enabled = True
    txtMap.Locked = False
    txtMap.BackColor = &H80000005
End If
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

Private Sub cmdRemove_Click()
Dim nTemp As Long
If lstChange.ListIndex < 0 Then Exit Sub
nTemp = lstChange.ListIndex
lstChange.RemoveItem (nTemp)
If nTemp > 0 Then lstChange.ListIndex = nTemp - 1

If lstChange.ListCount = 0 Then
    cmbDB.Enabled = True
    txtMap.Locked = False
    txtMap.BackColor = &H80000005
End If
End Sub

Private Sub cmdStart_Click()
On Error GoTo error:
Dim x As Integer, sFile As String
Dim StartTime As Long, nTotalTime As Double, sTotalTime As String
Dim nYesNo As Integer, bTest As Boolean, nTempList() As Long, nNewStart As Long
Dim sClip As String, nStartRange As Long, nEndRange As Long, nChageListPos As Long

'If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

If Not cmbDB.ListIndex = 9 And lstChange.ListCount = 0 Then Exit Sub

'set the datasource
If filFileList.FileName = "" Then
    MsgBox "You must select an export file first.", vbInformation + vbOKOnly
    Exit Sub
End If

DataSource = filFileList.FileName
If Right(Dir1.Path, 1) = "\" Then
    DataSource = Dir1.Path & DataSource
    sFile = Dir1.Path & "NMR-Log_RecChange_" & cmbDB.Text & ".txt"
Else
    DataSource = Dir1.Path & "\" & DataSource
    sFile = Dir1.Path & "\NMR-Log_RecChange_" & cmbDB.Text & ".txt"
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

If cmbDB.ListIndex = 9 Then
    nChangeMap(0) = Val(txtChangeMap(0).Text)
    nChangeMap(1) = Val(txtChangeMap(1).Text)
    If nChangeMap(0) <= 0 Or nChangeMap(1) <= 0 Then
        MsgBox "Enter valid map numbers!", vbExclamation
        GoTo out:
    End If
    GoTo change_map:
End If

'build change list array
ReDim nTempList(lstChange.ListCount - 1, 1 To 2)
For x = 0 To lstChange.ListCount - 1
    nTempList(x, 1) = Val(Mid(lstChange.List(x), 1, InStr(1, lstChange.List(x), " ") - 1))
    nTempList(x, 2) = Val(Mid(lstChange.List(x), InStr(1, lstChange.List(x), "-> ") + 3))
Next

nNewStart = nTempList(0, 2)
nStartRange = nTempList(0, 1)
nEndRange = nStartRange - 1
nChageListPos = 0
ReDim ChangeList(1 To 3, nChageListPos)

For x = 0 To UBound(nTempList(), 1)
    If nEndRange + 1 = nTempList(x, 1) Then
        nEndRange = nTempList(x, 1)
        GoTo nextrec:
    End If
    
no_range:
    ReDim Preserve ChangeList(1 To 3, nChageListPos)
    ChangeList(1, nChageListPos) = nStartRange 'start range
    ChangeList(2, nChageListPos) = nEndRange 'end range
    ChangeList(3, nChageListPos) = nNewStart 'starting "change to" number
    
    nChageListPos = nChageListPos + 1
    
    nStartRange = nTempList(x, 1)
    nNewStart = nTempList(x, 2)
    nEndRange = nStartRange
nextrec:
Next x

If bCancelProcess Then GoTo out:

ReDim Preserve ChangeList(1 To 3, nChageListPos)
ChangeList(1, nChageListPos) = nStartRange 'start range
ChangeList(2, nChageListPos) = nEndRange 'end range
ChangeList(3, nChageListPos) = nNewStart 'starting "change to" number

'Debug.Print nChageListPos
Erase nTempList()
DoEvents
'UnloadForms (Me.Name)

change_map:
'start log file
If fso.FileExists(sFile) Then
    fso.DeleteFile sFile, True
End If

Set ts = fso.OpenTextFile(sFile, ForWriting, True)

ts.WriteLine ("Record Number Change Job Started " & Date & " @ " & Time)
ts.WriteBlankLines (1)

If cmbDB.ListIndex = 9 Then 'change map
    ts.WriteLine ("Changing Map " & nChangeMap(0) & " to " & nChangeMap(1))
Else
    ts.WriteLine ("ChangeList:")
    
    sClip = "DB: " & cmbDB.Text & vbCrLf
    sClip = sClip & String(Len(sClip), "-") & vbCrLf
    For x = 0 To UBound(ChangeList(), 2)
        sClip = sClip & ChangeList(1, x) & IIf(ChangeList(2, x) > ChangeList(1, x), "-" & ChangeList(2, x), "")
        sClip = sClip & " -> " & ChangeList(3, x) & _
            IIf(ChangeList(2, x) > ChangeList(1, x), "-" & (ChangeList(3, x) + (ChangeList(2, x) - ChangeList(1, x))), "") & vbCrLf
    Next x
    ts.WriteLine sClip
End If

ts.WriteBlankLines (1)

'GoTo out:
nMapChange = Val(txtMap.Text)

Set DB = OpenDatabase(DataSource)
nYesNo = OpenTables
If nYesNo < 0 Then GoTo out:

'bTest = CheckVersion
'If Not bTest = True Then GoTo out:

'set up progress bar
Call SetRange(CalcTotalRecords + lstChange.ListCount)
ProgressBar.Visible = True

bCancelProcess = False
DoEvents
'0-Classes
'1-Items
'2-Messages
'3-Monsters
'4-Shops
'5-Spells
'6-Races
'7-Rooms
'8-Textblocks
Select Case cmbDB.ListIndex
    Case 1: 'Items
        Call ScanShops
        If bCancelProcess Then GoTo out:
        Call ScanTextblocks
        If bCancelProcess Then GoTo out:
        Call ScanRooms
        If bCancelProcess Then GoTo out:
        Call ScanMonsters
        If bCancelProcess Then GoTo out:
        Call ScanSpells
        
    Case 5: 'Spells
        Call ScanItems
        If bCancelProcess Then GoTo out:
        Call ScanMonsters
        If bCancelProcess Then GoTo out:
        Call ScanRooms
        If bCancelProcess Then GoTo out:
        Call ScanSpells
        If bCancelProcess Then GoTo out:
        Call ScanTextblocks
        
    Case 7, 9: 'Rooms
        Call ScanRooms
        If bCancelProcess Then GoTo out:
        Call ScanSpells
        If bCancelProcess Then GoTo out:
        Call ScanTextblocks
        
    Case 4: 'Shops
        Call ScanRooms
        
    Case 3: 'Monsters
        Call ScanMonsters
        If bCancelProcess Then GoTo out:
        Call ScanSpells
        If bCancelProcess Then GoTo out:
        Call ScanRooms
        If bCancelProcess Then GoTo out:
        Call ScanTextblocks
        
    Case 2: 'Messages
        Call ScanItems
        If bCancelProcess Then GoTo out:
        Call ScanMonsters
        If bCancelProcess Then GoTo out:
        Call ScanRooms
        If bCancelProcess Then GoTo out:
        Call ScanSpells
        If bCancelProcess Then GoTo out:
        Call ScanTextblocks
        
    Case 8: 'Textblocks
        Call ScanClasses
        If bCancelProcess Then GoTo out:
        Call ScanItems
        If bCancelProcess Then GoTo out:
        Call ScanRooms
        If bCancelProcess Then GoTo out:
        Call ScanSpells
        If bCancelProcess Then GoTo out:
        Call ScanMonsters
        If bCancelProcess Then GoTo out:
        Call ScanTextblocks
        
    Case 0: 'Classes
        Call ScanTextblocks
        If bCancelProcess Then GoTo out:
        Call ScanItems
        If bCancelProcess Then GoTo out:
        Call ScanRooms
        If bCancelProcess Then GoTo out:
        Call ScanShops
        
    Case 6: 'Races
        Call ScanTextblocks
        If bCancelProcess Then GoTo out:
        Call ScanItems
        If bCancelProcess Then GoTo out:
        Call ScanRooms
        
End Select

If bCancelProcess Then GoTo out:

'0-Classes
'1-Items
'2-Messages
'3-Monsters
'4-Shops
'5-Spells
'6-Races
'7-Rooms
'8-Textblocks
Select Case cmbDB.ListIndex
    Case 1: 'Items
        Call ChangeItems
    Case 5: 'Spells
        Call ChangeSpells
    Case 7: 'Rooms
        Call ChangeRooms
    Case 4: 'Shops
        Call ChangeShops
    Case 3: 'Monsters
        Call ChangeMonsters
    Case 2: 'Messages
        Call ChangeMessages
    Case 8: 'Textblocks
        Call ChangeTextblocks
    Case 0: 'Classes
        Call ChangeClasses
    Case 6: 'Races
        Call ChangeRaces
    Case 9: 'Rooms
        Call ChangeMap
End Select

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
Erase nTempList()
Set ts = Nothing
Set fso = Nothing

Exit Sub
error:
Call HandleError("cmdStart_Click")
Resume out:

End Sub


Private Sub ChangeRooms()
On Error GoTo error:
Dim nStatus As Integer, x As Long, y As Long, sLine As String

'-------------------------------
'       Rooms - CHANGE
'-------------------------------
If tabRooms.RecordCount = 0 Then
    MsgBox "There are no records in the Rooms table."
    Exit Sub
End If
tabRooms.MoveFirst

ts.WriteBlankLines 1
ts.WriteLine "Changing Actual Room Record Numbers"
ts.WriteLine "-----------------------------------"
stsStatusBar.Panels(1).Text = "Rooms"

tabRooms.Index = "idxRooms"
For x = 0 To UBound(ChangeList(), 2)
    stsStatusBar.Panels(2).Text = "Changing Record Numbers (" & x & ")"
    
    For y = ChangeList(1, x) To ChangeList(2, x)
        tabRooms.Seek "=", nMapChange, y
        If tabRooms.NoMatch = True Then
            ts.WriteLine "Room " & nMapChange & "/" & y & " -- Error: Record not found in export file."
            GoTo Skip:
        End If
        sLine = "Room " & nMapChange & "/" & y & " -- Retrieve=OK"
        
        tabRooms.Seek "=", nMapChange, ChangeList(3, x) + (y - ChangeList(1, x))
        If tabRooms.NoMatch = False Then
            sLine = sLine & ", Renumber->" & ChangeList(3, x) + (y - ChangeList(1, x)) & "=Error: Record number already being used"
            ts.WriteLine sLine
            GoTo Skip:
        End If
        
        tabRooms.Seek "=", nMapChange, y
        tabRooms.Edit
    
        tabRooms.Fields("Room Number") = ChangeList(3, x) + (y - ChangeList(1, x))
        
        tabRooms.Update
        sLine = sLine & ", Renumber->" & ChangeList(3, x) + (y - ChangeList(1, x)) & "=OK"
        
        ts.WriteLine sLine
Skip:
    Next y
    
    If Not bUseCPU Then DoEvents
    Call IncreaseProgressBar
    If bCancelProcess Then Exit For
Next x

Exit Sub
error:
Call HandleError("ChangeRooms")
End Sub

Private Sub ChangeMap()
On Error GoTo error:
Dim nStatus As Integer, x As Long, y As Long, sLine As String
Dim nCurRoom(1) As Long

'-------------------------------
'       Rooms - CHANGE
'-------------------------------
If tabRooms.RecordCount = 0 Then
    MsgBox "There are no records in the Rooms table."
    Exit Sub
End If
tabRooms.Index = "idxRooms"
tabRooms.MoveFirst

ts.WriteBlankLines 1
ts.WriteLine "Changing Actual Room Record Numbers"
ts.WriteLine "-----------------------------------"
stsStatusBar.Panels(1).Text = "Rooms"

Do While tabRooms.EOF = False And bCancelProcess = False
    If tabRooms.Fields("Map Number") = nChangeMap(0) Then
        nCurRoom(0) = tabRooms.Fields("Map Number")
        nCurRoom(1) = tabRooms.Fields("Room Number")
        
        sLine = "Room " & nCurRoom(0) & "/" & nCurRoom(1) & " -- Retrieve=OK"
        
        tabRooms.Seek "=", nChangeMap(1), nCurRoom(1)
        If tabRooms.NoMatch = False Then
            sLine = sLine & ", Renumber->" & nChangeMap(1) & "/" & nCurRoom(1) & "=Error: Map/Room number already exists!"
            ts.WriteLine sLine
            tabRooms.Seek "=", nCurRoom(0), nCurRoom(1)
            GoTo Skip:
        End If
        
        tabRooms.Seek "=", nCurRoom(0), nCurRoom(1)
        tabRooms.Edit
    
        tabRooms.Fields("Map Number") = nChangeMap(1)
        
        tabRooms.Update
        sLine = sLine & ", Renumber->" & nChangeMap(1) & "/" & nCurRoom(1) & "=OK"
        
        ts.WriteLine sLine
    End If
Skip:
    If Not bUseCPU Then DoEvents
    Call IncreaseProgressBar
    If bCancelProcess Then Exit Do
    tabRooms.MoveNext
Loop

Exit Sub
error:
Call HandleError("ChangeRooms")
End Sub

Private Sub ChangeShops()
On Error GoTo error:
Dim nStatus As Integer, x As Long, y As Long, sLine As String

'-------------------------------
'       Shops - CHANGE
'-------------------------------
If tabShops.RecordCount = 0 Then
    MsgBox "There are no records in the Shops table."
    Exit Sub
End If
tabShops.MoveFirst

ts.WriteBlankLines 1
ts.WriteLine "Changing Actual Shop Record Numbers"
ts.WriteLine "-----------------------------------"
stsStatusBar.Panels(1).Text = "Shops"

tabShops.Index = "pkShops"
For x = 0 To UBound(ChangeList(), 2)
    stsStatusBar.Panels(2).Text = "Changing Record Numbers (" & x & ")"
    
    For y = ChangeList(1, x) To ChangeList(2, x)
        tabShops.Seek "=", y
        If tabShops.NoMatch = True Then
            ts.WriteLine "Shop #" & y & " -- Error: Record not found in export file."
            GoTo Skip:
        End If
        sLine = "Shop #" & y & " -- Retrieve=OK"
        
        tabShops.Seek "=", ChangeList(3, x) + (y - ChangeList(1, x))
        If tabShops.NoMatch = False Then
            sLine = sLine & ", Renumber->" & ChangeList(3, x) + (y - ChangeList(1, x)) & "=Error: Record number already being used"
            ts.WriteLine sLine
            GoTo Skip:
        End If
        
        tabShops.Seek "=", y
        tabShops.Edit
    
        tabShops.Fields("Number") = ChangeList(3, x) + (y - ChangeList(1, x))
        
        tabShops.Update
        sLine = sLine & ", Renumber->" & ChangeList(3, x) + (y - ChangeList(1, x)) & "=OK"
        
        ts.WriteLine sLine
Skip:
    Next y
    
    If Not bUseCPU Then DoEvents
    Call IncreaseProgressBar
    If bCancelProcess Then Exit For
Next x

Exit Sub
error:
Call HandleError("ChangeShops")
End Sub

Private Sub ChangeItems()
On Error GoTo error:
Dim nStatus As Integer, x As Long, y As Long, sLine As String

'-------------------------------
'       Items - CHANGE
'-------------------------------
If tabItems.RecordCount = 0 Then
    MsgBox "There are no records in the Items table."
    Exit Sub
End If
tabItems.MoveFirst

ts.WriteBlankLines 1
ts.WriteLine "Changing Actual Item Record Numbers"
ts.WriteLine "-----------------------------------"
stsStatusBar.Panels(1).Text = "Items"

tabItems.Index = "pkItems"
For x = 0 To UBound(ChangeList(), 2)
    stsStatusBar.Panels(2).Text = "Changing Record Numbers (" & x & ")"
    
    For y = ChangeList(1, x) To ChangeList(2, x)
        tabItems.Seek "=", y
        If tabItems.NoMatch = True Then
            ts.WriteLine "Item #" & y & " -- Error: Record not found in export file."
            GoTo Skip:
        End If
        sLine = "Item #" & y & " -- Retrieve=OK"
        
        tabItems.Seek "=", ChangeList(3, x) + (y - ChangeList(1, x))
        If tabItems.NoMatch = False Then
            sLine = sLine & ", Renumber->" & ChangeList(3, x) + (y - ChangeList(1, x)) & "=Error: Record number already being used"
            ts.WriteLine sLine
            GoTo Skip:
        End If
        
        tabItems.Seek "=", y
        tabItems.Edit
    
        tabItems.Fields("Number") = ChangeList(3, x) + (y - ChangeList(1, x))
        
        tabItems.Update
        sLine = sLine & ", Renumber->" & ChangeList(3, x) + (y - ChangeList(1, x)) & "=OK"
        
        ts.WriteLine sLine
Skip:
    Next y
    
    If Not bUseCPU Then DoEvents
    Call IncreaseProgressBar
    If bCancelProcess Then Exit For
Next x

Exit Sub
error:
Call HandleError("ChangeItems")
End Sub

Private Sub ChangeMonsters()
On Error GoTo error:
Dim nStatus As Integer, x As Long, y As Long, sLine As String

'-------------------------------
'       Monsters - CHANGE
'-------------------------------
If tabMonsters.RecordCount = 0 Then
    MsgBox "There are no records in the Monsters table."
    Exit Sub
End If
tabMonsters.MoveFirst

ts.WriteBlankLines 1
ts.WriteLine "Changing Actual Monster Record Numbers"
ts.WriteLine "-----------------------------------"
stsStatusBar.Panels(1).Text = "Monsters"

tabMonsters.Index = "pkMonsters"
For x = 0 To UBound(ChangeList(), 2)
    stsStatusBar.Panels(2).Text = "Changing Record Numbers (" & x & ")"
    
    For y = ChangeList(1, x) To ChangeList(2, x)
        tabMonsters.Seek "=", y
        If tabMonsters.NoMatch = True Then
            ts.WriteLine "Monster #" & y & " -- Error: Record not found in export file."
            GoTo Skip:
        End If
        sLine = "Monster #" & y & " -- Retrieve=OK"
        
        tabMonsters.Seek "=", ChangeList(3, x) + (y - ChangeList(1, x))
        If tabMonsters.NoMatch = False Then
            sLine = sLine & ", Renumber->" & ChangeList(3, x) + (y - ChangeList(1, x)) & "=Error: Record number already being used"
            ts.WriteLine sLine
            GoTo Skip:
        End If
        
        tabMonsters.Seek "=", y
        tabMonsters.Edit
    
        tabMonsters.Fields("Number") = ChangeList(3, x) + (y - ChangeList(1, x))
        
        tabMonsters.Update
        sLine = sLine & ", Renumber->" & ChangeList(3, x) + (y - ChangeList(1, x)) & "=OK"
        
        ts.WriteLine sLine
Skip:
    Next y
    
    If Not bUseCPU Then DoEvents
    Call IncreaseProgressBar
    If bCancelProcess Then Exit For
Next x

Exit Sub
error:
Call HandleError("ChangeMonsters")
End Sub

Private Sub ChangeMessages()
On Error GoTo error:
Dim nStatus As Integer, x As Long, y As Long, sLine As String

'-------------------------------
'       Messages - CHANGE
'-------------------------------
If tabMessages.RecordCount = 0 Then
    MsgBox "There are no records in the Messages table."
    Exit Sub
End If
tabMessages.MoveFirst

ts.WriteBlankLines 1
ts.WriteLine "Changing Actual Message Record Numbers"
ts.WriteLine "-----------------------------------"
stsStatusBar.Panels(1).Text = "Messages"

tabMessages.Index = "pkMessages"
For x = 0 To UBound(ChangeList(), 2)
    stsStatusBar.Panels(2).Text = "Changing Record Numbers (" & x & ")"
    
    For y = ChangeList(1, x) To ChangeList(2, x)
        tabMessages.Seek "=", y
        If tabMessages.NoMatch = True Then
            ts.WriteLine "Message #" & y & " -- Error: Record not found in export file."
            GoTo Skip:
        End If
        sLine = "Message #" & y & " -- Retrieve=OK"
        
        tabMessages.Seek "=", ChangeList(3, x) + (y - ChangeList(1, x))
        If tabMessages.NoMatch = False Then
            sLine = sLine & ", Renumber->" & ChangeList(3, x) + (y - ChangeList(1, x)) & "=Error: Record number already being used"
            ts.WriteLine sLine
            GoTo Skip:
        End If
        
        tabMessages.Seek "=", y
        tabMessages.Edit
    
        tabMessages.Fields("Number") = ChangeList(3, x) + (y - ChangeList(1, x))
        
        tabMessages.Update
        sLine = sLine & ", Renumber->" & ChangeList(3, x) + (y - ChangeList(1, x)) & "=OK"
        
        ts.WriteLine sLine
Skip:
    Next y
    
    If Not bUseCPU Then DoEvents
    Call IncreaseProgressBar
    If bCancelProcess Then Exit For
Next x

Exit Sub
error:
Call HandleError("ChangeMessages")
End Sub

Private Sub ChangeSpells()
On Error GoTo error:
Dim nStatus As Integer, x As Long, y As Long, sLine As String

'-------------------------------
'       Spells - CHANGE
'-------------------------------
If tabSpells.RecordCount = 0 Then
    MsgBox "There are no records in the Spells table."
    Exit Sub
End If
tabSpells.MoveFirst

ts.WriteBlankLines 1
ts.WriteLine "Changing Actual Spell Record Numbers"
ts.WriteLine "-----------------------------------"
stsStatusBar.Panels(1).Text = "Spells"

tabSpells.Index = "pkSpells"
For x = 0 To UBound(ChangeList(), 2)
    stsStatusBar.Panels(2).Text = "Changing Record Numbers (" & x & ")"
    
    For y = ChangeList(1, x) To ChangeList(2, x)
        tabSpells.Seek "=", y
        If tabSpells.NoMatch = True Then
            ts.WriteLine "Spell #" & y & " -- Error: Record not found in export file."
            GoTo Skip:
        End If
        sLine = "Spell #" & y & " -- Retrieve=OK"
        
        tabSpells.Seek "=", ChangeList(3, x) + (y - ChangeList(1, x))
        If tabSpells.NoMatch = False Then
            sLine = sLine & ", Renumber->" & ChangeList(3, x) + (y - ChangeList(1, x)) & "=Error: Record number already being used"
            ts.WriteLine sLine
            GoTo Skip:
        End If
        
        tabSpells.Seek "=", y
        tabSpells.Edit
    
        tabSpells.Fields("Number") = ChangeList(3, x) + (y - ChangeList(1, x))
        
        tabSpells.Update
        sLine = sLine & ", Renumber->" & ChangeList(3, x) + (y - ChangeList(1, x)) & "=OK"
        
        ts.WriteLine sLine
Skip:
    Next y
    
    If Not bUseCPU Then DoEvents
    Call IncreaseProgressBar
    If bCancelProcess Then Exit For
Next x

Exit Sub
error:
Call HandleError("ChangeSpells")
End Sub

Private Sub ChangeTextblocks()
On Error GoTo error:
Dim nStatus As Integer, x As Long, y As Long, sLine As String, nPart As Integer

'-------------------------------
'       Textblocks - CHANGE
'-------------------------------
If tabTextblocks.RecordCount = 0 Then
    MsgBox "There are no records in the Textblocks table."
    Exit Sub
End If
tabTextblocks.MoveFirst

ts.WriteBlankLines 1
ts.WriteLine "Changing Actual Textblock Record Numbers"
ts.WriteLine "-----------------------------------"
stsStatusBar.Panels(1).Text = "Textblocks"

tabTextblocks.Index = "idxTextblocks"
For x = 0 To UBound(ChangeList(), 2)
    stsStatusBar.Panels(2).Text = "Changing Record Numbers (" & x & ")"
    
    For y = ChangeList(1, x) To ChangeList(2, x)
        nPart = 0
check_next_part:
        tabTextblocks.Seek "=", y, nPart
        If tabTextblocks.NoMatch = True Then
            If nPart > 0 Then GoTo Skip:
            ts.WriteLine "Textblock #" & y & ", p" & nPart & " -- Error: Record not found in export file."
            GoTo Skip:
        End If
        sLine = "Textblock #" & y & ", p" & nPart & " -- Retrieve=OK"
        
        tabTextblocks.Seek "=", ChangeList(3, x) + (y - ChangeList(1, x)), nPart
        If tabTextblocks.NoMatch = False Then
            sLine = sLine & ", Renumber->" & ChangeList(3, x) + (y - ChangeList(1, x)) & "=Error: Record number already being used"
            ts.WriteLine sLine
            GoTo Skip:
        End If
        
        tabTextblocks.Seek "=", y, nPart
        tabTextblocks.Edit
    
        tabTextblocks.Fields("Number") = ChangeList(3, x) + (y - ChangeList(1, x))
        
        tabTextblocks.Update
        sLine = sLine & ", Renumber->" & ChangeList(3, x) + (y - ChangeList(1, x)) & "=OK"
        
        ts.WriteLine sLine
        nPart = nPart + 1
        GoTo check_next_part:
Skip:
    Next y
    
    If Not bUseCPU Then DoEvents
    Call IncreaseProgressBar
    If bCancelProcess Then Exit For
Next x

Exit Sub
error:
Call HandleError("ChangeTextblocks")
End Sub

Private Sub ChangeRaces()
On Error GoTo error:
Dim nStatus As Integer, x As Long, y As Long, sLine As String

'-------------------------------
'       Races - CHANGE
'-------------------------------
If tabRaces.RecordCount = 0 Then
    MsgBox "There are no records in the Races table."
    Exit Sub
End If
tabRaces.MoveFirst

ts.WriteBlankLines 1
ts.WriteLine "Changing Actual Race Record Numbers"
ts.WriteLine "-----------------------------------"
stsStatusBar.Panels(1).Text = "Races"

tabRaces.Index = "pkRaces"
For x = 0 To UBound(ChangeList(), 2)
    stsStatusBar.Panels(2).Text = "Changing Record Numbers (" & x & ")"
    
    For y = ChangeList(1, x) To ChangeList(2, x)
        tabRaces.Seek "=", y
        If tabRaces.NoMatch = True Then
            ts.WriteLine "Race #" & y & " -- Error: Record not found in export file."
            GoTo Skip:
        End If
        sLine = "Race #" & y & " -- Retrieve=OK"
        
        tabRaces.Seek "=", ChangeList(3, x) + (y - ChangeList(1, x))
        If tabRaces.NoMatch = False Then
            sLine = sLine & ", Renumber->" & ChangeList(3, x) + (y - ChangeList(1, x)) & "=Error: Record number already being used"
            ts.WriteLine sLine
            GoTo Skip:
        End If
        
        tabRaces.Seek "=", y
        tabRaces.Edit
    
        tabRaces.Fields("Number") = ChangeList(3, x) + (y - ChangeList(1, x))
        
        tabRaces.Update
        sLine = sLine & ", Renumber->" & ChangeList(3, x) + (y - ChangeList(1, x)) & "=OK"
        
        ts.WriteLine sLine
Skip:
    Next y
    
    If Not bUseCPU Then DoEvents
    Call IncreaseProgressBar
    If bCancelProcess Then Exit For
Next x

Exit Sub
error:
Call HandleError("ChangeRaces")
End Sub

Private Sub ChangeClasses()
On Error GoTo error:
Dim nStatus As Integer, x As Long, y As Long, sLine As String

'-------------------------------
'       Classes - CHANGE
'-------------------------------
If tabClasses.RecordCount = 0 Then
    MsgBox "There are no records in the Classes table."
    Exit Sub
End If
tabClasses.MoveFirst

ts.WriteBlankLines 1
ts.WriteLine "Changing Actual Class Record Numbers"
ts.WriteLine "-----------------------------------"
stsStatusBar.Panels(1).Text = "Classes"

tabClasses.Index = "pkClasses"
For x = 0 To UBound(ChangeList(), 2)
    stsStatusBar.Panels(2).Text = "Changing Record Numbers (" & x & ")"
    
    For y = ChangeList(1, x) To ChangeList(2, x)
        tabClasses.Seek "=", y
        If tabClasses.NoMatch = True Then
            ts.WriteLine "Class #" & y & " -- Error: Record not found in export file."
            GoTo Skip:
        End If
        sLine = "Class #" & y & " -- Retrieve=OK"
        
        tabClasses.Seek "=", ChangeList(3, x) + (y - ChangeList(1, x))
        If tabClasses.NoMatch = False Then
            sLine = sLine & ", Renumber->" & ChangeList(3, x) + (y - ChangeList(1, x)) & "=Error: Record number already being used"
            ts.WriteLine sLine
            GoTo Skip:
        End If
        
        tabClasses.Seek "=", y
        tabClasses.Edit
    
        tabClasses.Fields("Number") = ChangeList(3, x) + (y - ChangeList(1, x))
        
        tabClasses.Update
        sLine = sLine & ", Renumber->" & ChangeList(3, x) + (y - ChangeList(1, x)) & "=OK"
        
        ts.WriteLine sLine
Skip:
    Next y
    
    If Not bUseCPU Then DoEvents
    Call IncreaseProgressBar
    If bCancelProcess Then Exit For
Next x

Exit Sub
error:
Call HandleError("ChangeClasses")
End Sub

Private Sub ScanRooms()
Dim nStatus As Integer, x As Long, y As Long, nRec As Long

'-------------------------------
'       ROOMS - SCAN
'-------------------------------
If tabRooms.RecordCount = 0 Then
    ts.WriteLine vbCrLf & "Rooms -- No records to scan." & vbCrLf
    Exit Sub
End If

stsStatusBar.Panels(1).Text = "Rooms"

If cmbDB.ListIndex = 9 Then 'map change
    If chkMapChange(0).Value = 0 Then Exit Sub
End If

nRec = 0
tabRooms.MoveFirst
Do Until tabRooms.EOF Or bCancelProcess
    nRec = nRec + 1
    stsStatusBar.Panels(2).Text = nRec
    
    tabRooms.Edit
    
    Select Case cmbDB.ListIndex
            Case 1: 'Items
                
                For x = 0 To 9 'room exits
                    If tabRooms.Fields("Exit " & x) > 0 Then
                        Select Case tabRooms.Fields("Type " & x)
                            Case 0: 'normal

                            Case 2, 3, 17: 'Key, Item, Ticket
                                If tabRooms.Fields("Para1 " & x) > 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para1 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para1 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") & " -- Exit " & GetRoomExits(x, False) & " (key/item/ticket): " & tabRooms.Fields("Para1 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para1 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para1 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para1 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If

                            Case 7, 11, 12: 'Door, Gate, Remote Action
                                If tabRooms.Fields("Para4 " & x) > 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para4 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para4 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") & " -- Exit " & GetRoomExits(x, False) & " (door/gate/remote): " & tabRooms.Fields("Para4 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para4 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para4 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para4 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If
                            
                        End Select
                    End If
                Next x
                
                For x = 0 To 9 'placed items
                    If tabRooms.Fields("Placed Item " & x) > 0 Then
                        For y = 0 To UBound(ChangeList(), 2)
                            If tabRooms.Fields("Placed Item " & x) >= ChangeList(1, y) And tabRooms.Fields("Placed Item " & x) <= ChangeList(2, y) Then
                                ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") & " -- Item(placed) #" & tabRooms.Fields("Placed Item " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Placed Item " & x) - ChangeList(1, y)))
                                tabRooms.Fields("Placed Item " & x) = (ChangeList(3, y) + (tabRooms.Fields("Placed Item " & x) - ChangeList(1, y)))
                            End If
                        Next y
                    End If
                Next x

                For x = 0 To 16
                    If tabRooms.Fields("Room Item " & x) > 0 Then
                        For y = 0 To UBound(ChangeList(), 2)
                            If tabRooms.Fields("Room Item " & x) >= ChangeList(1, y) And tabRooms.Fields("Room Item " & x) <= ChangeList(2, y) Then
                                ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") & " -- Item(floor) #" & tabRooms.Fields("Room Item " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Room Item " & x) - ChangeList(1, y)))
                                tabRooms.Fields("Room Item " & x) = (ChangeList(3, y) + (tabRooms.Fields("Room Item " & x) - ChangeList(1, y)))
                            End If
                        Next y
                    End If
                Next x

                For x = 0 To 14
                    If tabRooms.Fields("Hidden Item " & x) > 0 Then
                        For y = 0 To UBound(ChangeList(), 2)
                            If tabRooms.Fields("Hidden Item " & x) >= ChangeList(1, y) And tabRooms.Fields("Hidden Item " & x) <= ChangeList(2, y) Then
                                ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") & " -- Item(invis) #" & tabRooms.Fields("Hidden Item " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Hidden Item " & x) - ChangeList(1, y)))
                                tabRooms.Fields("Hidden Item " & x) = (ChangeList(3, y) + (tabRooms.Fields("Hidden Item " & x) - ChangeList(1, y)))
                            End If
                        Next y
                    End If
                Next x
            Case 5: 'Spells
                For x = 0 To 9 'room exits
                    If Not tabRooms.Fields("Exit " & x) = 0 Then
                        Select Case tabRooms.Fields("Type " & x)
                            Case 0: 'normal
                            Case 1: 'spell
                                If Not tabRooms.Fields("Para1 " & x) = 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para1 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para1 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") & " -- Exit " & GetRoomExits(x, False) & " (spell): " & tabRooms.Fields("Para1 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para1 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para1 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para1 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If
                                
                                If Not tabRooms.Fields("Para4 " & x) = 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para4 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para4 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") & " -- Exit " & GetRoomExits(x, False) & " (spell): " & tabRooms.Fields("Para4 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para4 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para4 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para4 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If
                            
                            Case 22: 'cast
                                If Not tabRooms.Fields("Para1 " & x) = 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para1 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para1 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") & " -- Exit " & GetRoomExits(x, False) & " (cast): " & tabRooms.Fields("Para1 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para1 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para1 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para1 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If
                                
                                If Not tabRooms.Fields("Para2 " & x) = 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para2 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para2 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") & " -- Exit " & GetRoomExits(x, False) & " (cast): " & tabRooms.Fields("Para2 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para2 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para2 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para2 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If
                                
                            Case 24: 'spell trap
                                If Not tabRooms.Fields("Para1 " & x) = 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para1 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para1 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") & " -- Exit " & GetRoomExits(x, False) & " (spell trap): " & tabRooms.Fields("Para1 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para1 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para1 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para1 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If
                                
                        End Select
                    End If
                Next x
                If Not tabRooms.Fields("Spell") = 0 Then
                    For y = 0 To UBound(ChangeList(), 2)
                        If tabRooms.Fields("Spell") >= ChangeList(1, y) And tabRooms.Fields("Spell") <= ChangeList(2, y) Then
                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") & " -- RoomSpell #" & tabRooms.Fields("Spell") & " --> " & (ChangeList(3, y) + (tabRooms.Fields("Spell") - ChangeList(1, y)))
                            tabRooms.Fields("Spell") = (ChangeList(3, y) + (tabRooms.Fields("Spell") - ChangeList(1, y)))
                            Exit For
                        End If
                    Next y
                End If
                
            Case 7: 'Rooms
                If tabRooms.Fields("Map Number") = nMapChange Then 'if this map is the same map
                
                    For x = 0 To 9 'each exit
                        If tabRooms.Fields("Exit " & x) > 0 Then  'if this exit is not 0
recheck_room:
                            If tabRooms.Fields("Type " & x) <> 8 Then
                                For y = 0 To UBound(ChangeList(), 2) 'check to see if the number matches any changes
                                    If tabRooms.Fields("Exit " & x) >= ChangeList(1, y) And tabRooms.Fields("Exit " & x) <= ChangeList(2, y) Then
                                        ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") & " -- Exit " & GetRoomExits(x, False) & ": " & tabRooms.Fields("Exit " & x) & " --> " & (ChangeList(3, y) + (tabRooms.Fields("Exit " & x) - ChangeList(1, y)))
                                        tabRooms.Fields("Exit " & x) = (ChangeList(3, y) + (tabRooms.Fields("Exit " & x) - ChangeList(1, y)))
                                    End If 'end if exit = change record
                                Next y 'next change record
                            ElseIf tabRooms.Fields("Type " & x) = 8 And tabRooms.Fields("Para1 " & x) = tabRooms.Fields("Map Number") Then
                                tabRooms.Fields("Type " & x) = 0
                                GoTo recheck_room:
                            End If

                        End If 'end if exit isn't 0
                    Next x 'next exit
                    
                    If tabRooms.Fields("Death Room") > 0 Then
                        For y = 0 To UBound(ChangeList(), 2)
                            If tabRooms.Fields("Death Room") >= ChangeList(1, y) And tabRooms.Fields("Death Room") <= ChangeList(2, y) Then
                                ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") & " -- Death Room " & tabRooms.Fields("Death Room") & " --> " & (ChangeList(3, y) + (tabRooms.Fields("Death Room") - ChangeList(1, y)))
                                tabRooms.Fields("Death Room") = (ChangeList(3, y) + (tabRooms.Fields("Death Room") - ChangeList(1, y)))
                                Exit For
                            End If
                        Next y
                    End If
                    
                    If tabRooms.Fields("Exit Room") > 0 Then
                        For y = 0 To UBound(ChangeList(), 2)
                            If tabRooms.Fields("Exit Room") >= ChangeList(1, y) And tabRooms.Fields("Exit Room") <= ChangeList(2, y) Then
                                ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") & " -- Exit Room " & tabRooms.Fields("Exit Room") & " --> " & (ChangeList(3, y) + (tabRooms.Fields("Exit Room") - ChangeList(1, y)))
                                tabRooms.Fields("Exit Room") = (ChangeList(3, y) + (tabRooms.Fields("Exit Room") - ChangeList(1, y)))
                                Exit For
                            End If
                        Next y
                    End If
                    
                    If tabRooms.Fields("Control Room") > 0 Then
                        For y = 0 To UBound(ChangeList(), 2)
                            If tabRooms.Fields("Control Room") >= ChangeList(1, y) And tabRooms.Fields("Control Room") <= ChangeList(2, y) Then
                                ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") & " -- Control Room " & tabRooms.Fields("Control Room") & " --> " & (ChangeList(3, y) + (tabRooms.Fields("Control Room") - ChangeList(1, y)))
                                tabRooms.Fields("Control Room") = (ChangeList(3, y) + (tabRooms.Fields("Control Room") - ChangeList(1, y)))
                                Exit For
                            End If
                        Next y
                    End If
                    
                Else
                
                    For x = 0 To 9 'each exit
                        If tabRooms.Fields("Exit " & x) > 0 Then  'if this exit is not 0
                            If tabRooms.Fields("Type " & x) = 8 Then 'if this exit is a map change
                                If tabRooms.Fields("Para1 " & x) = nMapChange Then 'if the map matches the map we're using
                                    For y = 0 To UBound(ChangeList(), 2) 'check to see if the number matches any changes
                                        If tabRooms.Fields("Exit " & x) >= ChangeList(1, y) And tabRooms.Fields("Para1 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") & " -- Map Change to " & nMapChange & "/" & tabRooms.Fields("Exit " & x) & " --> " & (ChangeList(3, y) + (tabRooms.Fields("Para1 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Exit " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para1 " & x) - ChangeList(1, y)))
                                        End If 'end if exit = change record
                                    Next y 'next change record
                                End If 'end if map change para1 is the map we're changing
                            End If 'end if map change
                        End If 'end if exit isn't 0
                    Next x 'next exit
                    
                End If
                
                
            Case 4: 'Shops
                If Not tabRooms.Fields("Shop Number") = 0 Then
                    For y = 0 To UBound(ChangeList(), 2)
                        If tabRooms.Fields("Shop Number") >= ChangeList(1, y) And tabRooms.Fields("Shop Number") <= ChangeList(2, y) Then
                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") & " -- Shop #" & tabRooms.Fields("Shop Number") & " --> " & (ChangeList(3, y) + (tabRooms.Fields("Shop Number") - ChangeList(1, y)))
                            tabRooms.Fields("Shop Number") = (ChangeList(3, y) + (tabRooms.Fields("Shop Number") - ChangeList(1, y)))
                            Exit For
                        End If
                    Next y
                End If
            Case 3: 'Monsters
                If Not tabRooms.Fields("Perm NPC") = 0 Then
                    For y = 0 To UBound(ChangeList(), 2)
                        If tabRooms.Fields("Perm NPC") >= ChangeList(1, y) And tabRooms.Fields("Perm NPC") <= ChangeList(2, y) Then
                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") & " -- Perm NPC #" & tabRooms.Fields("Perm NPC") & " --> " & (ChangeList(3, y) + (tabRooms.Fields("Perm NPC") - ChangeList(1, y)))
                            tabRooms.Fields("Perm NPC") = (ChangeList(3, y) + (tabRooms.Fields("Perm NPC") - ChangeList(1, y)))
                            Exit For
                        End If
                    Next y
                End If
            Case 2: 'Messages
                For x = 0 To 9 'room exits
                    If tabRooms.Fields("Exit " & x) > 0 Then
                        Select Case tabRooms.Fields("Type " & x)
                            Case 0: 'normal
                            
                            'action exit type(5): para1
                            'timed exit type(16): para1
                            Case 5, 16: 'action/timed
                                If tabRooms.Fields("Para1 " & x) > 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para1 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para1 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") _
                                                & " -- Exit " & GetRoomExits(x, False) & " (action/timed): " & tabRooms.Fields("Para1 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para1 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para1 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para1 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If
                            
                            'item exit type(3): para2, para3
                            'ticket exit type(17): para2, para3
                            Case 3, 17: 'item/ticket
                                If tabRooms.Fields("Para2 " & x) > 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para2 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para2 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") _
                                                & " -- Exit " & GetRoomExits(x, False) & " (item/ticket): " & tabRooms.Fields("Para2 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para2 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para2 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para2 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If
                                If tabRooms.Fields("Para3 " & x) > 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para3 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para3 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") _
                                                & " -- Exit " & GetRoomExits(x, False) & " (item/ticket): " & tabRooms.Fields("Para3 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para3 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para3 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para3 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If
                            
                            'remote action exit type(12): para1, para3
                            Case 12: 'remote action
                                If tabRooms.Fields("Para1 " & x) > 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para1 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para1 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") _
                                                & " -- Exit " & GetRoomExits(x, False) & " (remote): " & tabRooms.Fields("Para1 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para1 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para1 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para1 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If
                                If tabRooms.Fields("Para3 " & x) > 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para3 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para3 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") _
                                                & " -- Exit " & GetRoomExits(x, False) & " (remote): " & tabRooms.Fields("Para3 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para3 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para3 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para3 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If
                            
                            'text exit type(10): para1, para2, para3
                            Case 10: 'text
                                If tabRooms.Fields("Para1 " & x) > 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para1 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para1 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") _
                                                & " -- Exit " & GetRoomExits(x, False) & " (text): " & tabRooms.Fields("Para1 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para1 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para1 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para1 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If
                                If tabRooms.Fields("Para2 " & x) > 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para2 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para2 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") _
                                                & " -- Exit " & GetRoomExits(x, False) & " (text): " & tabRooms.Fields("Para2 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para2 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para2 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para2 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If
                                If tabRooms.Fields("Para3 " & x) > 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para3 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para3 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") _
                                                & " -- Exit " & GetRoomExits(x, False) & " (text): " & tabRooms.Fields("Para3 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para3 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para3 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para3 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If
                                
                            'spell exit type(1): para3
                            'class exit type(13): para3
                            'race exit type(14): para3
                            'level exit type(15): para3
                            Case 1, 13, 14, 15: 'spell/class/race/level
                                If tabRooms.Fields("Para3 " & x) > 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para3 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para3 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") _
                                                & " -- Exit " & GetRoomExits(x, False) & " (spell/class/race/level): " & tabRooms.Fields("Para3 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para3 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para3 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para3 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If
                                
                            'hidden exit type(6): para3, para4
                            'trap exit type(9): para3, para4
                            'cast exit type(22): para3, para4
                            'spell trap exit type(24): para3, para4
                            Case 6, 9, 22, 24: 'hidden/trap/cast/spell
                                If tabRooms.Fields("Para3 " & x) > 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para3 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para3 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") _
                                                & " -- Exit " & GetRoomExits(x, False) & " (hidden/trap/cast/spell): " & tabRooms.Fields("Para3 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para3 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para3 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para3 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If
                                If tabRooms.Fields("Para4 " & x) > 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para4 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para4 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") _
                                                & " -- Exit " & GetRoomExits(x, False) & " (hidden/trap/cast/spell): " & tabRooms.Fields("Para4 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para4 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para4 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para4 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If
                                
                            'ability exit type(23): para4
                            Case 23: 'ability
                                If tabRooms.Fields("Para4 " & x) > 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para4 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para4 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") _
                                                & " -- Exit " & GetRoomExits(x, False) & " (ability): " & tabRooms.Fields("Para4 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para4 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para4 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para4 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If
                        End Select
                    End If
                Next x
                
            Case 8: 'Textblocks
            Case 0: 'Classes
                For x = 0 To 9 'room exits
                    If Not tabRooms.Fields("Exit " & x) = 0 Then
                        Select Case tabRooms.Fields("Type " & x)
                            Case 0: 'normal
                            Case 13: 'class
                                If Not tabRooms.Fields("Para1 " & x) = 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para1 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para1 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") & " -- Exit " & GetRoomExits(x, False) & " (class-ok): " & tabRooms.Fields("Para1 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para1 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para1 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para1 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If
                                
                                If Not tabRooms.Fields("Para2 " & x) = 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para2 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para2 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") & " -- Exit " & GetRoomExits(x, False) & " (class-no): " & tabRooms.Fields("Para2 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para2 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para2 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para2 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If
                        End Select
                    End If
                Next x
            Case 6: 'Races
                For x = 0 To 9 'room exits
                    If Not tabRooms.Fields("Exit " & x) = 0 Then
                        Select Case tabRooms.Fields("Type " & x)
                            Case 0: 'normal
                            Case 14: 'race
                                If Not tabRooms.Fields("Para1 " & x) = 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para1 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para1 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") & " -- Exit " & GetRoomExits(x, False) & " (race-ok): " & tabRooms.Fields("Para1 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para1 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para1 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para1 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If
                                
                                If Not tabRooms.Fields("Para2 " & x) = 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabRooms.Fields("Para2 " & x) >= ChangeList(1, y) And tabRooms.Fields("Para2 " & x) <= ChangeList(2, y) Then
                                            ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") & " -- Exit " & GetRoomExits(x, False) & " (race-no): " & tabRooms.Fields("Para2 " & x) & " to " & (ChangeList(3, y) + (tabRooms.Fields("Para2 " & x) - ChangeList(1, y)))
                                            tabRooms.Fields("Para2 " & x) = (ChangeList(3, y) + (tabRooms.Fields("Para2 " & x) - ChangeList(1, y)))
                                        End If
                                    Next y
                                End If
                        End Select
                    End If
                Next x
        Case 9: 'change map
            If tabRooms.Fields("Map Number") <> nChangeMap(0) Then 'if this room is on a different map
                
                For x = 0 To 9 'each exit
                    If tabRooms.Fields("Exit " & x) > 0 Then  'if this exit is not 0
                        If tabRooms.Fields("Type " & x) = 8 Then 'if this exit is a map change
                            If tabRooms.Fields("Para1 " & x) = nChangeMap(0) Then 'if the map matches the map we're changing
                                
                                If chkMapChange(1).Value = 1 And tabRooms.Fields("Map Number") = nChangeMap(1) Then
                                    ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") _
                                        & " -- Map Change to " & nChangeMap(0) & "/" & tabRooms.Fields("Exit " & x) _
                                        & " --> Changed to normal exit type."
                                    tabRooms.Fields("Type " & x) = 0
                                Else
                                    ts.WriteLine "Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") _
                                        & " -- Map Change to " & nChangeMap(0) & "/" & tabRooms.Fields("Exit " & x) _
                                        & " --> " & nChangeMap(1) & "/" & tabRooms.Fields("Exit " & x)
                                    tabRooms.Fields("Para1 " & x) = nChangeMap(1)
                                End If
                            End If 'end if map change para1 is the map we're changing
                        End If 'end if map change
                    End If 'end if exit isn't 0
                Next x 'next exit
                
            End If
    End Select

    tabRooms.Update
    tabRooms.MoveNext
    Call IncreaseProgressBar
    If Not bUseCPU Then DoEvents
Loop

End Sub

Private Sub ScanMonsters()
Dim nStatus As Integer, x As Long, y As Long

'-------------------------------
'       MONSTERS - SCAN
'-------------------------------

If tabMonsters.RecordCount = 0 Then
    ts.WriteLine vbCrLf & "Monsters -- No records to scan." & vbCrLf
    Exit Sub
End If

tabMonsters.MoveFirst
stsStatusBar.Panels(1).Text = "Monsters"

Do Until tabMonsters.EOF Or bCancelProcess
    stsStatusBar.Panels(2).Text = tabMonsters.Fields("Number")
    
    tabMonsters.Edit
    Select Case cmbDB.ListIndex
            Case 1: 'Items
                For x = 0 To 9
                    If tabMonsters.Fields("Item Number " & x) > 0 Then
                        For y = 0 To UBound(ChangeList(), 2)
                            If tabMonsters.Fields("Item Number " & x) >= ChangeList(1, y) And tabMonsters.Fields("Item Number " & x) <= ChangeList(2, y) Then
                                ts.WriteLine "Monster #" & tabMonsters.Fields("Number") & " -- Item(drop) #" & tabMonsters.Fields("Item Number " & x) & " to " & (ChangeList(3, y) + (tabMonsters.Fields("Item Number " & x) - ChangeList(1, y)))
                                tabMonsters.Fields("Item Number " & x) = (ChangeList(3, y) + (tabMonsters.Fields("Item Number " & x) - ChangeList(1, y)))
                            End If
                        Next y
                    End If
                Next x
                
                If tabMonsters.Fields("Weapon Number") > 0 Then
                    For y = 0 To UBound(ChangeList(), 2)
                        If tabMonsters.Fields("Weapon Number") >= ChangeList(1, y) And tabMonsters.Fields("Weapon Number") <= ChangeList(2, y) Then
                            ts.WriteLine "Monster #" & tabMonsters.Fields("Number") & " -- Item(wepn) #" & tabMonsters.Fields("Weapon Number") & " to " & (ChangeList(3, y) + (tabMonsters.Fields("Weapon Number") - ChangeList(1, y)))
                            tabMonsters.Fields("Weapon Number") = (ChangeList(3, y) + (tabMonsters.Fields("Weapon Number") - ChangeList(1, y)))
                        End If
                    Next y
                End If
                
                For x = 0 To 9
                    If tabMonsters.Fields("Ability Value " & x) > 0 Then
                        Select Case tabMonsters.Fields("Ability " & x)
                            'bad attack
                            Case 185:
                                For y = 0 To UBound(ChangeList(), 2)
                                    If tabMonsters.Fields("Ability Value " & x) >= ChangeList(1, y) And tabMonsters.Fields("Ability Value " & x) <= ChangeList(2, y) Then
                                        ts.WriteLine "Monster #" & tabMonsters.Fields("Number") & " -- " _
                                            & GetAbilityName(tabMonsters.Fields("Ability " & x)) & " #" _
                                            & tabMonsters.Fields("Ability Value " & x) & " to " & (ChangeList(3, y) + (tabMonsters.Fields("Ability Value " & x) - ChangeList(1, y)))
                                        tabMonsters.Fields("Ability Value " & x) = (ChangeList(3, y) + (tabMonsters.Fields("Ability Value " & x) - ChangeList(1, y)))
                                    End If
                                Next y
                            
                        End Select
                    End If
                Next x
            Case 5: 'Spells
                If tabMonsters.Fields("Create Spell") > 0 Then
                    For y = 0 To UBound(ChangeList(), 2)
                        If tabMonsters.Fields("Create Spell") >= ChangeList(1, y) And tabMonsters.Fields("Create Spell") <= ChangeList(2, y) Then
                            ts.WriteLine "Monster #" & tabMonsters.Fields("Number") & " -- CreateSpell #" & tabMonsters.Fields("Create Spell") & " to " & (ChangeList(3, y) + (tabMonsters.Fields("Create Spell") - ChangeList(1, y)))
                            tabMonsters.Fields("Create Spell") = (ChangeList(3, y) + (tabMonsters.Fields("Create Spell") - ChangeList(1, y)))
                        End If
                    Next y
                End If
                If tabMonsters.Fields("Death Spell") > 0 Then
                    For y = 0 To UBound(ChangeList(), 2)
                        If tabMonsters.Fields("Death Spell") >= ChangeList(1, y) And tabMonsters.Fields("Death Spell") <= ChangeList(2, y) Then
                            ts.WriteLine "Monster #" & tabMonsters.Fields("Number") & " -- DeathSpell #" & tabMonsters.Fields("Death Spell") & " to " & (ChangeList(3, y) + (tabMonsters.Fields("Death Spell") - ChangeList(1, y)))
                            tabMonsters.Fields("Death Spell") = (ChangeList(3, y) + (tabMonsters.Fields("Death Spell") - ChangeList(1, y)))
                        End If
                    Next y
                End If
                For x = 0 To 4
                    If tabMonsters.Fields("Spell Number " & x) > 0 Then 'between round
                        For y = 0 To UBound(ChangeList(), 2)
                            If tabMonsters.Fields("Spell Number " & x) >= ChangeList(1, y) And tabMonsters.Fields("Spell Number " & x) <= ChangeList(2, y) Then
                                ts.WriteLine "Monster #" & tabMonsters.Fields("Number") & " -- BetweenRoundSpell #" & tabMonsters.Fields("Spell Number " & x) & " to " & (ChangeList(3, y) + (tabMonsters.Fields("Spell Number " & x) - ChangeList(1, y)))
                                tabMonsters.Fields("Spell Number " & x) = (ChangeList(3, y) + (tabMonsters.Fields("Spell Number " & x) - ChangeList(1, y)))
                            End If
                        Next y
                    End If
                    
                    If tabMonsters.Fields("Attack Type " & x) = 2 Then 'spell attack
                        If tabMonsters.Fields("Attack Accu/Spell " & x) > 0 Then
                            For y = 0 To UBound(ChangeList(), 2)
                                If tabMonsters.Fields("Attack Accu/Spell " & x) >= ChangeList(1, y) And tabMonsters.Fields("Attack Accu/Spell " & x) <= ChangeList(2, y) Then
                                    ts.WriteLine "Monster #" & tabMonsters.Fields("Number") & " -- AttackSpell(" & x & ") #" & tabMonsters.Fields("Attack Accu/Spell " & x) & " to " & (ChangeList(3, y) + (tabMonsters.Fields("Attack Accu/Spell " & x) - ChangeList(1, y)))
                                    tabMonsters.Fields("Attack Accu/Spell " & x) = (ChangeList(3, y) + (tabMonsters.Fields("Attack Accu/Spell " & x) - ChangeList(1, y)))
                                End If
                            Next y
                        End If
                    End If
                    
                    If tabMonsters.Fields("Attack Hit Spell " & x) > 0 Then
                        For y = 0 To UBound(ChangeList(), 2)
                            If tabMonsters.Fields("Attack Hit Spell " & x) >= ChangeList(1, y) And tabMonsters.Fields("Attack Hit Spell " & x) <= ChangeList(2, y) Then
                                ts.WriteLine "Monster #" & tabMonsters.Fields("Number") & " -- AttackHitSpell #" & tabMonsters.Fields("Attack Hit Spell " & x) & " to " & (ChangeList(3, y) + (tabMonsters.Fields("Attack Hit Spell " & x) - ChangeList(1, y)))
                                tabMonsters.Fields("Attack Hit Spell " & x) = (ChangeList(3, y) + (tabMonsters.Fields("Attack Hit Spell " & x) - ChangeList(1, y)))
                            End If
                        Next y
                    End If
                Next x
            Case 7: 'Rooms
            Case 4: 'Shops
            Case 3: 'Monsters
                For x = 0 To 9
                    If tabMonsters.Fields("Ability Value " & x) > 0 Then
                        Select Case tabMonsters.Fields("Ability " & x)
                            'mons guards
                            Case 146:
                                For y = 0 To UBound(ChangeList(), 2)
                                    If tabMonsters.Fields("Ability Value " & x) >= ChangeList(1, y) And tabMonsters.Fields("Ability Value " & x) <= ChangeList(2, y) Then
                                        ts.WriteLine "Monster #" & tabMonsters.Fields("Number") & " -- " _
                                            & GetAbilityName(tabMonsters.Fields("Ability " & x)) & " #" _
                                            & tabMonsters.Fields("Ability Value " & x) & " to " & (ChangeList(3, y) + (tabMonsters.Fields("Ability Value " & x) - ChangeList(1, y)))
                                        tabMonsters.Fields("Ability Value " & x) = (ChangeList(3, y) + (tabMonsters.Fields("Ability Value " & x) - ChangeList(1, y)))
                                    End If
                                Next y
                            
                        End Select
                    End If
                Next x
            Case 2: 'Messages
                If tabMonsters.Fields("Move Msg") > 0 Then
                    For y = 0 To UBound(ChangeList(), 2)
                        If tabMonsters.Fields("Move Msg") >= ChangeList(1, y) And tabMonsters.Fields("Move Msg") <= ChangeList(2, y) Then
                            ts.WriteLine "Monster #" & tabMonsters.Fields("Number") & " -- MoveMsg #" & tabMonsters.Fields("Move Msg") & " to " & (ChangeList(3, y) + (tabMonsters.Fields("Move Msg") - ChangeList(1, y)))
                            tabMonsters.Fields("Move Msg") = (ChangeList(3, y) + (tabMonsters.Fields("Move Msg") - ChangeList(1, y)))
                        End If
                    Next y
                End If
                
                If tabMonsters.Fields("Death Msg") > 0 Then
                    For y = 0 To UBound(ChangeList(), 2)
                        If tabMonsters.Fields("Death Msg") >= ChangeList(1, y) And tabMonsters.Fields("Death Msg") <= ChangeList(2, y) Then
                            ts.WriteLine "Monster #" & tabMonsters.Fields("Number") & " -- DeathMsg #" & tabMonsters.Fields("Death Msg") & " to " & (ChangeList(3, y) + (tabMonsters.Fields("Death Msg") - ChangeList(1, y)))
                            tabMonsters.Fields("Death Msg") = (ChangeList(3, y) + (tabMonsters.Fields("Death Msg") - ChangeList(1, y)))
                        End If
                    Next y
                End If
                
                
                For x = 0 To 4
                    If tabMonsters.Fields("Attack Hit Msg " & x) > 0 Then
                        For y = 0 To UBound(ChangeList(), 2)
                            If tabMonsters.Fields("Attack Hit Msg " & x) >= ChangeList(1, y) And tabMonsters.Fields("Attack Hit Msg " & x) <= ChangeList(2, y) Then
                                ts.WriteLine "Monster #" & tabMonsters.Fields("Number") & " -- HitMsg #" & tabMonsters.Fields("Attack Hit Msg " & x) & " to " & (ChangeList(3, y) + (tabMonsters.Fields("Attack Hit Msg " & x) - ChangeList(1, y)))
                                tabMonsters.Fields("Attack Hit Msg " & x) = (ChangeList(3, y) + (tabMonsters.Fields("Attack Hit Msg " & x) - ChangeList(1, y)))
                            End If
                        Next y
                    End If
                    
                    If tabMonsters.Fields("Attack Dodge Msg " & x) > 0 Then
                        For y = 0 To UBound(ChangeList(), 2)
                            If tabMonsters.Fields("Attack Dodge Msg " & x) >= ChangeList(1, y) And tabMonsters.Fields("Attack Dodge Msg " & x) <= ChangeList(2, y) Then
                                ts.WriteLine "Monster #" & tabMonsters.Fields("Number") & " -- DodgeMsg #" & tabMonsters.Fields("Attack Dodge Msg " & x) & " to " & (ChangeList(3, y) + (tabMonsters.Fields("Attack Dodge Msg " & x) - ChangeList(1, y)))
                                tabMonsters.Fields("Attack Dodge Msg " & x) = (ChangeList(3, y) + (tabMonsters.Fields("Attack Dodge Msg " & x) - ChangeList(1, y)))
                            End If
                        Next y
                    End If
                    
                    If tabMonsters.Fields("Attack Miss Msg " & x) > 0 Then
                        For y = 0 To UBound(ChangeList(), 2)
                            If tabMonsters.Fields("Attack Miss Msg " & x) >= ChangeList(1, y) And tabMonsters.Fields("Attack Miss Msg " & x) <= ChangeList(2, y) Then
                                ts.WriteLine "Monster #" & tabMonsters.Fields("Number") & " -- MissMsg #" & tabMonsters.Fields("Attack Miss Msg " & x) & " to " & (ChangeList(3, y) + (tabMonsters.Fields("Attack Miss Msg " & x) - ChangeList(1, y)))
                                tabMonsters.Fields("Attack Miss Msg " & x) = (ChangeList(3, y) + (tabMonsters.Fields("Attack Miss Msg " & x) - ChangeList(1, y)))
                            End If
                        Next y
                    End If
                Next x
                
'                For y = 0 To UBound(ChangeList(), 2)
'                    For x = 0 To 4
'                        If tabMonsters.Fields("Attack Hit Msg " & x) >= ChangeList(1, y) and zzz <= ChangeList(2, y) Then
'                            ts.WriteLine "Monster #" & tabMonsters.Fields("Number") & " -- HitMsg #" & tabMonsters.Fields("Attack Hit Msg " & x) & " to " & (ChangeList(3, y) + (zzz - ChangeList(1, y)))
'                            tabMonsters.Fields("Attack Hit Msg " & x) = (ChangeList(3, y) + (zzz - ChangeList(1, y)))
'                        End If
'
'                        If tabMonsters.Fields("Attack Dodge Msg " & x) >= ChangeList(1, y) and zzz <= ChangeList(2, y) Then
'                            ts.WriteLine "Monster #" & tabMonsters.Fields("Number") & " -- DodgeMsg #" & tabMonsters.Fields("Attack Dodge Msg " & x) & " to " & (ChangeList(3, y) + (zzz - ChangeList(1, y)))
'                            tabMonsters.Fields("Attack Dodge Msg " & x) = (ChangeList(3, y) + (zzz - ChangeList(1, y)))
'                        End If
'
'                        If tabMonsters.Fields("Attack Miss Msg " & x) >= ChangeList(1, y) and zzz <= ChangeList(2, y) Then
'                            ts.WriteLine "Monster #" & tabMonsters.Fields("Number") & " -- MissMsg #" & tabMonsters.Fields("Attack Miss Msg " & x) & " to " & (ChangeList(3, y) + (zzz - ChangeList(1, y)))
'                            tabMonsters.Fields("Attack Miss Msg " & x) = (ChangeList(3, y) + (zzz - ChangeList(1, y)))
'                        End If
'                    Next x
'                Next y
                
            Case 8: 'Textblocks
                If tabMonsters.Fields("Greet Txt") > 0 Then
                    For y = 0 To UBound(ChangeList(), 2)
                        If tabMonsters.Fields("Greet Txt") >= ChangeList(1, y) And tabMonsters.Fields("Greet Txt") <= ChangeList(2, y) Then
                            ts.WriteLine "Monster #" & tabMonsters.Fields("Number") & " -- GreetTXT #" & tabMonsters.Fields("Greet Txt") & " to " & (ChangeList(3, y) + (tabMonsters.Fields("Greet Txt") - ChangeList(1, y)))
                            tabMonsters.Fields("Greet Txt") = (ChangeList(3, y) + (tabMonsters.Fields("Greet Txt") - ChangeList(1, y)))
                        End If
                    Next y
                End If
                If tabMonsters.Fields("Desc Txt") > 0 Then
                    For y = 0 To UBound(ChangeList(), 2)
                        If tabMonsters.Fields("Desc Txt") >= ChangeList(1, y) And tabMonsters.Fields("Desc Txt") <= ChangeList(2, y) Then
                            ts.WriteLine "Monster #" & tabMonsters.Fields("Number") & " -- DescTXT #" & tabMonsters.Fields("Desc Txt") & " to " & (ChangeList(3, y) + (tabMonsters.Fields("Desc Txt") - ChangeList(1, y)))
                            tabMonsters.Fields("Desc Txt") = (ChangeList(3, y) + (tabMonsters.Fields("Desc Txt") - ChangeList(1, y)))
                        End If
                    Next y
                End If
                If tabMonsters.Fields("Talk Txt") > 0 Then
                    For y = 0 To UBound(ChangeList(), 2)
                        If tabMonsters.Fields("Talk Txt") >= ChangeList(1, y) And tabMonsters.Fields("Talk Txt") <= ChangeList(2, y) Then
                            ts.WriteLine "Monster #" & tabMonsters.Fields("Number") & " -- TalkTXT #" & tabMonsters.Fields("Talk Txt") & " to " & (ChangeList(3, y) + (tabMonsters.Fields("Talk Txt") - ChangeList(1, y)))
                            tabMonsters.Fields("Talk Txt") = (ChangeList(3, y) + (tabMonsters.Fields("Talk Txt") - ChangeList(1, y)))
                        End If
                    Next y
                End If
                For x = 0 To 9
                    If tabMonsters.Fields("Ability Value " & x) > 0 Then
                        Select Case tabMonsters.Fields("Ability " & x)
                            'deathtext
                            Case 155:
                                For y = 0 To UBound(ChangeList(), 2)
                                    If tabMonsters.Fields("Ability Value " & x) >= ChangeList(1, y) And tabMonsters.Fields("Ability Value " & x) <= ChangeList(2, y) Then
                                        ts.WriteLine "Monster #" & tabMonsters.Fields("Number") & " -- " _
                                            & GetAbilityName(tabMonsters.Fields("Ability " & x)) & " #" _
                                            & tabMonsters.Fields("Ability Value " & x) & " to " & (ChangeList(3, y) + (tabMonsters.Fields("Ability Value " & x) - ChangeList(1, y)))
                                        tabMonsters.Fields("Ability Value " & x) = (ChangeList(3, y) + (tabMonsters.Fields("Ability Value " & x) - ChangeList(1, y)))
                                    End If
                                Next y
                            
                        End Select
                    End If
                Next x
            Case 0: 'Classes
            Case 6: 'Races
    End Select

    tabMonsters.Update
    tabMonsters.MoveNext
    Call IncreaseProgressBar
    If Not bUseCPU Then DoEvents
Loop

End Sub


Private Sub ScanShops()
Dim nStatus As Integer, x As Long, y As Long

'-------------------------------
'       SHOPS - SCAN
'-------------------------------
If tabShops.RecordCount = 0 Then
    ts.WriteLine vbCrLf & "Shops -- No records to scan." & vbCrLf
    Exit Sub
End If

stsStatusBar.Panels(1).Text = "Shops"

tabShops.MoveFirst
Do Until tabShops.EOF Or bCancelProcess
    tabShops.Edit
    stsStatusBar.Panels(2).Text = tabShops.Fields("Number")
    
    Select Case cmbDB.ListIndex
            Case 1: 'Items
                For x = 0 To 19
                    If tabShops.Fields("Item " & x) > 0 Then
                        For y = 0 To UBound(ChangeList(), 2)
                            If tabShops.Fields("Item " & x) >= ChangeList(1, y) And tabShops.Fields("Item " & x) <= ChangeList(2, y) Then
                                ts.WriteLine "Shop #" & tabShops.Fields("Number") & " -- Item #" & tabShops.Fields("Item " & x) & " to " & (ChangeList(3, y) + (tabShops.Fields("Item " & x) - ChangeList(1, y)))
                                tabShops.Fields("Item " & x) = (ChangeList(3, y) + (tabShops.Fields("Item " & x) - ChangeList(1, y)))
                            End If
                        Next y
                    End If
                Next x
            Case 5: 'Spells
            Case 7: 'Rooms
            Case 4: 'Shops
            Case 3: 'Monsters
            Case 2: 'Messages
            Case 8: 'Textblocks
            Case 0: 'Classes
                If tabShops.Fields("Class Limit") > 0 Then
                    For y = 0 To UBound(ChangeList(), 2)
                        If tabShops.Fields("Class Limit") >= ChangeList(1, y) And tabShops.Fields("Class Limit") <= ChangeList(2, y) Then
                            ts.WriteLine "Shop #" & tabShops.Fields("Number") & " -- ClassLimit #" & tabShops.Fields("Class Limit") & " to " & (ChangeList(3, y) + (tabShops.Fields("Class Limit") - ChangeList(1, y)))
                            tabShops.Fields("Class Limit") = (ChangeList(3, y) + (tabShops.Fields("Class Limit") - ChangeList(1, y)))
                        End If
                    Next y
                End If
            Case 6: 'Races
    End Select

    tabShops.Update
    tabShops.MoveNext
    Call IncreaseProgressBar
    If Not bUseCPU Then DoEvents
Loop

End Sub

Private Sub ScanClasses()
Dim nStatus As Integer, x As Long, y As Long

'-------------------------------
'       Classes - SCAN
'-------------------------------
If tabClasses.RecordCount = 0 Then
    ts.WriteLine vbCrLf & "Classes -- No records to scan." & vbCrLf
    Exit Sub
End If

stsStatusBar.Panels(1).Text = "Classes"

tabClasses.MoveFirst
Do Until tabClasses.EOF Or bCancelProcess
    tabClasses.Edit
    stsStatusBar.Panels(2).Text = tabClasses.Fields("Number")
    
    Select Case cmbDB.ListIndex
            Case 1: 'Items
            Case 5: 'Spells
            Case 7: 'Rooms
            Case 4: 'Classes
            Case 3: 'Monsters
            Case 2: 'Messages
            Case 8: 'Textblocks
                If tabClasses.Fields("Title Text") > 0 Then
                    For y = 0 To UBound(ChangeList(), 2)
                        If tabClasses.Fields("Title Text") >= ChangeList(1, y) And tabClasses.Fields("Title Text") <= ChangeList(2, y) Then
                            ts.WriteLine "Class #" & tabClasses.Fields("Number") & " -- TitleText #" & tabClasses.Fields("Title Text") & " to " & (ChangeList(3, y) + (tabClasses.Fields("Title Text") - ChangeList(1, y)))
                            tabClasses.Fields("Title Text") = (ChangeList(3, y) + (tabClasses.Fields("Title Text") - ChangeList(1, y)))
                        End If
                    Next y
                End If
            Case 0: 'Classes
            Case 6: 'Races
    End Select

    tabClasses.Update
    tabClasses.MoveNext
    Call IncreaseProgressBar
    If Not bUseCPU Then DoEvents
Loop

End Sub

Private Sub ScanItems()
Dim nStatus As Integer, x As Long, y As Long

'-------------------------------
'       ItemS - SCAN
'-------------------------------
If tabItems.RecordCount = 0 Then
    ts.WriteLine vbCrLf & "Items -- No records to scan." & vbCrLf
    Exit Sub
End If

stsStatusBar.Panels(1).Text = "Items"

tabItems.MoveFirst
Do Until tabItems.EOF Or bCancelProcess
    tabItems.Edit
    stsStatusBar.Panels(2).Text = tabItems.Fields("Number")
    
    Select Case cmbDB.ListIndex
            Case 1: 'Items
            Case 5: 'Spells
                For x = 0 To 9
                    If tabItems.Fields("Negate " & x) > 0 Then
                        For y = 0 To UBound(ChangeList(), 2)
                            If tabItems.Fields("Negate " & x) >= ChangeList(1, y) And tabItems.Fields("Negate " & x) <= ChangeList(2, y) Then
                                ts.WriteLine "Item #" & tabItems.Fields("Number") & " -- NegateSpell #" & tabItems.Fields("Negate " & x) & " to " & (ChangeList(3, y) + (tabItems.Fields("Negate " & x) - ChangeList(1, y)))
                                tabItems.Fields("Negate " & x) = (ChangeList(3, y) + (tabItems.Fields("Negate " & x) - ChangeList(1, y)))
                            End If
                        Next y
                    End If
                Next x
                For x = 0 To 19
                    If tabItems.Fields("Ability Value " & x) > 0 Then
                        Select Case tabItems.Fields("Ability " & x)
                            'learnspell,castspell,dispell,removespell,endcast,killspell,givetempspell
                            Case 42, 43, 73, 122, 151, 153, 160:
                                For y = 0 To UBound(ChangeList(), 2)
                                    If tabItems.Fields("Ability Value " & x) >= ChangeList(1, y) And tabItems.Fields("Ability Value " & x) <= ChangeList(2, y) Then
                                        ts.WriteLine "Item #" & tabItems.Fields("Number") & " -- " _
                                            & GetAbilityName(tabItems.Fields("Ability " & x)) & " #" _
                                            & tabItems.Fields("Ability Value " & x) & " to " & (ChangeList(3, y) + (tabItems.Fields("Ability Value " & x) - ChangeList(1, y)))
                                        tabItems.Fields("Ability Value " & x) = (ChangeList(3, y) + (tabItems.Fields("Ability Value " & x) - ChangeList(1, y)))
                                    End If
                                Next y
                            
                        End Select
                    End If
                Next x
            Case 7: 'Rooms
            Case 4: 'Shops
            Case 3: 'Monsters
            Case 2: 'Messages
                If tabItems.Fields("Hit Msg") > 0 Then
                    For y = 0 To UBound(ChangeList(), 2)
                        If tabItems.Fields("Hit Msg") >= ChangeList(1, y) And tabItems.Fields("Hit Msg") <= ChangeList(2, y) Then
                            ts.WriteLine "Item #" & tabItems.Fields("Number") & " -- HitMsg #" & tabItems.Fields("Hit Msg") & " to " & (ChangeList(3, y) + (tabItems.Fields("Hit Msg") - ChangeList(1, y)))
                            tabItems.Fields("Hit Msg") = (ChangeList(3, y) + (tabItems.Fields("Hit Msg") - ChangeList(1, y)))
                        End If
                    Next y
                End If
                
                If tabItems.Fields("Miss Msg") > 0 Then
                    For y = 0 To UBound(ChangeList(), 2)
                        If tabItems.Fields("Miss Msg") >= ChangeList(1, y) And tabItems.Fields("Miss Msg") <= ChangeList(2, y) Then
                            ts.WriteLine "Item #" & tabItems.Fields("Number") & " -- MissMsg #" & tabItems.Fields("Miss Msg") & " to " & (ChangeList(3, y) + (tabItems.Fields("Miss Msg") - ChangeList(1, y)))
                            tabItems.Fields("Miss Msg") = (ChangeList(3, y) + (tabItems.Fields("Miss Msg") - ChangeList(1, y)))
                        End If
                    Next y
                End If
                
                If tabItems.Fields("Distruct Msg") > 0 Then
                    For y = 0 To UBound(ChangeList(), 2)
                        If tabItems.Fields("Distruct Msg") >= ChangeList(1, y) And tabItems.Fields("Distruct Msg") <= ChangeList(2, y) Then
                            ts.WriteLine "Item #" & tabItems.Fields("Number") & " -- DistructMsg #" & tabItems.Fields("Distruct Msg") & " to " & (ChangeList(3, y) + (tabItems.Fields("Distruct Msg") - ChangeList(1, y)))
                            tabItems.Fields("Distruct Msg") = (ChangeList(3, y) + (tabItems.Fields("Distruct Msg") - ChangeList(1, y)))
                        End If
                    Next y
                End If
                
                For x = 0 To 19
                    If tabItems.Fields("Ability Value " & x) > 0 Then
                        Select Case tabItems.Fields("Ability " & x)
                            'confusemsg,descmsg,startmsg,shock,shadowform
                            Case 101, 115, 120, 137, 178:
                                For y = 0 To UBound(ChangeList(), 2)
                                    If tabItems.Fields("Ability Value " & x) >= ChangeList(1, y) And tabItems.Fields("Ability Value " & x) <= ChangeList(2, y) Then
                                        ts.WriteLine "Item #" & tabItems.Fields("Number") & " -- " _
                                            & GetAbilityName(tabItems.Fields("Ability " & x)) & " #" _
                                            & tabItems.Fields("Ability Value " & x) & " to " & (ChangeList(3, y) + (tabItems.Fields("Ability Value " & x) - ChangeList(1, y)))
                                        tabItems.Fields("Ability Value " & x) = (ChangeList(3, y) + (tabItems.Fields("Ability Value " & x) - ChangeList(1, y)))
                                    End If
                                Next y
                            
                        End Select
                    End If
                Next x
                
            Case 8: 'Textblocks
                If tabItems.Fields("Read Msg") > 0 Then
                    For y = 0 To UBound(ChangeList(), 2)
                        If tabItems.Fields("Read Msg") >= ChangeList(1, y) And tabItems.Fields("Read Msg") <= ChangeList(2, y) Then
                            ts.WriteLine "Item #" & tabItems.Fields("Number") & " -- ReadTB #" & tabItems.Fields("Read Msg") & " to " & (ChangeList(3, y) + (tabItems.Fields("Read Msg") - ChangeList(1, y)))
                            tabItems.Fields("Read Msg") = (ChangeList(3, y) + (tabItems.Fields("Read Msg") - ChangeList(1, y)))
                        End If
                    Next y
                End If
            Case 0: 'Classes
                For x = 0 To 9
                    If tabItems.Fields("Class " & x) > 0 Then
                        For y = 0 To UBound(ChangeList(), 2)
                            If tabItems.Fields("Class " & x) >= ChangeList(1, y) And tabItems.Fields("Class " & x) <= ChangeList(2, y) Then
                                ts.WriteLine "Item #" & tabItems.Fields("Number") & " -- ClassRestrict #" & tabItems.Fields("Class " & x) & " to " & (ChangeList(3, y) + (tabItems.Fields("Class " & x) - ChangeList(1, y)))
                                tabItems.Fields("Class " & x) = (ChangeList(3, y) + (tabItems.Fields("Class " & x) - ChangeList(1, y)))
                            End If
                        Next y
                    End If
                Next x
                
                For x = 0 To 19
                    If tabItems.Fields("Ability " & x) = 59 And tabItems.Fields("Ability Value " & x) > 0 Then 'class ok
                        For y = 0 To UBound(ChangeList(), 2)
                            If tabItems.Fields("Ability Value " & x) >= ChangeList(1, y) And tabItems.Fields("Ability Value " & x) <= ChangeList(2, y) Then
                                ts.WriteLine "Item #" & tabItems.Fields("Number") & " -- ClassOK #" & tabItems.Fields("Ability Value " & x) & " to " & (ChangeList(3, y) + (tabItems.Fields("Ability Value " & x) - ChangeList(1, y)))
                                tabItems.Fields("Ability Value " & x) = (ChangeList(3, y) + (tabItems.Fields("Ability Value " & x) - ChangeList(1, y)))
                            End If
                        Next y
                    End If
                Next x
            Case 6: 'Races
                For x = 0 To 9
                    If tabItems.Fields("Race " & x) > 0 Then
                        For y = 0 To UBound(ChangeList(), 2)
                            If tabItems.Fields("Race " & x) >= ChangeList(1, y) And tabItems.Fields("Race " & x) <= ChangeList(2, y) Then
                                ts.WriteLine "Item #" & tabItems.Fields("Number") & " -- RaceRestrict #" & tabItems.Fields("Race " & x) & " to " & (ChangeList(3, y) + (tabItems.Fields("Race " & x) - ChangeList(1, y)))
                                tabItems.Fields("Race " & x) = (ChangeList(3, y) + (tabItems.Fields("Race " & x) - ChangeList(1, y)))
                            End If
                        Next y
                    End If
                Next x
                
    End Select

    tabItems.Update
    tabItems.MoveNext
    Call IncreaseProgressBar
    If Not bUseCPU Then DoEvents
Loop

End Sub

Private Sub ScanSpells()
Dim nStatus As Integer, x As Long, y As Long, nYesNo As Integer
Dim nAbilDif As Long, sAbil As String, bAbilRanged As Boolean, nAbilNewVal As Long

'-------------------------------
'       SpellS - SCAN
'-------------------------------
If tabSpells.RecordCount = 0 Then
    ts.WriteLine vbCrLf & "Spells -- No records to scan." & vbCrLf
    Exit Sub
End If

stsStatusBar.Panels(1).Text = "Spells"

If cmbDB.ListIndex = 9 Then 'map change
    If chkMapChange(3).Value = 0 Then Exit Sub
End If

tabSpells.MoveFirst
Do Until tabSpells.EOF Or bCancelProcess
    tabSpells.Edit
    stsStatusBar.Panels(2).Text = tabSpells.Fields("Number")
    
    Select Case cmbDB.ListIndex
            Case 1: 'Items
                For x = 0 To 9
                    If (tabSpells.Fields("Ability " & x) = 143 Or tabSpells.Fields("Ability " & x) = 56) _
                        And tabSpells.Fields("Ability Value " & x) > 0 Then 'clearitem/rechargeitem
                        For y = 0 To UBound(ChangeList(), 2)
                            If tabSpells.Fields("Ability Value " & x) >= ChangeList(1, y) And tabSpells.Fields("Ability Value " & x) <= ChangeList(2, y) Then
                                ts.WriteLine "Spell #" & tabSpells.Fields("Number") & " -- " _
                                    & GetAbilityName(tabSpells.Fields("Ability " & x)) & " #" _
                                    & tabSpells.Fields("Ability Value " & x) & " to " & (ChangeList(3, y) + (tabSpells.Fields("Ability Value " & x) - ChangeList(1, y)))
                                tabSpells.Fields("Ability Value " & x) = (ChangeList(3, y) + (tabSpells.Fields("Ability Value " & x) - ChangeList(1, y)))
                            End If
                        Next y
                    End If
                Next x
            Case 5: 'Spells
                For x = 0 To 9
                    If tabSpells.Fields("Ability Value " & x) > 0 Then
                        Select Case tabSpells.Fields("Ability " & x)
                            'learnspell,castspell,dispell,removespell,endcast,killspell,givetempspell
                            Case 42, 43, 73, 122, 151, 153, 160:
                                For y = 0 To UBound(ChangeList(), 2)
                                    If tabSpells.Fields("Ability Value " & x) >= ChangeList(1, y) And tabSpells.Fields("Ability Value " & x) <= ChangeList(2, y) Then
                                        ts.WriteLine "Spell #" & tabSpells.Fields("Number") & " -- " _
                                            & GetAbilityName(tabSpells.Fields("Ability " & x)) & " #" _
                                            & tabSpells.Fields("Ability Value " & x) & " to " & (ChangeList(3, y) + (tabSpells.Fields("Ability Value " & x) - ChangeList(1, y)))
                                        tabSpells.Fields("Ability Value " & x) = (ChangeList(3, y) + (tabSpells.Fields("Ability Value " & x) - ChangeList(1, y)))
                                    End If
                                Next y
                            
                        End Select
                    End If
                Next x
            Case 7: 'Rooms
                nStatus = 0
                For x = 0 To 9
                    If tabSpells.Fields("Ability " & x) = 141 Then 'teleport map
                        If tabSpells.Fields("Ability Value " & x) = nMapChange Then
                            nStatus = nMapChange
                            Exit For
                        End If
                    End If
                Next x
                
                For x = 0 To 9
                    If tabSpells.Fields("Ability " & x) > 0 Then
                        If tabSpells.Fields("Ability Value " & x) = 0 Then
                            sAbil = "Min"
                            nAbilDif = tabSpells.Fields("Max") - tabSpells.Fields("Min")
                            bAbilRanged = True
                        Else
                            sAbil = "Ability Value " & x
                            nAbilDif = 0
                            bAbilRanged = False
                        End If
                        
                        Select Case tabSpells.Fields("Ability " & x)
                            'teleport room
                            Case 140:
                                If nStatus > 0 Then
                                    For y = 0 To UBound(ChangeList(), 2)
                                        If tabSpells.Fields(sAbil) >= ChangeList(1, y) And tabSpells.Fields(sAbil) <= ChangeList(2, y) Then
                                            nAbilNewVal = (ChangeList(3, y) + (tabSpells.Fields(sAbil) - ChangeList(1, y)))
                                            
                                            ts.WriteLine "Spell #" & tabSpells.Fields("Number") & " -- " _
                                                & GetAbilityName(tabSpells.Fields("Ability " & x)) & " #" _
                                                & tabSpells.Fields(sAbil) & IIf(bAbilRanged, "-" & tabSpells.Fields("Max"), "") & " to " _
                                                & nAbilNewVal & IIf(bAbilRanged, "-" & (nAbilNewVal + nAbilDif), "")
                                                
                                            tabSpells.Fields(sAbil) = nAbilNewVal
                                            If bAbilRanged Then
                                                tabSpells.Fields("Max") = nAbilNewVal + nAbilDif
                                                GoTo done_spell:
                                            End If
                                        End If
                                    Next y
                                End If
                            Case 157: 'scatter items
                                For y = 0 To UBound(ChangeList(), 2)
                                    If tabSpells.Fields(sAbil) >= ChangeList(1, y) And tabSpells.Fields(sAbil) <= ChangeList(2, y) Then
                                        nYesNo = MsgBox("Spell #" & tabSpells.Fields("Number") _
                                            & " (" & ClipNull(tabSpells.Fields("Name")) & ") contains the 'ScatterItems' ability and the room(s) falls in range with the change list.  " _
                                            & "However, it is not possible to verify that this spell is intended for the specified map.  Make the change?" _
                                            , vbYesNo + vbQuestion + vbDefaultButton2, "Change spell record?")
                                        If nYesNo = vbYes Then
                                            nAbilNewVal = (ChangeList(3, y) + (tabSpells.Fields(sAbil) - ChangeList(1, y)))
                                        
                                            ts.WriteLine "Spell #" & tabSpells.Fields("Number") & " -- " _
                                                & GetAbilityName(tabSpells.Fields("Ability " & x)) & " #" _
                                                & tabSpells.Fields(sAbil) & IIf(bAbilRanged, "-" & tabSpells.Fields("Max"), "") & " to " _
                                                & nAbilNewVal & IIf(bAbilRanged, "-" & (nAbilNewVal + nAbilDif), "")
                                                
                                            tabSpells.Fields(sAbil) = nAbilNewVal
                                            If bAbilRanged Then
                                                tabSpells.Fields("Max") = nAbilNewVal + nAbilDif
                                                GoTo done_spell:
                                            End If
                                        Else
                                            ts.WriteLine "Spell #" & tabSpells.Fields("Number") & " -- " _
                                                & GetAbilityName(tabSpells.Fields("Ability " & x)) & " -- SKIPPED."
                                        End If
                                    End If
                                Next y
                        End Select
                    End If
                Next x
                
            Case 4: 'Shops
            Case 3: 'Monsters
                For x = 0 To 9
                    If tabSpells.Fields("Ability " & x) > 0 Then
                        If tabSpells.Fields("Ability Value " & x) = 0 Then
                            sAbil = "Min"
                            nAbilDif = tabSpells.Fields("Max") - tabSpells.Fields("Min")
                            bAbilRanged = True
                        Else
                            sAbil = "Ability Value " & x
                            nAbilDif = 0
                            bAbilRanged = False
                        End If
                        
                        Select Case tabSpells.Fields("Ability " & x)
                            'summon
                            Case 12:
                                For y = 0 To UBound(ChangeList(), 2)
                                    If tabSpells.Fields(sAbil) >= ChangeList(1, y) And tabSpells.Fields(sAbil) <= ChangeList(2, y) Then
                                        nAbilNewVal = (ChangeList(3, y) + (tabSpells.Fields(sAbil) - ChangeList(1, y)))
                                        
                                        ts.WriteLine "Spell #" & tabSpells.Fields("Number") & " -- " _
                                            & GetAbilityName(tabSpells.Fields("Ability " & x)) & " #" _
                                            & tabSpells.Fields(sAbil) & IIf(bAbilRanged, "-" & tabSpells.Fields("Max"), "") & " to " _
                                            & nAbilNewVal & IIf(bAbilRanged, "-" & (nAbilNewVal + nAbilDif), "")
                                            
                                        tabSpells.Fields(sAbil) = nAbilNewVal
                                        If bAbilRanged Then
                                            tabSpells.Fields("Max") = nAbilNewVal + nAbilDif
                                            GoTo done_spell:
                                        End If
                                    End If
                                Next y
                        End Select
                    End If
                Next x
            Case 2: 'Messages
                If tabSpells.Fields("Cast MSG A") > 0 Then
                    For y = 0 To UBound(ChangeList(), 2)
                        If tabSpells.Fields("Cast MSG A") >= ChangeList(1, y) And tabSpells.Fields("Cast MSG A") <= ChangeList(2, y) Then
                            ts.WriteLine "Spell #" & tabSpells.Fields("Number") & " -- " _
                                & "CastMsgA #" & tabSpells.Fields("Cast MSG A") & " to " & (ChangeList(3, y) + (tabSpells.Fields("Cast MSG A") - ChangeList(1, y)))
                            tabSpells.Fields("Cast MSG A") = (ChangeList(3, y) + (tabSpells.Fields("Cast MSG A") - ChangeList(1, y)))
                        End If
                    Next y
                End If
                If tabSpells.Fields("Cast MSG B") > 0 Then
                    For y = 0 To UBound(ChangeList(), 2)
                        If tabSpells.Fields("Cast MSG B") >= ChangeList(1, y) And tabSpells.Fields("Cast MSG B") <= ChangeList(2, y) Then
                            ts.WriteLine "Spell #" & tabSpells.Fields("Number") & " -- " _
                                & "CastMsgB #" & tabSpells.Fields("Cast MSG B") & " to " & (ChangeList(3, y) + (tabSpells.Fields("Cast MSG B") - ChangeList(1, y)))
                            tabSpells.Fields("Cast MSG B") = (ChangeList(3, y) + (tabSpells.Fields("Cast MSG B") - ChangeList(1, y)))
                        End If
                    Next y
                End If
                For x = 0 To 9
                    If tabSpells.Fields("Ability Value " & x) > 0 Then
                        Select Case tabSpells.Fields("Ability " & x)
                            'confusemsg,descmsg,startmsg,shock,shadowform
                            Case 101, 115, 120, 137, 178:
                                For y = 0 To UBound(ChangeList(), 2)
                                    If tabSpells.Fields("Ability Value " & x) >= ChangeList(1, y) And tabSpells.Fields("Ability Value " & x) <= ChangeList(2, y) Then
                                        ts.WriteLine "Spell #" & tabSpells.Fields("Number") & " -- " _
                                            & GetAbilityName(tabSpells.Fields("Ability " & x)) & " #" _
                                            & tabSpells.Fields("Ability Value " & x) & " to " & (ChangeList(3, y) + (tabSpells.Fields("Ability Value " & x) - ChangeList(1, y)))
                                        tabSpells.Fields("Ability Value " & x) = (ChangeList(3, y) + (tabSpells.Fields("Ability Value " & x) - ChangeList(1, y)))
                                    End If
                                Next y
                            
                        End Select
                    End If
                Next x
                
            Case 8: 'Textblocks
                For x = 0 To 9
                    If tabSpells.Fields("Ability " & x) > 0 Then
                        If tabSpells.Fields("Ability Value " & x) = 0 Then
                            sAbil = "Min"
                            nAbilDif = tabSpells.Fields("Max") - tabSpells.Fields("Min")
                            bAbilRanged = True
                        Else
                            sAbil = "Ability Value " & x
                            nAbilDif = 0
                            bAbilRanged = False
                        End If
                        
                        Select Case tabSpells.Fields("Ability " & x)
                            'textblock
                            Case 148:
                                For y = 0 To UBound(ChangeList(), 2)
                                    If tabSpells.Fields(sAbil) >= ChangeList(1, y) And tabSpells.Fields(sAbil) <= ChangeList(2, y) Then
                                        nAbilNewVal = (ChangeList(3, y) + (tabSpells.Fields(sAbil) - ChangeList(1, y)))
                                        
                                        ts.WriteLine "Spell #" & tabSpells.Fields("Number") & " -- " _
                                            & GetAbilityName(tabSpells.Fields("Ability " & x)) & " #" _
                                            & tabSpells.Fields(sAbil) & IIf(bAbilRanged, "-" & tabSpells.Fields("Max"), "") & " to " _
                                            & nAbilNewVal & IIf(bAbilRanged, "-" & (nAbilNewVal + nAbilDif), "")
                                            
                                        tabSpells.Fields(sAbil) = nAbilNewVal
                                        If bAbilRanged Then
                                            tabSpells.Fields("Max") = nAbilNewVal + nAbilDif
                                            GoTo done_spell:
                                        End If
                                    End If
                                Next y
                            'deathtext
                            Case 155:
                                For y = 0 To UBound(ChangeList(), 2)
                                    If tabSpells.Fields("Ability Value " & x) >= ChangeList(1, y) And tabSpells.Fields("Ability Value " & x) <= ChangeList(2, y) Then
                                        ts.WriteLine "Spell #" & tabSpells.Fields("Number") & " -- " _
                                            & GetAbilityName(tabSpells.Fields("Ability " & x)) & " #" _
                                            & tabSpells.Fields("Ability Value " & x) & " to " & (ChangeList(3, y) + (tabSpells.Fields("Ability Value " & x) - ChangeList(1, y)))
                                        tabSpells.Fields("Ability Value " & x) = (ChangeList(3, y) + (tabSpells.Fields("Ability Value " & x) - ChangeList(1, y)))
                                    End If
                                Next y
                            
                        End Select
                    End If
                Next x
            Case 0: 'Classes
            Case 6: 'Races
        Case 9: 'map change
            For x = 0 To 9
                If tabSpells.Fields("Ability " & x) = 141 Then 'teleport map
                    If tabSpells.Fields("Ability Value " & x) = nChangeMap(0) Then
                        ts.WriteLine "Spell #" & tabSpells.Fields("Number") & " -- " _
                            & GetAbilityName(tabSpells.Fields("Ability " & x)) & " #" _
                            & tabSpells.Fields("Ability Value " & x) & " to " & nChangeMap(1)
                        tabSpells.Fields("Ability Value " & x) = nChangeMap(1)
                        Exit For
                    End If
                End If
            Next x
    End Select
done_spell:

    tabSpells.Update
    tabSpells.MoveNext
    Call IncreaseProgressBar
    If Not bUseCPU Then DoEvents
Loop


End Sub

Private Function ChangeTextblockText(ByRef sData As String, ByVal sSearchText As String, _
    ByVal sDBType As String) As Boolean
Dim x As Long, y1 As Long, y2 As Long, z As Long, sLook As String, sChar As String
Dim sTemp As String, nValue As Currency

On Error GoTo error:

If InStr(1, sData, sSearchText) = 0 Then Exit Function

x = 1
sLook = LCase(sSearchText)

check_next:
If Not bUseCPU Then DoEvents

If Not InStr(x, LCase(sData), sLook) = 0 Then
    x = InStr(x, LCase(sData), sLook) 'sets x to the position of the matched string
    
    x = x + Len(sLook)
    y1 = x
    y2 = y1
    
get_next_val:
    sChar = Mid(sData, y2, 1)
    Select Case sChar
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
            y2 = y2 + 1
            GoTo get_next_val:
        Case " ":
            If y2 = y1 Then
                y2 = y2 + 1
                GoTo get_next_val:
            End If
            
        Case Else:
            If y2 = y1 Then GoTo check_next:
    End Select
    
    sTemp = Mid(sData, y1, y2 - y1)
    nValue = Fix(Val(sTemp))
    
    If nValue > 0 Then
        For z = 0 To UBound(ChangeList(), 2)
            If nValue >= ChangeList(1, z) And nValue <= ChangeList(2, z) Then
                sData = Mid(sData, 1, y1 - 1) & (ChangeList(3, z) + (nValue - ChangeList(1, z))) _
                    & Mid(sData, y2)
                
                ts.WriteLine "Textblock #" & tabTextblocks.Fields("Number") & ", p" _
                    & tabTextblocks.Fields("Part #") & " -- " & sDBType & " #" _
                    & nValue & " to " & (ChangeList(3, z) + (nValue - ChangeList(1, z)))
                    
                ChangeTextblockText = True
            End If
        Next z
    End If
    
    x = y2
    GoTo check_next:
End If

out:
Exit Function
error:
Call HandleError("ChangeTextblockText")
Resume out:

End Function

Private Function ChangeTextblockText2ndVal(ByRef sData As String, ByVal sSearchText As String, _
    ByVal sDBType As String) As Boolean
Dim x As Long, y1 As Long, y2 As Long, z As Long, sLook As String, sChar As String
Dim sTemp As String, nValue As Currency

On Error GoTo error:

If InStr(1, sData, sSearchText) = 0 Then Exit Function

x = 1
sLook = LCase(sSearchText)

check_next:
If Not bUseCPU Then DoEvents

If Not InStr(x, LCase(sData), sLook) = 0 Then
    x = InStr(x, LCase(sData), sLook) 'sets x to the position of the matched string
    
    x = x + Len(sLook)
    y1 = x
    
move_to_next_val:
    sChar = Mid(sData, y1, 1)
    Select Case sChar
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
            y1 = y1 + 1
            GoTo move_to_next_val:
            
        Case " ":
            If y1 = x Then
                y1 = y1 + 1
                GoTo move_to_next_val:
            End If
            y1 = y1 + 1
            
        Case Else: 'no numbers after
            If y1 = x Then GoTo check_next:
            
    End Select
    
    y2 = y1
    
get_next_val:
    sChar = Mid(sData, y2, 1)
    Select Case sChar
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
            y2 = y2 + 1
            GoTo get_next_val:
        Case " ":
            If y2 = y1 Then
                y2 = y2 + 1
                GoTo get_next_val:
            End If
            
        Case Else:
            If y2 = y1 Then GoTo check_next:
    End Select
    
    sTemp = Mid(sData, y1, y2 - y1)
    nValue = Fix(Val(sTemp))
    
    If nValue > 0 Then
        If cmbDB.ListIndex = 9 Then 'map change
            If nValue = nChangeMap(0) Then
                sData = Mid(sData, 1, y1 - 1) & nChangeMap(1) _
                    & Mid(sData, y2)
                
                ts.WriteLine "Textblock #" & tabTextblocks.Fields("Number") & ", p" _
                    & tabTextblocks.Fields("Part #") & " -- " & sDBType & " #" _
                    & nValue & " to " & nChangeMap(1)
                
                ChangeTextblockText2ndVal = True
            End If
        Else
            For z = 0 To UBound(ChangeList(), 2)
                If nValue >= ChangeList(1, z) And nValue <= ChangeList(2, z) Then
                    sData = Mid(sData, 1, y1 - 1) & (ChangeList(3, z) + (nValue - ChangeList(1, z))) _
                        & Mid(sData, y2)
                    
                    ts.WriteLine "Textblock #" & tabTextblocks.Fields("Number") & ", p" _
                        & tabTextblocks.Fields("Part #") & " -- " & sDBType & " #" _
                        & nValue & " to " & (ChangeList(3, z) + (nValue - ChangeList(1, z)))
                    
                    ChangeTextblockText2ndVal = True
                End If
            Next z
        End If
    End If
    
    x = y2
    GoTo check_next:
End If

out:
Exit Function
error:
Call HandleError("ChangeTextblockText2ndVal")
Resume out:

End Function

Private Function ChangeTextblockTestSkill(ByRef sData As String) As Boolean
Dim x As Long, y1 As Long, y2 As Long, z As Long, sLook As String, sChar As String
Dim sTemp As String, nValue As Currency, sSearchText As String, sDBType As String

On Error GoTo error:

sSearchText = "testskill "
sDBType = "TestSkillText"

If InStr(1, sData, sSearchText) = 0 Then Exit Function

x = 1
sLook = LCase(sSearchText)

check_next:
If Not bUseCPU Then DoEvents

If Not InStr(x, LCase(sData), sLook) = 0 Then
    x = InStr(x, LCase(sData), sLook) 'sets x to the position of the matched string
    
    x = x + Len(sLook)
    y1 = x
    
move_to_next_val:
    sChar = Mid(sData, y1, 1)
    Select Case LCase(sChar)
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
            "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", _
            "r", "s", "t", "u", "v", "w", "x", "y", "z":
            y1 = y1 + 1
            GoTo move_to_next_val:
            
        Case " ":
            If y1 = x Then
                y1 = y1 + 1
                GoTo move_to_next_val:
            End If
            y1 = y1 + 1
            
        Case Else: 'no chars after
            If y1 = x Then GoTo check_next:
            
    End Select
    y2 = y1
    
get_next_val:
    sChar = Mid(sData, y2, 1)
    Select Case sChar
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
            y2 = y2 + 1
            GoTo get_next_val:
        Case " ":
            If y2 = y1 Then
                y2 = y2 + 1
                GoTo get_next_val:
            End If
            
        Case Else:
            If y2 = y1 Then GoTo check_next:
    End Select
    y2 = y1
    
get_3rd_val:
    sChar = Mid(sData, y2, 1)
    Select Case sChar
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
            y2 = y2 + 1
            GoTo get_3rd_val:
        Case " ":
            If y2 = y1 Then
                y2 = y2 + 1
                GoTo get_3rd_val:
            End If
            
        Case Else:
            If y2 = y1 Then GoTo check_next:
    End Select
    
    sTemp = Mid(sData, y1, y2 - y1)
    nValue = Fix(Val(sTemp))
    
    If nValue > 0 Then
        For z = 0 To UBound(ChangeList(), 2)
            If nValue >= ChangeList(1, z) And nValue <= ChangeList(2, z) Then
                sData = Mid(sData, 1, y1 - 1) & (ChangeList(3, z) + (nValue - ChangeList(1, z))) _
                    & Mid(sData, y2)
                
                ts.WriteLine "Textblock #" & tabTextblocks.Fields("Number") & ", p" _
                    & tabTextblocks.Fields("Part #") & " -- " & sDBType & " #" _
                    & nValue & " to " & (ChangeList(3, z) + (nValue - ChangeList(1, z)))
                
                ChangeTextblockTestSkill = True
            End If
        Next z
    End If
    
    x = y2
    GoTo check_next:
End If

out:
Exit Function
error:
Call HandleError("ChangeTextblockTestSkill")
Resume out:

End Function

Private Sub ScanTextblocks()
Dim nStatus As Integer, x As Long, y As Long, bChanged As Boolean, bTemp As Boolean
Dim sLook As String, sData As String, sChar As String
Dim y1 As Integer, y2 As Integer, z As Long, nItem As Long
Dim sDataTemp As String, nTB_Number As Long, nTB_Part As Long
'-------------------------------
'       TEXTBLOCKS - SCAN
'-------------------------------
If tabTextblocks.RecordCount = 0 Then
    ts.WriteLine vbCrLf & "Textblocks -- No records to scan." & vbCrLf
    Exit Sub
End If

nTB_Number = -1
nTB_Part = -1
stsStatusBar.Panels(1).Text = "Textblocks"

If cmbDB.ListIndex = 9 Then 'map change
    If chkMapChange(2).Value = 0 Then Exit Sub
End If

tabTextblocks.MoveFirst
Do Until tabTextblocks.EOF Or bCancelProcess
    stsStatusBar.Panels(2).Text = tabTextblocks.Fields("Number")
    
    sData = tabTextblocks.Fields("Data")
    
    tabTextblocks.Edit
    bChanged = False
    Select Case cmbDB.ListIndex
            Case 1: 'Items
                'covers takeitem, checkitem, failitem, roomitem, giveitem, failroomitem, clearitem
                bTemp = ChangeTextblockText(sData, "item ", "Item")
                If bTemp Then bChanged = True
                
            Case 5: 'Spells
                bTemp = ChangeTextblockText(sData, "cast ", "Cast")
                If bTemp Then bChanged = True
                bTemp = ChangeTextblockText(sData, "learnspell ", "LearnSpell")
                If bTemp Then bChanged = True
                bTemp = ChangeTextblockText(sData, "checkspell ", "CheckSpell")
                If bTemp Then bChanged = True
                
            Case 7: 'Rooms
                For z = 0 To UBound(ChangeList(), 2)
                    For y = ChangeList(1, z) To ChangeList(2, z)
                        x = 1
                        sLook = "teleport " & y & " " & nMapChange 'teleport <room>     '...<map>
room_check_next:
                        If Not InStr(x, sData, sLook) = 0 Then
                            x = InStr(x, sData, sLook) 'sets x to the position of the matched string
                            
                            y1 = x + 9 'len of "teleport " (to position y1 at first number)
                            y2 = y1 + Len(sLook) - 9 'positions y2 after the last number
                            
                            sChar = Mid(sData, y2, 1) 'check to make sure this match isn't part of another number
                            Select Case sChar
                                Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
                                    x = y1
                                    GoTo room_check_next:
                                Case Else:
                            End Select
                            
                            sData = Mid(sData, 1, y1 - 1) & (ChangeList(3, z) + (y - ChangeList(1, z))) _
                                & " " & nMapChange & Mid(sData, y2)
                            ts.WriteLine "Textblock #" & tabTextblocks.Fields("Number") & ", p" _
                                & tabTextblocks.Fields("Part #") & " -- Teleport " & y & " " & nMapChange & " --> " _
                                & (ChangeList(3, z) + (y - ChangeList(1, z))) & " " & nMapChange
                            bChanged = True
                            
                            x = y1 + Len(CStr((ChangeList(3, z) + (y - ChangeList(1, z)))))
                            GoTo room_check_next:
                        End If
                    Next y
                Next z
                
                'remoteaction [room]
                bTemp = ChangeTextblockText(sData, "remoteaction ", "RemoteActionOnRoom")
                If bTemp Then bChanged = True
                
            Case 4: 'Shops
            Case 3: 'Monsters
                'summon <monster>
                bTemp = ChangeTextblockText(sData, "summon ", "SummonMonster")
                If bTemp Then bChanged = True
                
                'needmonster <monster>
                bTemp = ChangeTextblockText(sData, "needmonster ", "NeedMonster")
                If bTemp Then bChanged = True
                
            Case 2: 'Messages
                'message <msg>
                bTemp = ChangeTextblockText(sData, "message ", "Msg")
                If bTemp Then bChanged = True
                
                'nomonsters <msg>
                bTemp = ChangeTextblockText(sData, "nomonsters ", "NoMonsMsg")
                If bTemp Then bChanged = True
                
                'checkitem <item> <msg>
                'failitem <item> <msg>
                'failroomitem <item> <msg>
                'takeitem <item> <msg>
                'roomitem <item> <msg>
                bTemp = ChangeTextblockText2ndVal(sData, "checkitem ", "CheckItemMsg")
                If bTemp Then bChanged = True
                bTemp = ChangeTextblockText2ndVal(sData, "failitem ", "FailItemMsg")
                If bTemp Then bChanged = True
                bTemp = ChangeTextblockText2ndVal(sData, "failroomitem ", "FailRoomItemMsg")
                If bTemp Then bChanged = True
                bTemp = ChangeTextblockText2ndVal(sData, "takeitem ", "TakeItemMsg")
                If bTemp Then bChanged = True
                bTemp = ChangeTextblockText2ndVal(sData, "roomitem ", "RoomItemMsg")
                If bTemp Then bChanged = True
                
                'maxlevel <value> <msg>
                'minlevel <value> <msg>
                bTemp = ChangeTextblockText2ndVal(sData, "maxlevel ", "MaxLvlMsg")
                If bTemp Then bChanged = True
                bTemp = ChangeTextblockText2ndVal(sData, "minlevel ", "MinLvlMsg")
                If bTemp Then bChanged = True
                
                'needmonster <monster> <msg>
                bTemp = ChangeTextblockText2ndVal(sData, "needmonster ", "NeedMonMsg")
                If bTemp Then bChanged = True
                
                'price <value> <msg>
                bTemp = ChangeTextblockText2ndVal(sData, "price ", "PriceMsg")
                If bTemp Then bChanged = True
                
                'evilaligned <value> <msg>
                'goodaligned <value> <msg>
                bTemp = ChangeTextblockText2ndVal(sData, "evilaligned ", "EvilAlignMsg")
                If bTemp Then bChanged = True
                bTemp = ChangeTextblockText2ndVal(sData, "goodaligned ", "GoodAlignMsg")
                If bTemp Then bChanged = True
                
                'failability <value> <msg>
                bTemp = ChangeTextblockText2ndVal(sData, "failability ", "FailAbilMsg")
                If bTemp Then bChanged = True
                
                'remoteaction [room] [message] [????] [exit]
                bTemp = ChangeTextblockText2ndVal(sData, "remoteaction ", "RemoteActionMsg")
                If bTemp Then bChanged = True
                
            Case 8: 'Textblocks
                bTemp = ChangeTextblockText(sData, "text ", "Text")
                If bTemp Then bChanged = True
                bTemp = ChangeTextblockText(sData, "random ", "RandomText")
                If bTemp Then bChanged = True
                bTemp = ChangeTextblockText(sData, ":", "GreetCommand")
                If bTemp Then bChanged = True
                bTemp = ChangeTextblockText2ndVal(sData, "checkspell ", "CheckSpellText")
                If bTemp Then bChanged = True
                bTemp = ChangeTextblockTestSkill(sData)
                If bTemp Then bChanged = True
                
                If tabTextblocks.Fields("Link To") > 0 Then
                    For y = 0 To UBound(ChangeList(), 2)
                        If tabTextblocks.Fields("Link To") >= ChangeList(1, y) And tabTextblocks.Fields("Link To") <= ChangeList(2, y) Then
                            ts.WriteLine "Textblock #" & tabTextblocks.Fields("Number") & " -- LinkTo #" & tabTextblocks.Fields("Link To") & " to " & (ChangeList(3, y) + (tabTextblocks.Fields("Link To") - ChangeList(1, y)))
                            tabTextblocks.Fields("Link To") = (ChangeList(3, y) + (tabTextblocks.Fields("Link To") - ChangeList(1, y)))
                            bChanged = True
                        End If
                    Next y
                End If
                
            Case 0: 'Classes
                bTemp = ChangeTextblockText(sData, "class ", "Class")
                If bTemp Then bChanged = True
                
            Case 6: 'Races
                bTemp = ChangeTextblockText(sData, "race ", "Race")
                If bTemp Then bChanged = True
            
            Case 9: 'map change
                bTemp = ChangeTextblockText2ndVal(sData, "teleport ", "Teleport")

    End Select
    
    If bChanged Then
        If Len(sData) > 2000 Then
            nTB_Number = tabTextblocks.Fields("Number")
            nTB_Part = tabTextblocks.Fields("Part #")
        End If
check_length:
        If Len(sData) > 2000 Then
            x = InStr(1, sData, Chr(10))
            If x > 2000 Or x < 1 Then
                x = InStr(1, sData, ":")
                If x > 0 Then
                    y = x
                    Do Until x < 1 Or x > 2000
                        x = InStr(y + 1, sData, ":")
                        If x > 0 And x < 2000 Then y = x
                    Loop
                    If y > 2000 Then y = 2000
                    sDataTemp = Right(sData, Len(sData) - y)
                    sData = Left(sData, y)
                Else
                    sDataTemp = Right(sData, Len(sData) - 2000)
                    sData = Left(sData, 2000)
                End If
            Else
                y = x
                Do Until x < 1 Or x > 2000
                    x = InStr(y + 1, sData, Chr(10))
                    If x > 0 And x < 2000 Then y = x
                Loop
                If y > 2000 Then y = 2000
                sDataTemp = Right(sData, Len(sData) - y)
                sData = Left(sData, y)
            End If
            
            tabTextblocks.Fields("Data") = sData
            tabTextblocks.Update
            
            x = tabTextblocks.Fields("Number")
            y = tabTextblocks.Fields("Part #") + 1

            tabTextblocks.Index = "idxTextblocks"
            tabTextblocks.Seek "=", x, y
            If tabTextblocks.NoMatch Then
                tabTextblocks.AddNew
                tabTextblocks.Fields("Number") = x
                tabTextblocks.Fields("Part #") = y
                tabTextblocks.Fields("Link To") = 0
                tabTextblocks.Fields("Data") = Chr(0)
                tabTextblocks.Update
                
                tabTextblocks.Seek "=", x, y
                tabTextblocks.Edit
                
                sData = sDataTemp
                ts.WriteLine "Textblock #" & x & ", p" & y & " added because of new length."
            Else
                ts.WriteLine "Textblock #" & x & ", p" & y & " appended because of new length."
                'Debug.Print sDataTemp
                sData = sDataTemp & tabTextblocks.Fields("Data")
                tabTextblocks.Edit
            End If
            
            If bUseCPU Then DoEvents
            GoTo check_length:
        Else
            tabTextblocks.Fields("Data") = sData
            tabTextblocks.Update
            
            If nTB_Number >= 0 Then
                tabTextblocks.Index = "idxTextblocks"
                tabTextblocks.Seek "=", nTB_Number, nTB_Part
                
                nTB_Number = -1
                nTB_Part = -1
            End If
        End If
        
    End If
    
    tabTextblocks.MoveNext
    Call IncreaseProgressBar
    If Not bUseCPU Then DoEvents
Loop

End Sub



Private Sub ScanUsers()
'Dim nStatus As Integer, x As Long, y As Long, bChanged As Boolean, nRec As Long
'
''-------------------------------
''       USERS - SCAN
''-------------------------------
'nStatus = BTRCALL(BGETFIRST, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0)
'If Not nStatus = 0 Then
'    MsgBox "Couldn't get first User record."
'    Exit Sub
'End If
'
'stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_USERS
'
'nRec = 1
'Do While nStatus = 0 And bCancelProcess = False
'    UserRowToStruct Userdatabuf.buf
'    stsStatusBar.Panels(2).Text = "Searching Users (" & nRec & ")"
'
'    bChanged = False
'    Select Case cmbDB.ListIndex
'            Case 1: 'Items
'
'                For x = 0 To 19
'                    If Not Userrec.WornItem(x) = 0 Then
'                        For y = 0 To UBound(ChangeList(), 2)
'                            If Userrec.WornItem(x) >= ChangeList(1, y) And zzz <= ChangeList(2, y) Then
'                                ts.WriteLine "User " & ClipNull(Userrec.FirstName) & "/" & ClipNull(Userrec.BBSName) & " -- Item(worn) #" & Userrec.WornItem(x) & " to " & (ChangeList(3, y) + (zzz - ChangeList(1, y)))
'                                Userrec.WornItem(x) = (ChangeList(3, y) + (zzz - ChangeList(1, y)))
'                                bChanged = True
'                            End If
'                        Next y
'                    End If
'
'                    If Not Userrec.Key(x) = 0 Then
'                        For y = 0 To UBound(ChangeList(), 2)
'                            If Userrec.Key(x) >= ChangeList(1, y) And zzz <= ChangeList(2, y) Then
'                                ts.WriteLine "User " & ClipNull(Userrec.FirstName) & "/" & ClipNull(Userrec.BBSName) & " -- Item(key) #" & Userrec.Key(x) & " to " & (ChangeList(3, y) + (zzz - ChangeList(1, y)))
'                                Userrec.Key(x) = (ChangeList(3, y) + (zzz - ChangeList(1, y)))
'                                bChanged = True
'                            End If
'                        Next y
'                    End If
'
'                    If Not Userrec.Item(x) = 0 Then
'                        For y = 0 To UBound(ChangeList(), 2)
'                            If Userrec.Item(x) >= ChangeList(1, y) And zzz <= ChangeList(2, y) Then
'                                ts.WriteLine "User " & ClipNull(Userrec.FirstName) & "/" & ClipNull(Userrec.BBSName) & " -- Item(inven) #" & Userrec.Item(x) & " to " & (ChangeList(3, y) + (zzz - ChangeList(1, y)))
'                                Userrec.Item(x) = (ChangeList(3, y) + (zzz - ChangeList(1, y)))
'                                bChanged = True
'                            End If
'                        Next y
'                    End If
'                Next x
'
'                For x = 20 To 49
'                    If Not Userrec.Key(x) = 0 Then
'                        For y = 0 To UBound(ChangeList(), 2)
'                            If Userrec.Key(x) >= ChangeList(1, y) And zzz <= ChangeList(2, y) Then
'                                ts.WriteLine "User " & ClipNull(Userrec.FirstName) & "/" & ClipNull(Userrec.BBSName) & " -- Item(key) #" & Userrec.Key(x) & " to " & (ChangeList(3, y) + (zzz - ChangeList(1, y)))
'                                Userrec.Key(x) = (ChangeList(3, y) + (zzz - ChangeList(1, y)))
'                                bChanged = True
'                            End If
'                        Next y
'                    End If
'
'                    If Not Userrec.Item(x) = 0 Then
'                        For y = 0 To UBound(ChangeList(), 2)
'                            If Userrec.Item(x) >= ChangeList(1, y) And zzz <= ChangeList(2, y) Then
'                                ts.WriteLine "User " & ClipNull(Userrec.FirstName) & "/" & ClipNull(Userrec.BBSName) & " -- Item(inven) #" & Userrec.Item(x) & " to " & (ChangeList(3, y) + (zzz - ChangeList(1, y)))
'                                Userrec.Item(x) = (ChangeList(3, y) + (zzz - ChangeList(1, y)))
'                                bChanged = True
'                            End If
'                        Next y
'                    End If
'                Next x
'
'                For x = 50 To 99
'                    If Not Userrec.Item(x) = 0 Then
'                        For y = 0 To UBound(ChangeList(), 2)
'                            If Userrec.Item(x) >= ChangeList(1, y) And zzz <= ChangeList(2, y) Then
'                                ts.WriteLine "User " & ClipNull(Userrec.FirstName) & "/" & ClipNull(Userrec.BBSName) & " -- Item(inven) #" & Userrec.Item(x) & " to " & (ChangeList(3, y) + (zzz - ChangeList(1, y)))
'                                Userrec.Item(x) = (ChangeList(3, y) + (zzz - ChangeList(1, y)))
'                                bChanged = True
'                            End If
'                        Next y
'                    End If
'                Next x
'
'                For y = 0 To UBound(ChangeList(), 2)
'                    If Userrec.WeaponHand >= ChangeList(1, y) And zzz <= ChangeList(2, y) Then
'                        ts.WriteLine "User " & ClipNull(Userrec.FirstName) & "/" & ClipNull(Userrec.BBSName) & " -- Item(wepn) #" & Userrec.WeaponHand & " to " & (ChangeList(3, y) + (zzz - ChangeList(1, y)))
'                        Userrec.WeaponHand = (ChangeList(3, y) + (zzz - ChangeList(1, y)))
'                        bChanged = True
'                    End If
'                Next y
'
'            Case 5: 'Spells
'            Case 7: 'Rooms
'                If Userrec.MapNumber = nMapChange Then
'                    For y = 0 To UBound(ChangeList(), 2)
'                        If Userrec.RoomNum >= ChangeList(1, y) And zzz <= ChangeList(2, y) Then
'                            ts.WriteLine "User " & ClipNull(Userrec.FirstName) & "/" & ClipNull(Userrec.BBSName) & " -- Current Room " & nMapChange & "/" & Userrec.RoomNum & " to " & (ChangeList(3, y) + (zzz - ChangeList(1, y)))
'                            Userrec.RoomNum = (ChangeList(3, y) + (zzz - ChangeList(1, y)))
'                            bChanged = True
'                        End If
'                    Next y
'                End If
'
'            Case 4: 'Shops
'            Case 3: 'Monsters
'            Case 2: 'Messages
'            Case 8: 'Textblocks
'            Case 0: 'Classes
'            Case 6: 'Races
'    End Select
'
'    If bChanged Then
'        nStatus = UpdateUser
'    End If
'    nStatus = BTRCALL(BGETNEXT, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0)
'    If Not bUseCPU Then DoEvents
'    nRec = nRec + 1
'    Call IncreaseProgressBar
'Loop

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
    Call WriteINI("Windows", "RecChgTop", Me.Top)
    Call WriteINI("Windows", "RecChgLeft", Me.Left)
End If

Set fso = Nothing
Set ts = Nothing
End Sub

Private Sub lstRecordList_Click()
Dim sLine As String
On Error GoTo error:

If lstRecordList.ListCount < 1 Then Exit Sub
If lstRecordList.ListIndex < 0 Then Exit Sub

sLine = lstRecordList.List(lstRecordList.ListIndex)
If txtMap.Visible And Not txtMap.Locked Then
    txtMap.Text = Left(sLine, InStr(1, sLine, "/") - 1)
End If
If InStr(1, sLine, "/") > 0 Then
    sLine = Mid(sLine, InStr(1, sLine, "/") + 1)
End If

If InStr(1, sLine, "-") > 0 Then
    txtFromStart.Text = Val(Left(sLine, Mid(sLine, 1, InStr(1, sLine, "-") - 1)))
    txtFromEnd.Text = Val(Mid(sLine, InStr(1, sLine, "-") + 1))
Else
    txtFromStart.Text = Val(sLine)
    txtFromEnd.Text = Val(sLine)
End If

Call CalcRange

out:
Exit Sub
error:
Call HandleError("lstRecordList_Click")
Resume out:
End Sub

Private Sub txtFromEnd_GotFocus()
Call SelectAll(txtFromEnd)

End Sub

Private Sub txtFromEnd_KeyUp(KeyCode As Integer, Shift As Integer)
Call CalcRange
End Sub

Private Sub txtFromStart_GotFocus()
Call SelectAll(txtFromStart)

End Sub

Private Sub txtFromStart_KeyUp(KeyCode As Integer, Shift As Integer)
Call CalcRange
End Sub
Private Sub CalcRange()
txtToEnd.Text = Val(txtToStart.Text) + (Val(txtFromEnd.Text) - Val(txtFromStart.Text))
If Val(txtToEnd.Text) < Val(txtToStart.Text) Then txtToEnd.Text = "INVALID"
End Sub

Private Sub txtMap_GotFocus()
Call SelectAll(txtMap)

End Sub

Private Sub txtToStart_GotFocus()
Call SelectAll(txtToStart)

End Sub

Private Sub txtToStart_KeyUp(KeyCode As Integer, Shift As Integer)
Call CalcRange
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

Select Case cmbDB.ListIndex
    Case 1: 'Items
        CalcTotalRecords = CalcTotalRecords + tabShops.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabTextblocks.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabItems.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabRooms.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabMonsters.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabSpells.RecordCount
    Case 5: 'Spells
        CalcTotalRecords = CalcTotalRecords + tabItems.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabMonsters.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabRooms.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabTextblocks.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabSpells.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabSpells.RecordCount
    Case 7: 'Rooms
        CalcTotalRecords = CalcTotalRecords + tabTextblocks.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabSpells.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabRooms.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabRooms.RecordCount
    Case 4: 'Shops
        CalcTotalRecords = CalcTotalRecords + tabRooms.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabShops.RecordCount
    Case 3: 'Monsters
        CalcTotalRecords = CalcTotalRecords + tabMonsters.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabSpells.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabRooms.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabTextblocks.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabMonsters.RecordCount
    Case 2: 'Messages
        CalcTotalRecords = CalcTotalRecords + tabItems.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabMonsters.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabRooms.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabSpells.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabTextblocks.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabMessages.RecordCount
    Case 8: 'Textblocks
        CalcTotalRecords = CalcTotalRecords + tabClasses.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabItems.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabRooms.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabSpells.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabMonsters.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabTextblocks.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabTextblocks.RecordCount
    Case 0: 'Classes
        CalcTotalRecords = CalcTotalRecords + tabTextblocks.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabItems.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabRooms.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabShops.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabClasses.RecordCount
    Case 6: 'Races
        CalcTotalRecords = CalcTotalRecords + tabTextblocks.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabItems.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabRooms.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabRaces.RecordCount
    Case 9: 'map change
        If chkMapChange(2).Value = 1 Then CalcTotalRecords = CalcTotalRecords + tabTextblocks.RecordCount
        If chkMapChange(3).Value = 1 Then CalcTotalRecords = CalcTotalRecords + tabSpells.RecordCount
        If chkMapChange(0).Value = 1 Then CalcTotalRecords = CalcTotalRecords + tabRooms.RecordCount
        CalcTotalRecords = CalcTotalRecords + tabRooms.RecordCount
End Select

If CalcTotalRecords <= 0 Then CalcTotalRecords = 1
'If CalcTotalRecords > 32767 Then CalcTotalRecords = 32767

Exit Function

error:
Call HandleError
End Function

