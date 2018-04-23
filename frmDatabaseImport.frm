VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDatabaseImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Importer"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   ClipControls    =   0   'False
   Icon            =   "frmDatabaseImport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   8745
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview First"
      Height          =   315
      Left            =   7080
      TabIndex        =   60
      Top             =   2820
      Width           =   1335
   End
   Begin VB.Frame fraLog 
      Caption         =   "Logging Options"
      Height          =   1215
      Left            =   6900
      TabIndex        =   71
      Top             =   60
      Width           =   1695
      Begin VB.OptionButton optErrorsOnly 
         Caption         =   "Errors Only"
         Height          =   195
         Left            =   180
         TabIndex        =   74
         Top             =   900
         Width           =   1215
      End
      Begin VB.OptionButton optErrorsNSkips 
         Caption         =   "Errors && Skips"
         Height          =   195
         Left            =   180
         TabIndex        =   73
         Top             =   600
         Width           =   1395
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   195
         Left            =   180
         TabIndex        =   72
         Top             =   300
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame fraOpt2 
      Caption         =   "Existing Records"
      Height          =   975
      Left            =   6900
      TabIndex        =   75
      Top             =   1500
      Width           =   1695
      Begin VB.OptionButton optSkip 
         Caption         =   "Skip"
         Height          =   255
         Left            =   180
         TabIndex        =   58
         Top             =   300
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton optUpdate 
         Caption         =   "Update"
         Height          =   255
         Left            =   180
         TabIndex        =   59
         Top             =   600
         Width           =   1155
      End
   End
   Begin MSComctlLib.StatusBar stsStatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   77
      Top             =   5445
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12806
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   120
      TabIndex        =   76
      Top             =   5100
      Visible         =   0   'False
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   7800
      TabIndex        =   64
      Top             =   4140
      Width           =   795
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Import"
      Height          =   375
      Left            =   6900
      TabIndex        =   63
      Top             =   4140
      Width           =   795
   End
   Begin VB.Frame fraFile 
      Caption         =   "Select import file:"
      Height          =   4935
      Left            =   3780
      TabIndex        =   70
      Top             =   60
      Width           =   2955
      Begin VB.FileListBox filFileList 
         Height          =   1650
         Left            =   120
         Pattern         =   "*.mdb"
         TabIndex        =   67
         ToolTipText     =   "Double Click to Open"
         Top             =   3180
         Width           =   2715
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   65
         Top             =   240
         Width           =   2715
      End
      Begin VB.DirListBox Dir1 
         Height          =   2565
         Left            =   120
         TabIndex        =   66
         Top             =   540
         Width           =   2715
      End
   End
   Begin VB.Frame fraOpt 
      Caption         =   "Databases to Import"
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   3555
      Begin VB.CommandButton cmdQ 
         Caption         =   "?"
         Height          =   255
         Left            =   1560
         TabIndex        =   56
         Top             =   4080
         Width           =   210
      End
      Begin VB.CheckBox chkNotItems 
         Caption         =   "Not Items"
         Enabled         =   0   'False
         Height          =   255
         Left            =   420
         TabIndex        =   55
         Top             =   4200
         Width           =   1335
      End
      Begin VB.CheckBox chkOnlyItems 
         Caption         =   "Only Items"
         Enabled         =   0   'False
         Height          =   255
         Left            =   420
         TabIndex        =   54
         Top             =   3900
         Width           =   1335
      End
      Begin VB.CommandButton cmdCopyTo 
         Caption         =   ">"
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
         Index           =   0
         Left            =   2520
         TabIndex        =   6
         Top             =   420
         Width           =   255
      End
      Begin VB.CommandButton cmdCopyTo 
         Caption         =   ">"
         Height          =   255
         Index           =   9
         Left            =   2700
         TabIndex        =   52
         Top             =   3900
         Width           =   255
      End
      Begin VB.CommandButton cmdCopyTo 
         Caption         =   ">"
         Height          =   255
         Index           =   8
         Left            =   2580
         TabIndex        =   46
         Top             =   3180
         Width           =   135
      End
      Begin VB.CommandButton cmdCopyTo 
         Caption         =   ">"
         Height          =   255
         Index           =   7
         Left            =   2580
         TabIndex        =   41
         Top             =   2820
         Width           =   135
      End
      Begin VB.CommandButton cmdCopyTo 
         Caption         =   ">"
         Height          =   255
         Index           =   6
         Left            =   2580
         TabIndex        =   36
         Top             =   2460
         Width           =   135
      End
      Begin VB.CommandButton cmdCopyTo 
         Caption         =   ">"
         Height          =   255
         Index           =   5
         Left            =   2580
         TabIndex        =   31
         Top             =   2100
         Width           =   135
      End
      Begin VB.CommandButton cmdCopyTo 
         Caption         =   ">"
         Height          =   255
         Index           =   4
         Left            =   2580
         TabIndex        =   26
         Top             =   1740
         Width           =   135
      End
      Begin VB.CommandButton cmdCopyTo 
         Caption         =   ">"
         Height          =   255
         Index           =   3
         Left            =   2580
         TabIndex        =   21
         Top             =   1380
         Width           =   135
      End
      Begin VB.CommandButton cmdCopyTo 
         Caption         =   ">"
         Height          =   255
         Index           =   2
         Left            =   2580
         TabIndex        =   16
         Top             =   1020
         Width           =   135
      End
      Begin VB.CommandButton cmdCopyTo 
         Caption         =   ">"
         Height          =   255
         Index           =   1
         Left            =   2580
         TabIndex        =   11
         Top             =   660
         Width           =   135
      End
      Begin VB.CommandButton cmdNone 
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   300
         Width           =   555
      End
      Begin VB.CheckBox chkActionsAll 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1560
         TabIndex        =   69
         Top             =   4560
         Value           =   2  'Grayed
         Width           =   195
      End
      Begin VB.TextBox txtRoomsTo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2835
         TabIndex        =   53
         Text            =   "9999"
         Top             =   3555
         Width           =   555
      End
      Begin VB.CheckBox chkRoomsAll 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1575
         TabIndex        =   49
         Top             =   3600
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.TextBox txtRoomsMap 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1875
         TabIndex        =   50
         Text            =   "1"
         Top             =   3555
         Width           =   315
      End
      Begin VB.TextBox txtRoomsFrom 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2235
         TabIndex        =   51
         Text            =   "1"
         Top             =   3555
         Width           =   555
      End
      Begin VB.TextBox txtMessagesTo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   42
         Text            =   "9999"
         Top             =   2820
         Width           =   615
      End
      Begin VB.TextBox txtTextblocksTo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   47
         Text            =   "9999"
         Top             =   3180
         Width           =   615
      End
      Begin VB.CheckBox chkMessagesAll 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1560
         TabIndex        =   39
         Top             =   2880
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox chkTextblocksAll 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1560
         TabIndex        =   44
         Top             =   3240
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.TextBox txtMessagesFrom 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   40
         Text            =   "1"
         Top             =   2820
         Width           =   615
      End
      Begin VB.TextBox txtTextblocksFrom 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   45
         Text            =   "0"
         Top             =   3180
         Width           =   615
      End
      Begin VB.TextBox txtItemsTo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   12
         Text            =   "9999"
         Top             =   675
         Width           =   615
      End
      Begin VB.CheckBox chkItemsAll 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1560
         TabIndex        =   9
         Top             =   720
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.TextBox txtMonstersTo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   17
         Text            =   "9999"
         Top             =   1035
         Width           =   615
      End
      Begin VB.TextBox txtSpellsTo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   22
         Text            =   "9999"
         Top             =   1395
         Width           =   615
      End
      Begin VB.TextBox txtShopsTo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   27
         Text            =   "9999"
         Top             =   1755
         Width           =   615
      End
      Begin VB.TextBox txtRacesTo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   32
         Text            =   "9999"
         Top             =   2115
         Width           =   615
      End
      Begin VB.TextBox txtClassesTo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   37
         Text            =   "9999"
         Top             =   2475
         Width           =   615
      End
      Begin VB.CheckBox chkMonstersAll 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1560
         TabIndex        =   14
         Top             =   1080
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox chkSpellsAll 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1560
         TabIndex        =   19
         Top             =   1440
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox chkShopsAll 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1560
         TabIndex        =   24
         Top             =   1800
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox chkRacesAll 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1560
         TabIndex        =   29
         Top             =   2175
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox chkClassesAll 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1560
         TabIndex        =   34
         Top             =   2535
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.TextBox txtItemsFrom 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   10
         Text            =   "1"
         Top             =   675
         Width           =   615
      End
      Begin VB.TextBox txtMonstersFrom 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   15
         Text            =   "1"
         Top             =   1035
         Width           =   615
      End
      Begin VB.TextBox txtSpellsFrom 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   20
         Text            =   "1"
         Top             =   1395
         Width           =   615
      End
      Begin VB.TextBox txtShopsFrom 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   25
         Text            =   "1"
         Top             =   1755
         Width           =   615
      End
      Begin VB.TextBox txtRacesFrom 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   30
         Text            =   "1"
         Top             =   2115
         Width           =   615
      End
      Begin VB.TextBox txtClassesFrom 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   35
         Text            =   "1"
         Top             =   2475
         Width           =   615
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   555
      End
      Begin VB.CheckBox chkMonsters 
         Caption         =   "Monsters"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   1080
         Width           =   1395
      End
      Begin VB.CheckBox chkSpells 
         Caption         =   "Spells"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   1440
         Width           =   1395
      End
      Begin VB.CheckBox chkShops 
         Caption         =   "Shops"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   23
         Top             =   1800
         Width           =   1395
      End
      Begin VB.CheckBox chkRaces 
         Caption         =   "Races"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   28
         Top             =   2160
         Width           =   1395
      End
      Begin VB.CheckBox chkClasses 
         Caption         =   "Classes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   33
         Top             =   2520
         Width           =   1395
      End
      Begin VB.CheckBox chkRooms 
         Caption         =   "Rooms"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   48
         Top             =   3600
         Width           =   1395
      End
      Begin VB.CheckBox chkActions 
         Caption         =   "Actions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   57
         Top             =   4560
         Width           =   1275
      End
      Begin VB.CheckBox chkMessages 
         Caption         =   "Messages"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   38
         Top             =   2880
         Width           =   1395
      End
      Begin VB.CheckBox chkTextblocks 
         Caption         =   "Textblocks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   43
         Top             =   3240
         Width           =   1395
      End
      Begin VB.CheckBox chkItems 
         Caption         =   "Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label17 
         Caption         =   "Map"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1875
         TabIndex        =   68
         Top             =   3825
         Width           =   315
      End
      Begin VB.Label lblAll 
         Caption         =   "All"
         Height          =   195
         Left            =   1560
         TabIndex        =   4
         Top             =   405
         Width           =   195
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         Top             =   405
         Width           =   615
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "From"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   405
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "| -------------- range -------------- |"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1500
         TabIndex        =   1
         Top             =   180
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "&Log"
      Height          =   375
      Left            =   7800
      TabIndex        =   62
      Top             =   3600
      Width           =   795
   End
   Begin VB.CommandButton cmdNotes 
      Caption         =   "*&READ!*"
      Height          =   375
      Left            =   6900
      TabIndex        =   61
      Top             =   3600
      Width           =   795
   End
End
Attribute VB_Name = "frmDatabaseImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim DB As Database
Dim tabActions As Recordset
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

Dim bStopImport As Boolean
Dim bPreview As Boolean
Dim bSkipMissing As Boolean
Dim sDataSource As String
Dim nFilesToImport As Integer
Dim sLogFile As String
Dim ts As TextStream
Dim fso As FileSystemObject


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

Private Sub cmdQ_Click()
MsgBox "Using these options will make the importer either skip the hidden" & vbCrLf _
    & "and visible items in the room, or import only them and skip everything else." _
    & vbCrLf & vbCrLf & "(placed items are unaffected by this)", vbInformation
End Sub



Private Sub chkNotItems_Click()
If chkOnlyItems.Value = 1 And chkNotItems.Value = 1 Then chkOnlyItems.Value = 0
End Sub

Private Sub chkOnlyItems_Click()
If chkOnlyItems.Value = 1 And chkNotItems.Value = 1 Then chkNotItems.Value = 0
End Sub

Private Sub chkRooms_Click()
If chkRooms.Value = 1 Then
    chkOnlyItems.Enabled = True
    chkNotItems.Enabled = True
Else
    chkOnlyItems.Enabled = False
    chkNotItems.Enabled = False
End If
End Sub

Private Sub cmdCopyTo_Click(Index As Integer)
Dim x As Integer

x = Index
again:

Select Case x
    Case 1: txtItemsTo.Text = txtItemsFrom.Text
    Case 2: txtMonstersTo.Text = txtMonstersFrom.Text
    Case 3: txtSpellsTo.Text = txtSpellsFrom.Text
    Case 4: txtShopsTo.Text = txtShopsFrom.Text
    Case 5: txtRacesTo.Text = txtRacesFrom.Text
    Case 6: txtClassesTo.Text = txtClassesFrom.Text
    Case 7: txtMessagesTo.Text = txtMessagesFrom.Text
    Case 8: txtTextblocksTo.Text = txtTextblocksFrom.Text
    Case 9: txtRoomsTo.Text = txtRoomsFrom.Text
End Select

If Index = 0 Then
    x = x + 1
    If x <= 9 Then GoTo again:
End If

End Sub

Private Sub Form_Load()
On Error Resume Next

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FolderExists(ReadINI("Options", "ImportPath")) = True Then
    Dir1.Path = ReadINI("Options", "ImportPath")
Else
    Dir1.Path = App.Path
End If
Drive1.Drive = Dir1.Path

chkMessagesAll.Value = ReadINI("Options", "ImportMessagesAll")
txtMessagesFrom.Text = ReadINI("Options", "ImportMessagesFrom")
txtMessagesTo.Text = ReadINI("Options", "ImportMessagesTo")
chkRoomsAll.Value = ReadINI("Options", "ImportRoomsAll")
txtRoomsFrom.Text = ReadINI("Options", "ImportRoomsFrom")
txtRoomsTo.Text = ReadINI("Options", "ImportRoomsTo")
txtRoomsMap.Text = ReadINI("Options", "ImportRoomsMap")
chkShopsAll.Value = ReadINI("Options", "ImportShopsAll")
txtShopsFrom.Text = ReadINI("Options", "ImportShopsFrom")
txtShopsTo.Text = ReadINI("Options", "ImportShopsTo")
chkSpellsAll.Value = ReadINI("Options", "ImportSpellsAll")
txtSpellsFrom.Text = ReadINI("Options", "ImportSpellsFrom")
txtSpellsTo.Text = ReadINI("Options", "ImportSpellsTo")
chkItemsAll.Value = ReadINI("Options", "ImportItemsAll")
txtItemsFrom.Text = ReadINI("Options", "ImportItemsFrom")
txtItemsTo.Text = ReadINI("Options", "ImportItemsTo")
chkTextblocksAll.Value = ReadINI("Options", "ImportTextblocksAll")
txtTextblocksFrom.Text = ReadINI("Options", "ImportTextblocksFrom")
txtTextblocksTo.Text = ReadINI("Options", "ImportTextblocksTo")
chkRacesAll.Value = ReadINI("Options", "ImportRacesAll")
txtRacesFrom.Text = ReadINI("Options", "ImportRacesFrom")
txtRacesTo.Text = ReadINI("Options", "ImportRacesTo")
chkClassesAll.Value = ReadINI("Options", "ImportClassesAll")
txtClassesFrom.Text = ReadINI("Options", "ImportClassesFrom")
txtClassesTo.Text = ReadINI("Options", "ImportClassesTo")
chkMonstersAll.Value = ReadINI("Options", "ImportMonstersAll")
txtMonstersFrom.Text = ReadINI("Options", "ImportMonstersFrom")
txtMonstersTo.Text = ReadINI("Options", "ImportMonstersTo")

Me.Top = ReadINI("Windows", "ImportTop")
Me.Left = ReadINI("Windows", "ImportLeft")
If ReadINI("Options", "ImportUpdate") = "1" Then optUpdate = True
If ReadINI("Options", "ImportErrorsSkips") = "1" Then
    optErrorsNSkips.Value = True
ElseIf ReadINI("Options", "ImportErrorsOnly") = "1" Then
    optErrorsOnly.Value = True
Else
    optAll.Value = True
End If

Me.Show
Me.SetFocus
cmdCancel.SetFocus

End Sub

Private Sub chkClassesAll_Click()
If chkClassesAll.Value = 1 Then
    txtClassesFrom.Enabled = False
    txtClassesTo.Enabled = False
Else
    txtClassesFrom.Enabled = True
    txtClassesTo.Enabled = True
End If
End Sub

Private Sub chkItemsAll_Click()
If chkItemsAll.Value = 1 Then
    txtItemsFrom.Enabled = False
    txtItemsTo.Enabled = False
Else
    txtItemsFrom.Enabled = True
    txtItemsTo.Enabled = True
End If

End Sub

Private Sub chkMessagesAll_Click()
If chkMessagesAll.Value = 1 Then
    txtMessagesFrom.Enabled = False
    txtMessagesTo.Enabled = False
Else
    txtMessagesFrom.Enabled = True
    txtMessagesTo.Enabled = True
End If
End Sub

Private Sub chkMonstersAll_Click()
If chkMonstersAll.Value = 1 Then
    txtMonstersFrom.Enabled = False
    txtMonstersTo.Enabled = False
Else
    txtMonstersFrom.Enabled = True
    txtMonstersTo.Enabled = True
End If
End Sub

Private Sub chkRacesAll_Click()
If chkRacesAll.Value = 1 Then
    txtRacesFrom.Enabled = False
    txtRacesTo.Enabled = False
Else
    txtRacesFrom.Enabled = True
    txtRacesTo.Enabled = True
End If
End Sub

Private Sub chkRoomsAll_Click()
If chkRoomsAll.Value = 1 Then
    txtRoomsFrom.Enabled = False
    txtRoomsTo.Enabled = False
    txtRoomsMap.Enabled = False
Else
    txtRoomsFrom.Enabled = True
    txtRoomsTo.Enabled = True
    txtRoomsMap.Enabled = True
End If
End Sub

Private Sub chkShopsAll_Click()
If chkShopsAll.Value = 1 Then
    txtShopsFrom.Enabled = False
    txtShopsTo.Enabled = False
Else
    txtShopsFrom.Enabled = True
    txtShopsTo.Enabled = True
End If
End Sub

Private Sub chkSpellsAll_Click()
If chkSpellsAll.Value = 1 Then
    txtSpellsFrom.Enabled = False
    txtSpellsTo.Enabled = False
Else
    txtSpellsFrom.Enabled = True
    txtSpellsTo.Enabled = True
End If
End Sub

Private Sub chkTextblocksAll_Click()
If chkTextblocksAll.Value = 1 Then
    txtTextblocksFrom.Enabled = False
    txtTextblocksTo.Enabled = False
Else
    txtTextblocksFrom.Enabled = True
    txtTextblocksTo.Enabled = True
End If
End Sub

Private Sub filFileList_DblClick()
Dim fso As FileSystemObject

If filFileList.FileName = "" Then
    MsgBox "You must select a file first!", vbInformation + vbOKOnly
    Exit Sub
End If

Set fso = CreateObject("Scripting.FileSystemObject")

sDataSource = filFileList.FileName
If Right(Dir1.Path, 1) = "\" Then
    sDataSource = Dir1.Path & sDataSource
Else
    sDataSource = Dir1.Path & "\" & sDataSource
End If


If fso.FileExists(sDataSource) = True Then
    Call ShellExecute(0&, "open", sDataSource, vbNullString, vbNullString, vbNormalFocus)
Else
    MsgBox sDataSource & " was not found.", vbInformation
End If

Set fso = Nothing
End Sub


Private Sub OpenTables()
On Error GoTo error:

Set tabRooms = DB.OpenRecordset("Rooms")
Set tabItems = DB.OpenRecordset("Items")
Set tabClasses = DB.OpenRecordset("Classes")
Set tabRaces = DB.OpenRecordset("Races")
Set tabSpells = DB.OpenRecordset("Spells")
Set tabActions = DB.OpenRecordset("Actions")
Set tabMonsters = DB.OpenRecordset("Monsters")
Set tabShops = DB.OpenRecordset("Shops")
Set tabMessages = DB.OpenRecordset("Messages")
Set tabTextblocks = DB.OpenRecordset("Textblocks")
Set tabInfo = DB.OpenRecordset("Info")

Exit Sub
error:
Call HandleError
Resume Next

End Sub
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
tabActions.Close

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
Set tabActions = Nothing

Set DB = Nothing

End Sub

Private Sub cmdAll_Click()

chkMessages.Value = 1
chkItems.Value = 1
chkSpells.Value = 1
chkClasses.Value = 1
chkRaces.Value = 1
chkShops.Value = 1
chkRooms.Value = 1
chkActions.Value = 1
chkMonsters.Value = 1
chkTextblocks.Value = 1

End Sub

Private Sub cmdLog_Click()

On Error GoTo error:

If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")

If Right(Dir1.Path, 1) = "\" Then
    sLogFile = Dir1.Path & "NMR-Log_Import.txt"
Else
    sLogFile = Dir1.Path & "\NMR-Log_Import.txt"
End If

If fso.FileExists(sLogFile) = True Then
    Call ShellExecute(0&, "open", sLogFile, vbNullString, vbNullString, vbNormalFocus)
Else
    MsgBox sLogFile & " was not found.", vbInformation
End If

out:
Exit Sub
error:
Call HandleError("cmdLog_Click")
Resume out:
End Sub

Private Sub cmdNone_Click()
chkMessages.Value = 0
chkItems.Value = 0
chkSpells.Value = 0
chkClasses.Value = 0
chkRaces.Value = 0
chkShops.Value = 0
chkRooms.Value = 0
chkActions.Value = 0
chkMonsters.Value = 0
chkTextblocks.Value = 0
End Sub

Private Sub cmdNotes_Click()
MsgBox "Notes on the Database Importer:" & vbCrLf _
& "------------------------------------------" & vbCrLf _
& "While NMR edits a lot, there are still values within the dats that are undiscovered." & vbCrLf _
& "What this means is that you need to keep in mind that some exported records may be incomplete." & vbCrLf _
& "This would really only be an issue if you were importing original records that didn't exist currently in your installation.", vbInformation
End Sub


Private Sub Dir1_Change()
filFileList.Path = Dir1.Path
End Sub

Private Sub cmdGo_Click()

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

If filFileList.FileName = "" Then
    MsgBox "You must select a file to import from first!", vbInformation + vbOKOnly
    Exit Sub
End If

bPreview = False
Call DoImport

End Sub
Private Sub cmdPreview_Click()
Dim nYesNo As Integer

If filFileList.FileName = "" Then
    MsgBox "You must select a file to import from first!", vbInformation + vbOKOnly
    Exit Sub
End If

nYesNo = MsgBox("PREVIEW this import job?", vbYesNo, "Preview?")
If nYesNo = vbNo Then Exit Sub

bPreview = True
Call DoImport

End Sub
Private Sub DoImport()
On Error GoTo error:
Dim x As Integer, bTest As Boolean, nYesNo As Integer
Dim CheckboxArray(1 To 10) As Object

bSkipMissing = False
bStopImport = False

sDataSource = filFileList.FileName
If Right(Dir1.Path, 1) = "\" Then
    sDataSource = Dir1.Path & sDataSource
    sLogFile = Dir1.Path & "NMR-Log_Import.txt"
Else
    sDataSource = Dir1.Path & "\" & sDataSource
    sLogFile = Dir1.Path & "\NMR-Log_Import.txt"
End If

If Not bPreview Then
    nYesNo = MsgBox("Are you sure you want to import from" & vbCrLf & sDataSource, vbYesNo, "Confirm Action")
    If nYesNo = vbNo Then Exit Sub
End If

If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(sDataSource) = False Then MsgBox sDataSource & " was not found.": Exit Sub

UnloadForms (Me.Name)

DoEvents
'cmdCancel.Enabled = False
'frmMain.Enabled = False
cmdGo.Enabled = False
cmdPreview.Enabled = False
cmdNotes.Enabled = False
cmdLog.Enabled = False
fraOpt2.Enabled = False
fraLog.Enabled = False
fraFile.Enabled = False
fraOpt.Enabled = False
cmdCancel.Caption = "&Cancel"
nFilesToImport = 0
Call LockMenus

Set CheckboxArray(1) = chkMessages
Set CheckboxArray(2) = chkItems
Set CheckboxArray(3) = chkSpells
Set CheckboxArray(4) = chkClasses
Set CheckboxArray(5) = chkRaces
Set CheckboxArray(6) = chkShops
Set CheckboxArray(7) = chkRooms
Set CheckboxArray(8) = chkActions
Set CheckboxArray(9) = chkMonsters
Set CheckboxArray(10) = chkTextblocks

Set DB = OpenDatabase(sDataSource)
Call OpenTables

bTest = CheckVersion
If bTest <> True Then GoTo out:

Call CreateLogFile(sLogFile)

Set ts = fso.OpenTextFile(sLogFile, ForWriting)

Call SetRange(CalcTotalRecords)
ProgressBar.Visible = True

ts.WriteLine ("Import job started " & Date & " @ " & Time)

If bPreview Then
    ts.WriteBlankLines (1)
    ts.WriteLine ("*** PREVIEW ONLY -- NO INSERTS/UPDATES ARE ACTUALLY EXECUTED ***")
End If

For x = 1 To UBound(CheckboxArray())
    If bStopImport Then Exit For
    If CheckboxArray(x).Value = 1 Then
        If x = 1 Then Call ImportMessages
        If x = 2 Then Call ImportItems
        If x = 3 Then Call ImportSpells
        If x = 4 Then Call ImportClasses
        If x = 5 Then Call ImportRaces
        If x = 6 Then Call ImportShops
        If x = 7 Then Call ImportRooms
        If x = 8 Then Call ImportActions
        If x = 9 Then Call ImportMonsters
        If x = 10 Then Call ImportTextblocks
        DoEvents
    End If
Next

If bStopImport Then
    ts.WriteBlankLines (1)
    ts.WriteLine ("* CANCELED BY USER * - " & Date & " @ " & Time)
    ts.Close
    
    nYesNo = MsgBox("Import canceled, view log file?", vbYesNo + vbInformation)
    If nYesNo = vbYes Then Call cmdLog_Click

    GoTo out:
Else
    ts.WriteBlankLines (1)
    ts.WriteLine ("Complete - " & Date & " @ " & Time)
    ts.Close
End If

ProgressBar.Value = ProgressBar.Max
nYesNo = MsgBox("Import complete, view log file?", vbYesNo + vbInformation)
If nYesNo = vbYes Then Call cmdLog_Click

out:
On Error Resume Next
Call CloseAll
DoEvents
cmdGo.Enabled = True
cmdPreview.Enabled = True
cmdNotes.Enabled = True
cmdLog.Enabled = True
fraOpt2.Enabled = True
fraLog.Enabled = True
fraFile.Enabled = True
fraOpt.Enabled = True
cmdCancel.Enabled = True
cmdCancel.Caption = "&Close"

ProgressBar.Visible = False
stsStatusBar.Panels(1).Text = ""
stsStatusBar.Panels(2).Text = ""
frmMain.Enabled = True
Call UnLockMenus

Exit Sub

error:
Call HandleError("DoImport")
Resume out:
End Sub

Private Sub cmdCancel_Click()
Dim nYesNo As Integer

If cmdGo.Enabled = False Then
    nYesNo = MsgBox("Are you sure you want to cancel?", vbYesNo + vbQuestion + vbDefaultButton2)
    If Not nYesNo = vbYes Then Exit Sub

    cmdCancel.Enabled = False
    bStopImport = True
    DoEvents
Else
    Unload Me
End If

End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
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
Private Sub ImportTextblocks()
On Error GoTo error:
Dim nStatus As Integer, decrypted As String, nLastRec(1) As Long
Dim recnum As Long, x As Integer, ExistingRecord As Boolean

stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_TEXT

If tabTextblocks.RecordCount = 0 Then Exit Sub
tabTextblocks.MoveFirst
ts.WriteBlankLines (2)

tabTextblocks.Index = "idxTextblocks"
If chkTextblocksAll.Value = 0 Then
    tabTextblocks.Seek "=", Val(txtTextblocksFrom.Text), 0
    If tabTextblocks.NoMatch Then tabTextblocks.MoveFirst
End If
nLastRec(0) = tabTextblocks.Fields("Number")
nLastRec(1) = 0

Do While tabTextblocks.EOF = False And bStopImport = False
    
    'check for extra textblock parts
    If nLastRec(0) <> tabTextblocks.Fields("Number") Then
        TextblockKey.Number = nLastRec(0)
        TextblockKey.PartNum = nLastRec(1)
part_check:
        TextblockKey.PartNum = TextblockKey.PartNum + 1
        
        nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKeyStructToRow(), KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            nStatus = BTRCALL(BDELETE, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
            If nStatus <> 0 Then
                ts.WriteLine ("Textblock #" & TextblockKey.Number & ", Part " & TextblockKey.PartNum & " -- Error deleting extra textblock part: " & nStatus)
            Else
                If optErrorsNSkips.Value = False And optErrorsOnly.Value = False Then
                    ts.WriteLine ("Textblock #" & TextblockKey.Number & ", Part " & TextblockKey.PartNum & " -- Deleted Extra Textblock Part.")
                End If
                GoTo part_check:
            End If
        End If
    End If
    
    recnum = tabTextblocks.Fields("Number")
    stsStatusBar.Panels(2).Text = recnum
    IncreaseProgressBar
    
    TextblockKey.PartNum = tabTextblocks.Fields("Part #")
    TextblockKey.Number = tabTextblocks.Fields("Number")
    
    If chkTextblocksAll.Value = 0 Then
        If TextblockKey.Number < Val(txtTextblocksFrom.Text) Then
            If optErrorsOnly.Value = True Then GoTo SkipRecord
            ts.WriteLine ("Textblock #" & tabTextblocks.Fields("Number") & " -- Skipped, out of import range.")
            GoTo SkipRecord:
        ElseIf TextblockKey.Number > Val(txtTextblocksTo.Text) Then
            Exit Sub
        End If
    End If
    
    ExistingRecord = True
    nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKeyStructToRow(), KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nStatus = 4 Then
            ExistingRecord = False
        Else
            ts.WriteLine ("Textblock #" & tabTextblocks.Fields("Number") & ", Part " & tabTextblocks.Fields("Part #") & " -- Error " & nStatus & " determining if record exists.")
            GoTo SkipRecord:
        End If
    Else
        If optSkip.Value = True Then
            If optErrorsOnly.Value = True Then GoTo SkipRecord
            ts.WriteLine ("Textblock #" & tabTextblocks.Fields("Number") & ", Part " & tabTextblocks.Fields("Part #") & " -- Existing record. Skipped due to setting.")
            GoTo SkipRecord:
        End If
    End If
    
    TextblockRowToStruct TextblockDataBuf.buf

    TextblockRec.Number = tabTextblocks.Fields("Number")
    TextblockRec.PartNum = tabTextblocks.Fields("Part #")
    TextblockRec.LinkTo = tabTextblocks.Fields("Link To")
    
    decrypted = ""
'    For x = 1 To 8
'        decrypted = decrypted & tabTextblocks.Fields(CStr("Data Part " & x))
'    Next

    decrypted = tabTextblocks.Fields("Data")
    
    TextblockRec.Data = EncryptTextblock(decrypted)
    
    If ExistingRecord = True Then
        iUpdateTextblock
    Else
        For x = 1 To 14
            TextblockRec.LeadIn(x) = TextblockKey.LeadIn(x)
        Next

        TextblockStructToRow TextblockDataBuf.buf
        
        If bPreview Then
            ts.WriteLine ("Textblock #" & tabTextblocks.Fields("Number") & ", Part " & tabTextblocks.Fields("Part #") & " -- Non-Existing Record, would be inserted.")
        Else
            nStatus = BTRCALL(BINSERT, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
            If Not nStatus = 0 Then
                ts.WriteLine ("Textblock #" & tabTextblocks.Fields("Number") & ", Part " & tabTextblocks.Fields("Part #") & " -- Insert Error: " & nStatus)
            Else
                If optErrorsNSkips.Value = True Then GoTo SkipRecord:
                If optErrorsOnly.Value = True Then GoTo SkipRecord:
                ts.WriteLine ("Textblock #" & tabTextblocks.Fields("Number") & ", Part " & tabTextblocks.Fields("Part #") & " -- Insert Successful.")
            End If
        End If
    End If
        
SkipRecord:
    nLastRec(0) = tabTextblocks.Fields("Number")
    nLastRec(1) = tabTextblocks.Fields("Part #")
    tabTextblocks.MoveNext
    If Not bUseCPU Then DoEvents
Loop

Exit Sub
error:
If CheckError = True Then Resume Next
End Sub
Private Sub iUpdateTextblock()
Dim nStatus As Integer

TextblockStructToRow TextblockDataBuf.buf

If bPreview Then
    ts.WriteLine ("Textblock #" & TextblockRec.Number & ", Part " & TextblockRec.PartNum & " -- Existing Record, would be updated.")
Else
    nStatus = BTRCALL(BUPDATE, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        ts.WriteLine ("Textblock #" & TextblockRec.Number & ", Part " & TextblockRec.PartNum & " -- Update Error: " & nStatus)
    Else
        If optErrorsOnly.Value = True Or optErrorsNSkips.Value = True Then Exit Sub
        ts.WriteLine ("Textblock #" & TextblockRec.Number & ", Part " & TextblockRec.PartNum & " -- Update Successful.")
    End If
End If
End Sub

Private Sub ImportMessages()
On Error GoTo error:
Dim nStatus As Integer, recnum As Long, x As Long
Dim ExistingRecord As Boolean

stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_MSG

If tabMessages.RecordCount = 0 Then Exit Sub
tabMessages.MoveFirst
ts.WriteBlankLines (2)

tabMessages.Index = "pkMessages"
If chkMessagesAll.Value = 0 Then
    tabMessages.Seek "=", Val(txtMessagesFrom.Text)
    If tabMessages.NoMatch Then tabMessages.MoveFirst
End If

Do While tabMessages.EOF = False And bStopImport = False
    
    recnum = tabMessages.Fields("Number")
    stsStatusBar.Panels(2).Text = recnum
    IncreaseProgressBar
    
    x = tabMessages.Fields("Number")

    If chkMessagesAll.Value = 0 Then
        If x < Val(txtMessagesFrom.Text) Then
            If optErrorsOnly.Value = True Then GoTo SkipRecord
            ts.WriteLine ("Message #" & tabMessages.Fields("Number") & " -- Skipped, out of import range.")
            GoTo SkipRecord:
        ElseIf x > Val(txtMessagesTo.Text) Then
            Exit Sub
        End If
    End If
    
    ExistingRecord = True
    nStatus = BTRCALL(BGETEQUAL, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), x, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nStatus = 4 Then
            ExistingRecord = False
        Else
            ts.WriteLine ("Message #" & tabMessages.Fields("Number") & " -- Error " & nStatus & " determining if record exists.")
            GoTo SkipRecord:
        End If
    Else
        If optSkip.Value = True Then
            If optErrorsOnly.Value = True Then GoTo SkipRecord
            ts.WriteLine ("Message #" & tabMessages.Fields("Number") & " -- Existing record. Skipped due to setting.")
            GoTo SkipRecord:
        End If
    End If
    
    MessageRowToStruct Messagedatabuf.buf

    Messagerec.Number = tabMessages.Fields("Number")
    Messagerec.MessageLine1 = Trim(tabMessages.Fields("Line 1"))
    Messagerec.MessageLine2 = Trim(tabMessages.Fields("Line 2"))
    Messagerec.MessageLine3 = Trim(tabMessages.Fields("Line 3"))
    
            
        If ExistingRecord = True Then
            iUpdateMessage
        Else

            MessageStructToRow Messagedatabuf.buf
            
            If bPreview Then
                ts.WriteLine ("Message #" & tabMessages.Fields("Number") & " -- Non-Existing Record, would be inserted.")
            Else
                nStatus = BTRCALL(BINSERT, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
                If Not nStatus = 0 Then
                    ts.WriteLine ("Message #" & tabMessages.Fields("Number") & " -- Insert Error: " & nStatus)
                Else
                    If optErrorsNSkips.Value = True Then GoTo SkipRecord:
                    If optErrorsOnly.Value = True Then GoTo SkipRecord:
                    ts.WriteLine ("Message #" & tabMessages.Fields("Number") & " -- Insert Successful.")
                End If
            End If
        End If
        
SkipRecord:
        tabMessages.MoveNext
        If Not bUseCPU Then DoEvents
Loop

Exit Sub
error:
If CheckError = True Then Resume Next
End Sub
Private Sub iUpdateMessage()
Dim nStatus As Integer

MessageStructToRow Messagedatabuf.buf

If bPreview Then
    ts.WriteLine ("Message #" & Messagerec.Number & " -- Existing Record, would be updated.")
Else
    nStatus = BTRCALL(BUPDATE, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        ts.WriteLine ("Message #" & Messagerec.Number & " -- Update Error: " & nStatus)
    Else
        If optErrorsOnly.Value = True Or optErrorsNSkips.Value = True Then Exit Sub
        ts.WriteLine ("Message #" & Messagerec.Number & " -- Update Successful.")
    End If
End If
End Sub
Private Sub ImportItems()
On Error GoTo error:
Dim nStatus As Integer, recnum As Long, x As Long
Dim ExistingRecord As Boolean

stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_ITEMS

If tabItems.RecordCount = 0 Then Exit Sub
tabItems.MoveFirst
ts.WriteBlankLines (2)

tabItems.Index = "pkItems"
If chkItemsAll.Value = 0 Then
    tabItems.Seek "=", Val(txtItemsFrom.Text)
    If tabItems.NoMatch Then tabItems.MoveFirst
End If

Do While tabItems.EOF = False And bStopImport = False
    
    recnum = tabItems.Fields("Number")
    stsStatusBar.Panels(2).Text = recnum
    IncreaseProgressBar
    
    x = tabItems.Fields("Number")
    
    If chkItemsAll.Value = 0 Then
        If x < Val(txtItemsFrom.Text) Then
            If optErrorsOnly.Value = True Then GoTo SkipRecord:
            ts.WriteLine ("Item #" & tabItems.Fields("Number") & " [" & ClipNull(tabItems.Fields("Name")) & "] -- Skipped, out of import range.")
            GoTo SkipRecord:
        ElseIf x > Val(txtItemsTo.Text) Then
            Exit Sub
        End If
    End If
    
    ExistingRecord = True
    nStatus = BTRCALL(BGETEQUAL, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), x, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nStatus = 4 Then
            ExistingRecord = False
        Else
            ts.WriteLine ("Item #" & tabItems.Fields("Number") & " [" & ClipNull(tabItems.Fields("Name")) & "] -- Error " & nStatus & " determining if record exists.")
            GoTo SkipRecord:
        End If
    Else
        If optSkip.Value = True Then
            If optErrorsOnly.Value = True Then GoTo SkipRecord:
            ts.WriteLine ("Item #" & tabItems.Fields("Number") & " [" & ClipNull(tabItems.Fields("Name")) & "] -- Existing record. Skipped due to setting.")
            GoTo SkipRecord:
        End If
    End If
    
    ItemRowToStruct Itemdatabuf.buf

    Itemrec.Number = tabItems.Fields("Number")
    Itemrec.Name = Trim(tabItems.Fields("Name"))
    Itemrec.GameLimit = tabItems.Fields("Game Limit")
    Itemrec.Desc1 = Trim(tabItems.Fields("Desc1"))
    Itemrec.Desc2 = Trim(tabItems.Fields("Desc2"))
    Itemrec.Desc3 = Trim(tabItems.Fields("Desc3"))
    Itemrec.Desc4 = Trim(tabItems.Fields("Desc4"))
    Itemrec.Desc5 = Trim(tabItems.Fields("Desc5"))
    Itemrec.Desc6 = Trim(tabItems.Fields("Desc6"))
    Itemrec.Weight = tabItems.Fields("Weight")
    Itemrec.Type = tabItems.Fields("Type")
    Itemrec.Uses = tabItems.Fields("Uses")
    Itemrec.Cost = tabItems.Fields("Cost")
    Itemrec.CostType = tabItems.Fields("Cost Type")
    Itemrec.Minhit = tabItems.Fields("Min Hit")
    Itemrec.Maxhit = tabItems.Fields("Max Hit")
    Itemrec.AC = tabItems.Fields("AC")
    Itemrec.DR = tabItems.Fields("DR")
    Itemrec.Weapon = tabItems.Fields("Weapon")
    Itemrec.Armour = tabItems.Fields("Armour")
    Itemrec.WornOn = tabItems.Fields("Worn On")
    Itemrec.Accuracy = tabItems.Fields("Accuracy")
    Itemrec.Gettable = tabItems.Fields("Gettable")
    Itemrec.ReqStr = tabItems.Fields("Req Str")
    Itemrec.Speed = tabItems.Fields("Speed")
    Itemrec.Robable = tabItems.Fields("Robable")
    Itemrec.HitMsg = tabItems.Fields("Hit Msg")
    Itemrec.MissMsg = tabItems.Fields("Miss Msg")
    Itemrec.ReadTB = tabItems.Fields("Read Msg")
    Itemrec.DistructMsg = tabItems.Fields("Distruct Msg")
    Itemrec.NotDroppable = tabItems.Fields("Not Droppable")
    Itemrec.DestroyOnDeath = tabItems.Fields("Destroy On Death")
    Itemrec.RetainAfterUses = tabItems.Fields("Retain After Uses")
    Itemrec.OpenRunic = tabItems.Fields("OpenRunic")
    Itemrec.OpenPlatinum = tabItems.Fields("OpenPlatinum")
    Itemrec.OpenGold = tabItems.Fields("OpenGold")
    Itemrec.OpenSilver = tabItems.Fields("OpenSilver")
    Itemrec.OpenCopper = tabItems.Fields("OpenCopper")
    
    For x = 0 To 9
        Itemrec.Class(x) = tabItems.Fields("Class " & x)
    Next

    For x = 0 To 9
        Itemrec.Race(x) = tabItems.Fields("Race " & x)
    Next

    For x = 0 To 9
        Itemrec.Negate(x * 2) = tabItems.Fields("Negate " & x)
    Next
    
    For x = 0 To 19
        Itemrec.AbilityA(x) = tabItems.Fields("Ability " & x)
        Itemrec.AbilityB(x) = tabItems.Fields("Ability Value " & x)
    Next
            
        If ExistingRecord = True Then
            iUpdateItem
        Else

            ItemStructToRow Itemdatabuf.buf
            
            If bPreview Then
                ts.WriteLine ("Item #" & tabItems.Fields("Number") & " [" & ClipNull(tabItems.Fields("Name")) & "] -- Non-Existing Record, would be inserted.")
            Else
                nStatus = BTRCALL(BINSERT, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
                If Not nStatus = 0 Then
                    ts.WriteLine ("Item #" & tabItems.Fields("Number") & " [" & ClipNull(tabItems.Fields("Name")) & "] -- Insert Error: " & nStatus)
                Else
                    If optErrorsNSkips.Value = True Then GoTo SkipRecord:
                    If optErrorsOnly.Value = True Then GoTo SkipRecord:
                    ts.WriteLine ("Item #" & tabItems.Fields("Number") & " [" & ClipNull(tabItems.Fields("Name")) & "] -- Insert Successful.")
                End If
            End If
        End If
        
SkipRecord:
        tabItems.MoveNext
        If Not bUseCPU Then DoEvents
Loop

Exit Sub
error:
If CheckError = True Then Resume Next
End Sub
Private Sub iUpdateItem()
Dim nStatus As Integer

ItemStructToRow Itemdatabuf.buf

If bPreview Then
    ts.WriteLine ("Item #" & Itemrec.Number & " [" & ClipNull(Itemrec.Name) & "] -- Existing Record, would be updated.")
Else
    nStatus = BTRCALL(BUPDATE, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        ts.WriteLine ("Item #" & Itemrec.Number & " [" & ClipNull(Itemrec.Name) & "] -- Update Error: " & nStatus)
    Else
        If optErrorsOnly.Value = True Or optErrorsNSkips.Value = True Then Exit Sub
        ts.WriteLine ("Item #" & Itemrec.Number & " [" & ClipNull(Itemrec.Name) & "] -- Update Successful.")
    End If
End If
End Sub
Private Sub ImportRooms()
On Error GoTo error:
Dim nStatus As Integer, recnum As Long, x As Long
Dim ExistingRecord As Boolean

stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_MP

recnum = 0
If tabRooms.RecordCount = 0 Then Exit Sub
tabRooms.MoveFirst
ts.WriteBlankLines (2)
 
tabRooms.Index = "idxRooms"
If chkRoomsAll.Value = 0 Then
    tabRooms.Seek "=", Val(txtRoomsMap.Text), Val(txtRoomsFrom.Text)
    If tabRooms.NoMatch Then tabRooms.MoveFirst
End If
 
Do While tabRooms.EOF = False And bStopImport = False
    
    recnum = recnum + 1
    stsStatusBar.Panels(2).Text = recnum
    IncreaseProgressBar
    
    RoomKeyStruct.MapNum = tabRooms.Fields("Map Number")
    RoomKeyStruct.RoomNum = tabRooms.Fields("Room Number")

    If chkRoomsAll.Value = 0 Then
        If Not Val(txtRoomsMap.Text) = RoomKeyStruct.MapNum Or RoomKeyStruct.RoomNum < Val(txtRoomsFrom.Text) Then
            If optErrorsOnly.Value = True Then GoTo SkipRecord
            ts.WriteLine ("Room " & tabRooms.Fields("Map Number") & "/" _
                & tabRooms.Fields("Room Number") & " [" & ClipNull(tabRooms.Fields("Name")) _
                & "] -- Skipped, out of import range.")
            GoTo SkipRecord:
        ElseIf RoomKeyStruct.RoomNum > Val(txtRoomsTo.Text) Then
            Exit Sub
        End If
    End If
    
    ExistingRecord = True
    nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), RoomKeyStruct, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nStatus = 4 Then
            ExistingRecord = False
            For x = LBound(Roomdatabuf.buf()) To UBound(Roomdatabuf.buf())
                Roomdatabuf.buf(x) = &H0
            Next x
            'Call RoomRowToStruct(Roomdatabuf.buf)
        Else
            ts.WriteLine ("Room " & tabRooms.Fields("Map Number") & "/" _
                & tabRooms.Fields("Room Number") & " [" & ClipNull(tabRooms.Fields("Name")) _
                & "] -- Error " & nStatus & " determining if record exists.")
            GoTo SkipRecord:
        End If
    Else
        If optSkip.Value = True Then
            If optErrorsOnly.Value = True Then GoTo SkipRecord
            ts.WriteLine ("Room " & tabRooms.Fields("Map Number") & "/" _
                & tabRooms.Fields("Room Number") & " [" & ClipNull(tabRooms.Fields("Name")) _
                & "] -- Existing record. Skipped due to setting.")
            GoTo SkipRecord:
        End If
    End If
    
    Call RoomRowToStruct(Roomdatabuf.buf)

    Roomrec.MapNumber = tabRooms.Fields("Map Number")
    Roomrec.RoomNumber = tabRooms.Fields("Room Number")
    
    If chkNotItems.Value = 1 Then GoTo not_items:
    
    For x = 0 To 16
        Roomrec.RoomItems(x) = tabRooms.Fields("Room Item " & x)
        Roomrec.RoomItemQty(x) = tabRooms.Fields("Room Item " & x & " QTY")
        Roomrec.RoomItemUses(x) = tabRooms.Fields("Room Item " & x & " USES")
    Next

    For x = 0 To 14
        Roomrec.InvisItems(x) = tabRooms.Fields("Hidden Item " & x)
        Roomrec.InvisItemQty(x) = tabRooms.Fields("Hidden Item " & x & " QTY")
        Roomrec.InvisItemUses(x) = tabRooms.Fields("Hidden Item " & x & " USES")
    Next

not_items:
    
    If chkOnlyItems.Value = 1 Then GoTo only_items:
    
    Roomrec.Name = Trim(tabRooms.Fields("Name"))
    Roomrec.AnsiMap = Trim(tabRooms.Fields("AnsiMap"))
    Roomrec.Type = tabRooms.Fields("Type")
    Roomrec.ShopNum = tabRooms.Fields("Shop Number")
    Roomrec.GangHouseNumber = tabRooms.Fields("Gang House Number")
    Roomrec.MinIndex = tabRooms.Fields("Min Index")
    Roomrec.MaxIndex = tabRooms.Fields("Max Index")
    Roomrec.PermNPC = tabRooms.Fields("Perm NPC")
    Roomrec.Light = tabRooms.Fields("Light")
    Roomrec.MonsterType = tabRooms.Fields("Mon Type")
    Roomrec.MaxRegen = tabRooms.Fields("Max Regen")
    Roomrec.DeathRoom = tabRooms.Fields("Death Room")
    Roomrec.CmdText = tabRooms.Fields("Command Text")
    Roomrec.Delay = tabRooms.Fields("Delay")
    Roomrec.MaxArea = tabRooms.Fields("Max Area")
    Roomrec.ControlRoom = tabRooms.Fields("Control Room")
    Roomrec.Runic = tabRooms.Fields("Runic")
    Roomrec.Platinum = tabRooms.Fields("Platinum")
    Roomrec.Gold = tabRooms.Fields("Gold")
    Roomrec.Silver = tabRooms.Fields("Silver")
    Roomrec.Copper = tabRooms.Fields("Copper")
    Roomrec.InvisRunic = tabRooms.Fields("InvisRunic")
    Roomrec.InvisPlatinum = tabRooms.Fields("InvisPlatinum")
    Roomrec.InvisGold = tabRooms.Fields("InvisGold")
    Roomrec.InvisSilver = tabRooms.Fields("InvisSilver")
    Roomrec.InvisCopper = tabRooms.Fields("InvisCopper")
    Roomrec.Spell = tabRooms.Fields("Spell")
    Roomrec.ExitRoom = tabRooms.Fields("Exit Room")
    Roomrec.Attributes = tabRooms.Fields("Attributes")

    For x = 0 To 6
         Roomrec.Desc(x) = Trim(tabRooms.Fields("Desc " & x))
    Next
    
    For x = 0 To 9
        Roomrec.RoomExit(x) = tabRooms.Fields("Exit " & x)
        Roomrec.RoomType(x) = tabRooms.Fields("Type " & x)
        Roomrec.Para1(x) = tabRooms.Fields("Para1 " & x)
        Roomrec.Para2(x) = tabRooms.Fields("Para2 " & x)
        Roomrec.Para3(x) = tabRooms.Fields("Para3 " & x)
        Roomrec.Para4(x) = tabRooms.Fields("Para4 " & x)
        Roomrec.PlacedItems(x) = tabRooms.Fields("Placed Item " & x)
    Next
    
    'For x = 0 To 14
    '    Roomrec.CurrentRoomMon(x) = tabRooms.Fields("CurrentRoomMon " & x)
    'Next x
    
only_items:

    If ExistingRecord = True Then
        iUpdateRoom
    Else
        RoomStructToRow Roomdatabuf.buf

        If bPreview Then
            ts.WriteLine ("Room " & tabRooms.Fields("Map Number") & "/" _
                & tabRooms.Fields("Room Number") & " [" & ClipNull(tabRooms.Fields("Name")) _
                & "] -- Non-Existing Record, would be inserted.")
        Else
            nStatus = BTRCALL(BINSERT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
            If Not nStatus = 0 Then
                ts.WriteLine ("Room " & tabRooms.Fields("Map Number") & "/" _
                    & tabRooms.Fields("Room Number") & " [" & ClipNull(tabRooms.Fields("Name")) _
                    & "] -- Insert Error: " & nStatus)
            Else
                If optErrorsNSkips.Value = True Then GoTo SkipRecord:
                If optErrorsOnly.Value = True Then GoTo SkipRecord:
                ts.WriteLine ("Room " & tabRooms.Fields("Map Number") & "/" _
                    & tabRooms.Fields("Room Number") & " [" & ClipNull(tabRooms.Fields("Name")) _
                    & "] -- Insert Successful.")
            End If
        End If
    End If
    
SkipRecord:
    tabRooms.MoveNext
    If Not bUseCPU Then DoEvents
Loop

Exit Sub
error:
If CheckError = True Then Resume Next
End Sub
Private Sub iUpdateRoom()
Dim nStatus As Integer

RoomStructToRow Roomdatabuf.buf

If bPreview Then
    ts.WriteLine ("Room " & Roomrec.MapNumber & "/" & Roomrec.RoomNumber & " [" & ClipNull(Roomrec.Name) & "] -- Existing Record, would be updated.")
Else
    nStatus = BTRCALL(BUPDATE, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        ts.WriteLine ("Room " & Roomrec.MapNumber & "/" & Roomrec.RoomNumber & " [" & ClipNull(Roomrec.Name) & "] -- Update Error: " & nStatus)
    Else
        If optErrorsOnly.Value = True Or optErrorsNSkips.Value = True Then Exit Sub
        ts.WriteLine ("Room " & Roomrec.MapNumber & "/" & Roomrec.RoomNumber & " [" & ClipNull(Roomrec.Name) & "] -- Update Successful.")
    End If
End If
End Sub
Private Sub ImportSpells()
On Error GoTo error:
Dim nStatus As Integer, recnum As Long, x As Long
Dim ExistingRecord As Boolean

stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_SPELS

If tabSpells.RecordCount = 0 Then Exit Sub
tabSpells.MoveFirst
ts.WriteBlankLines (2)

tabSpells.Index = "pkSpells"
If chkSpellsAll.Value = 0 Then
    tabSpells.Seek "=", Val(txtSpellsFrom.Text)
    If tabSpells.NoMatch Then tabSpells.MoveFirst
End If

Do While tabSpells.EOF = False And bStopImport = False
    
    recnum = tabSpells.Fields("Number")
    stsStatusBar.Panels(2).Text = recnum
    IncreaseProgressBar
    
    x = tabSpells.Fields("Number")
    
    If chkSpellsAll.Value = 0 Then
        If x < Val(txtSpellsFrom.Text) Then
            If optErrorsOnly.Value = True Then GoTo SkipRecord
            ts.WriteLine ("Spell #" & tabSpells.Fields("Number") & " [" & ClipNull(tabSpells.Fields("Name")) & "] -- Skipped, out of import range.")
            GoTo SkipRecord:
        ElseIf x > Val(txtSpellsTo.Text) Then
            Exit Sub
        End If
    End If
    
    ExistingRecord = True
    nStatus = BTRCALL(BGETEQUAL, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), x, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nStatus = 4 Then
            ExistingRecord = False
        Else
            ts.WriteLine ("Spell #" & tabSpells.Fields("Number") & " [" & ClipNull(tabSpells.Fields("Name")) & "] -- Error " & nStatus & " determining if record exists.")
            GoTo SkipRecord:
        End If
    Else
        If optSkip.Value = True Then
            If optErrorsOnly.Value = True Then GoTo SkipRecord
            ts.WriteLine ("Spell #" & tabSpells.Fields("Number") & " [" & ClipNull(tabSpells.Fields("Name")) & "] -- Existing record. Skipped due to setting.")
            GoTo SkipRecord:
        End If
    End If
    
    SpellRowToStruct Spelldatabuf.buf
    
        Spellrec.Number = tabSpells.Fields("Number")
        Spellrec.Name = Trim(tabSpells.Fields("Name"))
        Spellrec.ShortName = Trim(tabSpells.Fields("Short Name"))
        Spellrec.Level = tabSpells.Fields("Level")
        Spellrec.DescA = Trim(tabSpells.Fields("Desc 1"))
        Spellrec.DescB = Trim(tabSpells.Fields("Desc 2"))
        Spellrec.CastMsgA = tabSpells.Fields("Cast MSG A")
        Spellrec.CastMsgB = tabSpells.Fields("Cast MSG B")
        Spellrec.MsgStyle = tabSpells.Fields("MSG Style")
        Spellrec.Energy = tabSpells.Fields("Energy")
        Spellrec.Mana = tabSpells.Fields("Mana")
        Spellrec.Min = tabSpells.Fields("Min")
        Spellrec.Max = tabSpells.Fields("Max")
        Spellrec.SpellType = tabSpells.Fields("Spell Type")
        Spellrec.TypeOfResists = tabSpells.Fields("Type of Resists")
        Spellrec.Difficulty = tabSpells.Fields("Difficulty")
        Spellrec.Target = tabSpells.Fields("Target")
        Spellrec.duration = tabSpells.Fields("Duration")
        Spellrec.TypeOfAttack = tabSpells.Fields("Attack Type")
        Spellrec.ResistAbility = tabSpells.Fields("Resist Ability")
        Spellrec.MageryA = tabSpells.Fields("Magery A")
        Spellrec.MageryB = tabSpells.Fields("Magery B")
        Spellrec.LevelCap = tabSpells.Fields("Level Cap")
        Spellrec.LVLSMaxIncr = tabSpells.Fields("LVLS Max Increase")
        Spellrec.MaxIncrease = tabSpells.Fields("Max Increase")
        Spellrec.LVLSMinIncr = tabSpells.Fields("LVLS Min Increase")
        Spellrec.MinIncrease = tabSpells.Fields("Min Increase")
        Spellrec.LVLSDurIncr = tabSpells.Fields("LVLS Dur Increase")
        Spellrec.DurIncrease = tabSpells.Fields("Dur Increase")
        Spellrec.UNDEFINED01 = tabSpells.Fields("UNDEFINED01")
        Spellrec.UNDEFINED02 = tabSpells.Fields("UNDEFINED02")
 
        For x = 0 To 9
            Spellrec.AbilityA(x) = tabSpells.Fields("Ability " & x)
            Spellrec.AbilityB(x) = tabSpells.Fields("Ability Value " & x)
        Next
        
        If ExistingRecord = True Then
            iUpdateSpell
        Else

            SpellStructToRow Spelldatabuf.buf
            
            If bPreview Then
                ts.WriteLine ("Spell #" & tabSpells.Fields("Number") & " [" & ClipNull(tabSpells.Fields("Name")) & "] -- Non-Existing Record, would be Inserted.")
            Else
                nStatus = BTRCALL(BINSERT, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
                If Not nStatus = 0 Then
                    ts.WriteLine ("Spell #" & tabSpells.Fields("Number") & " [" & ClipNull(tabSpells.Fields("Name")) & "] -- Insert Error: " & nStatus)
                Else
                    If optErrorsNSkips.Value = True Then GoTo SkipRecord:
                    If optErrorsOnly.Value = True Then GoTo SkipRecord:
                    ts.WriteLine ("Spell #" & tabSpells.Fields("Number") & " [" & ClipNull(tabSpells.Fields("Name")) & "] -- Insert Successful.")
                End If
            End If
        End If
        
SkipRecord:
        tabSpells.MoveNext
        If Not bUseCPU Then DoEvents
Loop

Exit Sub
error:
If CheckError = True Then Resume Next
End Sub
Private Sub iUpdateSpell()
Dim nStatus As Integer

SpellStructToRow Spelldatabuf.buf

If bPreview Then
    ts.WriteLine ("Spell #" & Spellrec.Number & " [" & ClipNull(Spellrec.Name) & "] -- Existing Record, would be updated.")
Else
    nStatus = BTRCALL(BUPDATE, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        ts.WriteLine ("Spell #" & Spellrec.Number & " [" & ClipNull(Spellrec.Name) & "] -- Update Error: " & nStatus)
    Else
        If optErrorsOnly.Value = True Or optErrorsNSkips.Value = True Then Exit Sub
        ts.WriteLine ("Spell #" & Spellrec.Number & " [" & ClipNull(Spellrec.Name) & "] -- Update Successful.")
    End If
End If
End Sub

Private Sub ImportActions()
On Error GoTo error:
Dim nStatus As Integer, recnum As Long, x As Long, ActionName As String * 30
Dim ExistingRecord As Boolean

stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_ACTS

If tabActions.RecordCount = 0 Then Exit Sub
tabActions.MoveFirst
ts.WriteBlankLines (2)

tabActions.Index = "pkActions"
recnum = 0

Do While tabActions.EOF = False And bStopImport = False
    
    recnum = recnum + 1
    stsStatusBar.Panels(2).Text = recnum
    IncreaseProgressBar
    
    ActionName = tabActions.Fields("Action") & Chr(0)
    
    'tabActions.Index = "pkActions"
    
    ExistingRecord = True
    nStatus = BTRCALL(BGETEQUAL, ActionPosBlock, ActionDatabuf, Len(ActionDatabuf), ByVal ActionName, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nStatus = 4 Then
            ExistingRecord = False
        Else
            ts.WriteLine ("Action: " & RemoveCharacter(tabActions.Fields("Action"), " ") & "-- Error " & nStatus & " determining if record exists.")
            GoTo SkipRecord:
        End If
    Else
        If optSkip.Value = True Then
            If optErrorsOnly.Value = True Then GoTo SkipRecord
            ts.WriteLine ("Action: " & RemoveCharacter(tabActions.Fields("Action"), " ") & "-- Existing record. Skipped due to setting.")
            GoTo SkipRecord:
        End If
    End If
    
    ActionRowToStruct ActionDatabuf.buf
    
        Actionrec.Name = Trim(tabActions.Fields("Action"))
        Actionrec.SingleToUser = Trim(tabActions.Fields("Single to User"))
        Actionrec.SingleToRoom = Trim(tabActions.Fields("Single to Room"))
        Actionrec.UserToUser = Trim(tabActions.Fields("User to User"))
        Actionrec.UserToOtherUser = Trim(tabActions.Fields("User to Other User"))
        Actionrec.UserToRoom = Trim(tabActions.Fields("User to Room"))
        Actionrec.MonsterToUser = Trim(tabActions.Fields("Monster to User"))
        Actionrec.MonsterToRoom = Trim(tabActions.Fields("Monster to Room"))
        Actionrec.InventoryToUser = Trim(tabActions.Fields("Inventory to User"))
        Actionrec.InventoryToRoom = Trim(tabActions.Fields("Inventory to Room"))
        Actionrec.FloorItemToUser = Trim(tabActions.Fields("Floor Item to User"))
        Actionrec.FloorItemToRoom = Trim(tabActions.Fields("Floor Item to Room"))

        If ExistingRecord = True Then
            iUpdateAction
        Else

            ActionStructToRow ActionDatabuf.buf
            
            If bPreview Then
                ts.WriteLine ("Action: " & RemoveCharacter(tabActions.Fields("Action"), " ") & "-- Non-Existing Record, would be inserted.")
            Else
                nStatus = BTRCALL(BINSERT, ActionPosBlock, ActionDatabuf, Len(ActionDatabuf), ByVal ActionKeyBuffer, KEY_BUF_LEN, 0)
                If Not nStatus = 0 Then
                    ts.WriteLine ("Action: " & RemoveCharacter(tabActions.Fields("Action"), " ") & "-- Insert Error: " & nStatus)
                Else
                    If optErrorsNSkips.Value = True Then GoTo SkipRecord:
                    If optErrorsOnly.Value = True Then GoTo SkipRecord:
                    ts.WriteLine ("Action: " & RemoveCharacter(tabActions.Fields("Action"), " ") & "-- Insert Successful.")
                End If
            End If
        End If
        
SkipRecord:
        tabActions.MoveNext
        If Not bUseCPU Then DoEvents
Loop

Exit Sub
error:
If CheckError = True Then Resume Next
End Sub
Private Sub iUpdateAction()
Dim nStatus As Integer

ActionStructToRow ActionDatabuf.buf

If bPreview Then
    ts.WriteLine ("Action: " & RemoveCharacter(Actionrec.Name, " ") & "-- Existing Record, would be updated.")
Else
    nStatus = BTRCALL(BUPDATE, ActionPosBlock, ActionDatabuf, Len(ActionDatabuf), ByVal ActionKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        ts.WriteLine ("Action: " & RemoveCharacter(Actionrec.Name, " ") & "-- Update Error: " & nStatus)
    Else
        If optErrorsOnly.Value = True Or optErrorsNSkips.Value = True Then Exit Sub
        ts.WriteLine ("Action: " & RemoveCharacter(Actionrec.Name, " ") & "-- Update Successful.")
    End If
End If
End Sub
Private Sub ImportClasses()
On Error GoTo error:
Dim nStatus As Integer, recnum As Long, x As Long
Dim ExistingRecord As Boolean

stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_CLASS
    
If tabClasses.RecordCount = 0 Then Exit Sub
tabClasses.MoveFirst
ts.WriteBlankLines (2)

tabClasses.Index = "pkClasses"
If chkClassesAll.Value = 0 Then
    tabClasses.Seek "=", Val(txtClassesFrom.Text)
    If tabClasses.NoMatch Then tabClasses.MoveFirst
End If

Do While tabClasses.EOF = False And bStopImport = False
    
    recnum = tabClasses.Fields("Number")
    stsStatusBar.Panels(2).Text = recnum
    IncreaseProgressBar
    
    x = tabClasses.Fields("Number")
    
    If chkClassesAll.Value = 0 Then
        If x < Val(txtClassesFrom.Text) Then
            If optErrorsOnly.Value = True Then GoTo SkipRecord
            ts.WriteLine ("Class #" & tabClasses.Fields("Number") & " [" & ClipNull(tabClasses.Fields("Name")) & "] -- Skipped, out of import range.")
            GoTo SkipRecord:
        ElseIf x > Val(txtClassesTo.Text) Then
            Exit Sub
        End If
    End If
    
    ExistingRecord = True
    nStatus = BTRCALL(BGETEQUAL, ClassPosBlock, Classdatabuf, Len(Classdatabuf), x, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nStatus = 4 Then
            ExistingRecord = False
        Else
            ts.WriteLine ("Class #" & tabClasses.Fields("Number") & " [" & ClipNull(tabClasses.Fields("Name")) & "] -- Error " & nStatus & " determining if record exists.")
            GoTo SkipRecord:
        End If
    Else
        If optSkip.Value = True Then
            If optErrorsOnly.Value = True Then GoTo SkipRecord
            ts.WriteLine ("Class #" & tabClasses.Fields("Number") & " [" & ClipNull(tabClasses.Fields("Name")) & "] -- Existing record. Skipped due to setting.")
            GoTo SkipRecord:
        End If
    End If
    
    ClassRowToStruct Classdatabuf.buf
    
        Classrec.Number = tabClasses.Fields("Number")
        Classrec.Name = Trim(tabClasses.Fields("Name"))
        Classrec.MinHp = tabClasses.Fields("Min HP")
        Classrec.MaxHP = tabClasses.Fields("Max HP")
        Classrec.Exp = tabClasses.Fields("EXP %")
        Classrec.MagicType = tabClasses.Fields("Magic Type")
        Classrec.MagicLvL = tabClasses.Fields("Magic LVL")
        Classrec.Weapon = tabClasses.Fields("Weapon")
        Classrec.Armour = tabClasses.Fields("Armour")
        Classrec.Combat = tabClasses.Fields("Combat")
        Classrec.TitleText = tabClasses.Fields("Title Text")

        For x = 0 To 9
            Classrec.AbilityA(x) = tabClasses.Fields("Ability " & x)
            Classrec.AbilityB(x) = tabClasses.Fields("Ability Value " & x)
        Next

        If ExistingRecord = True Then
            iUpdateClass
        Else

            ClassStructToRow Classdatabuf.buf
            
            If bPreview Then
                ts.WriteLine ("Class #" & tabClasses.Fields("Number") & " [" & ClipNull(tabClasses.Fields("Name")) & "] -- Non-Existing Record, would be inserted.")
            Else
                nStatus = BTRCALL(BINSERT, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
                If Not nStatus = 0 Then
                    ts.WriteLine ("Class #" & tabClasses.Fields("Number") & " [" & ClipNull(tabClasses.Fields("Name")) & "] -- Insert Error: " & nStatus)
                Else
                    If optErrorsNSkips.Value = True Then GoTo SkipRecord:
                    If optErrorsOnly.Value = True Then GoTo SkipRecord:
                    ts.WriteLine ("Class #" & tabClasses.Fields("Number") & " [" & ClipNull(tabClasses.Fields("Name")) & "] -- Insert Successful.")
                End If
            End If
        End If
        
SkipRecord:
        tabClasses.MoveNext
        If Not bUseCPU Then DoEvents
Loop

Call LoadClassArray

Exit Sub
error:
If CheckError = True Then Resume Next
End Sub
Private Sub iUpdateClass()
Dim nStatus As Integer

ClassStructToRow Classdatabuf.buf

If bPreview Then
    ts.WriteLine ("Class #" & Classrec.Number & " [" & ClipNull(Classrec.Name) & "] -- Existing Record, would be updated.")
Else
    nStatus = BTRCALL(BUPDATE, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        ts.WriteLine ("Class #" & Classrec.Number & " [" & ClipNull(Classrec.Name) & "] -- Update Error: " & nStatus)
    Else
        If optErrorsOnly.Value = True Or optErrorsNSkips.Value = True Then Exit Sub
        ts.WriteLine ("Class #" & Classrec.Number & " [" & ClipNull(Classrec.Name) & "] -- Update Successful.")
    End If
End If
End Sub
Private Sub ImportRaces()
On Error GoTo error:
Dim nStatus As Integer, recnum As Long, x As Long
Dim ExistingRecord As Boolean

stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_RACE

If tabRaces.RecordCount = 0 Then Exit Sub
tabRaces.MoveFirst
ts.WriteBlankLines (2)

tabRaces.Index = "pkRaces"
If chkRacesAll.Value = 0 Then
    tabRaces.Seek "=", Val(txtRacesFrom.Text)
    If tabRaces.NoMatch Then tabRaces.MoveFirst
End If

Do While tabRaces.EOF = False And bStopImport = False
    
    recnum = tabRaces.Fields("Number")
    stsStatusBar.Panels(2).Text = recnum
    IncreaseProgressBar
    
    x = tabRaces.Fields("Number")

    If chkRacesAll.Value = 0 Then
        If x < Val(txtRacesFrom.Text) Then
            If optErrorsOnly.Value = True Then GoTo SkipRecord
            ts.WriteLine ("Race #" & tabRaces.Fields("Number") & " [" & ClipNull(tabRaces.Fields("Name")) & "] -- Skipped, out of import range.")
            GoTo SkipRecord:
        ElseIf x > Val(txtRacesTo.Text) Then
            Exit Sub
        End If
    End If
    
    ExistingRecord = True
    nStatus = BTRCALL(BGETEQUAL, RacePosBlock, Racedatabuf, Len(Racedatabuf), x, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nStatus = 4 Then
            ExistingRecord = False
        Else
            ts.WriteLine ("Race #" & tabRaces.Fields("Number") & " [" & ClipNull(tabRaces.Fields("Name")) & "] -- Error " & nStatus & " determining if record exists.")
            GoTo SkipRecord:
        End If
    Else
        If optSkip.Value = True Then
            If optErrorsOnly.Value = True Then GoTo SkipRecord
            ts.WriteLine ("Race #" & tabRaces.Fields("Number") & " [" & ClipNull(tabRaces.Fields("Name")) & "] -- Existing record. Skipped due to setting.")
            GoTo SkipRecord:
        End If
    End If
    
    RaceRowToStruct Racedatabuf.buf
    
        Racerec.Number = tabRaces.Fields("Number")
        Racerec.Name = Trim(tabRaces.Fields("Name"))
        Racerec.MinInt = tabRaces.Fields("Min INT")
        Racerec.MinWil = tabRaces.Fields("Min WIL")
        Racerec.MinStr = tabRaces.Fields("Min STR")
        Racerec.MinHea = tabRaces.Fields("Min HEA")
        Racerec.MinAgl = tabRaces.Fields("Min AGL")
        Racerec.MinChm = tabRaces.Fields("Min CHM")
        Racerec.MaxInt = tabRaces.Fields("Max INT")
        Racerec.MaxWil = tabRaces.Fields("Max WIL")
        Racerec.MaxStr = tabRaces.Fields("Max STR")
        Racerec.MaxHea = tabRaces.Fields("Max HEA")
        Racerec.MaxAgl = tabRaces.Fields("Max AGL")
        Racerec.MaxChm = tabRaces.Fields("Max CHM")
        Racerec.HPBonus = tabRaces.Fields("HP Bonus")
        Racerec.CP = tabRaces.Fields("CP")
        Racerec.ExpChart = tabRaces.Fields("EXP %")
 
        For x = 0 To 9
            Racerec.AbilityA(x) = tabRaces.Fields("Ability " & x)
            Racerec.AbilityB(x) = tabRaces.Fields("Ability Value " & x)
        Next

        If ExistingRecord = True Then
            iUpdateRace
        Else

            RaceStructToRow Racedatabuf.buf
            
            If bPreview Then
                ts.WriteLine ("Race #" & tabRaces.Fields("Number") & " [" & ClipNull(tabRaces.Fields("Name")) & "] -- Non-Existing Record, would be inserted.")
            Else
                nStatus = BTRCALL(BINSERT, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
                If Not nStatus = 0 Then
                    ts.WriteLine ("Race #" & tabRaces.Fields("Number") & " [" & ClipNull(tabRaces.Fields("Name")) & "] -- Insert Error: " & nStatus)
                Else
                    If optErrorsNSkips.Value = True Then GoTo SkipRecord:
                    If optErrorsOnly.Value = True Then GoTo SkipRecord:
                    ts.WriteLine ("Race #" & tabRaces.Fields("Number") & " [" & ClipNull(tabRaces.Fields("Name")) & "] -- Insert Successful.")
                End If
            End If
        End If
        
SkipRecord:
        tabRaces.MoveNext
        If Not bUseCPU Then DoEvents
Loop

Call LoadRaceArray

Exit Sub
error:
If CheckError = True Then Resume Next
End Sub
Private Sub iUpdateRace()
Dim nStatus As Integer

RaceStructToRow Racedatabuf.buf

If bPreview Then
    ts.WriteLine ("Race #" & Racerec.Number & " [" & ClipNull(Racerec.Name) & "] -- Existing Record, would be updated.")
Else
    nStatus = BTRCALL(BUPDATE, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        ts.WriteLine ("Race #" & Racerec.Number & " [" & ClipNull(Racerec.Name) & "] -- Update Error: " & nStatus)
    Else
        If optErrorsOnly.Value = True Or optErrorsNSkips.Value = True Then Exit Sub
        ts.WriteLine ("Race #" & Racerec.Number & " [" & ClipNull(Racerec.Name) & "] -- Update Successful.")
    End If
End If
End Sub
Private Sub ImportShops()
On Error GoTo error:
Dim nStatus As Integer, recnum As Long, x As Long
Dim ExistingRecord As Boolean

stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_SHOPS

If tabShops.RecordCount = 0 Then Exit Sub
tabShops.MoveFirst
ts.WriteBlankLines (2)

tabShops.Index = "pkShops"
If chkShopsAll.Value = 0 Then
    tabShops.Seek "=", Val(txtShopsFrom.Text)
    If tabShops.NoMatch Then tabShops.MoveFirst
End If

Do While tabShops.EOF = False And bStopImport = False
    
    recnum = tabShops.Fields("Number")
    stsStatusBar.Panels(2).Text = recnum
    IncreaseProgressBar
    
    x = tabShops.Fields("Number")
    
    If chkShopsAll.Value = 0 Then
        If x < Val(txtShopsFrom.Text) Then
            If optErrorsOnly.Value = True Then GoTo SkipRecord
            ts.WriteLine ("Shop #" & tabShops.Fields("Number") & " [" & ClipNull(tabShops.Fields("Name")) & "] -- Skipped, out of import range.")
            GoTo SkipRecord:
        ElseIf x > Val(txtShopsTo.Text) Then
            Exit Sub
        End If
    End If
    
    ExistingRecord = True
    nStatus = BTRCALL(BGETEQUAL, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), x, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nStatus = 4 Then
            ExistingRecord = False
        Else
            ts.WriteLine ("Shop #" & tabShops.Fields("Number") & " [" & ClipNull(tabShops.Fields("Name")) & "] -- Error " & nStatus & " determining if record exists.")
            GoTo SkipRecord:
        End If
    Else
        If optSkip.Value = True Then
            If optErrorsOnly.Value = True Then GoTo SkipRecord
            ts.WriteLine ("Shop #" & tabShops.Fields("Number") & " [" & ClipNull(tabShops.Fields("Name")) & "] -- Existing record. Skipped due to setting.")
            GoTo SkipRecord:
        End If
    End If
    
    ShopRowToStruct Shopdatabuf.buf
    
        Shoprec.Number = tabShops.Fields("Number")
        Shoprec.Name = Trim(tabShops.Fields("Name"))
        Shoprec.ShopDescriptionA = Trim(tabShops.Fields("Desc A"))
        Shoprec.ShopDescriptionB = Trim(tabShops.Fields("Desc B"))
        Shoprec.ShopDescriptionC = Trim(tabShops.Fields("Desc C"))
        Shoprec.ShopType = tabShops.Fields("Type")
        Shoprec.ShopMinLvL = tabShops.Fields("Min Lvl")
        Shoprec.ShopMaxLvl = tabShops.Fields("Max Lvl")
        Shoprec.ShopMarkUp = tabShops.Fields("MarkUp")
        Shoprec.ShopClassLimit = tabShops.Fields("Class Limit")

        For x = 0 To 19
            Shoprec.ShopItemNumber(x) = tabShops.Fields("Item " & x)
            Shoprec.ShopMax(x) = tabShops.Fields("Max " & x)
            Shoprec.ShopNow(x) = tabShops.Fields("Normal " & x)
            Shoprec.ShopRgnTime(x) = tabShops.Fields("Regen Time " & x)
            Shoprec.ShopRgnNumber(x) = tabShops.Fields("Regen Number" & x)
            Shoprec.ShopRgnPercentage(x) = tabShops.Fields("Regen %" & x)
        Next

        If ExistingRecord = True Then
            iUpdateShop
        Else

            ShopStructToRow Shopdatabuf.buf
            
            If bPreview Then
                ts.WriteLine ("Shop #" & tabShops.Fields("Number") & " [" & ClipNull(tabShops.Fields("Name")) & "] -- Non-Existing Record, would be inserted.")
            Else
                nStatus = BTRCALL(BINSERT, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
                If Not nStatus = 0 Then
                    ts.WriteLine ("Shop #" & tabShops.Fields("Number") & " [" & ClipNull(tabShops.Fields("Name")) & "] -- Insert Error: " & nStatus)
                Else
                    If optErrorsNSkips.Value = True Then GoTo SkipRecord:
                    If optErrorsOnly.Value = True Then GoTo SkipRecord:
                    ts.WriteLine ("Shop #" & tabShops.Fields("Number") & " [" & ClipNull(tabShops.Fields("Name")) & "] -- Insert Successful.")
                End If
            End If
        End If
        
SkipRecord:
        tabShops.MoveNext
        If Not bUseCPU Then DoEvents
Loop

Exit Sub
error:
If CheckError = True Then Resume Next
End Sub
Private Sub iUpdateShop()
Dim nStatus As Integer

ShopStructToRow Shopdatabuf.buf

If bPreview Then
    ts.WriteLine ("Shop #" & Shoprec.Number & " [" & ClipNull(Shoprec.Name) & "] -- Existing Record, would be updated.")
Else
    nStatus = BTRCALL(BUPDATE, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        ts.WriteLine ("Shop #" & Shoprec.Number & " [" & ClipNull(Shoprec.Name) & "] -- Update Error: " & nStatus)
    Else
        If optErrorsOnly.Value = True Or optErrorsNSkips.Value = True Then Exit Sub
        ts.WriteLine ("Shop #" & Shoprec.Number & " [" & ClipNull(Shoprec.Name) & "] -- Update Successful.")
    End If
End If
End Sub
Private Function TestMonsterFields() As Boolean
On Error GoTo error:
'Dim adoConnect As Database, tabMonsters As Recordset
Dim nTemp As Integer, fldTemp As field

'this function is just to test if the "Exp Multiplier" field exists. if not, it errors out

TestMonsterFields = False

'Set adoConnect = OpenDatabase(sDataSource)
'Set tabMonsters = adoConnect.OpenRecordset("Monsters")

nTemp = 0
For Each fldTemp In tabMonsters.Fields()
    If fldTemp.Name = "Exp Multiplier" Then nTemp = 1
Next

If nTemp = 1 Then TestMonsterFields = True

'tabMonsters.Close
'adoConnect.Close
'
'Set tabMonsters = Nothing
'Set adoConnect = Nothing

Exit Function
error:
End Function
Private Sub ImportMonsters()
On Error GoTo error:
Dim nStatus As Integer, recnum As Long, x As Long, test As Boolean, nYesNo As Integer, ExpMulti1 As Boolean
Dim ExistingRecord As Boolean

stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_KNMSR

If tabMonsters.RecordCount = 0 Then Exit Sub
tabMonsters.MoveFirst
ts.WriteBlankLines (2)

ExpMulti1 = False
If eDatFileVersion >= v111j Then
    test = TestMonsterFields
    If test = False Then
        nYesNo = MsgBox("Exported monster table does not contain the 'Exp Multiplier' Field." & vbCrLf _
            & "Import monsters anyway and set the 'Exp Multiplier' Field to 1 for each monster?" & vbCrLf & vbCrLf _
            & "Any monster with experience greater than 2,147,483,646 will be set to that." & vbCrLf & vbCrLf _
            & "NOTE: This can be disasterous if importing original-game monsters as their stats prior to v1.11j were very different.", vbQuestion + vbYesNo)
        If nYesNo = vbYes Then
            ExpMulti1 = True
        Else
            Exit Sub
        End If
    End If
End If

tabMonsters.Index = "pkMonsters"
If chkMonstersAll.Value = 0 Then
    tabMonsters.Seek "=", Val(txtMonstersFrom.Text)
    If tabMonsters.NoMatch Then tabMonsters.MoveFirst
End If

Do While tabMonsters.EOF = False And bStopImport = False
    
    recnum = tabMonsters.Fields("Number")
    stsStatusBar.Panels(2).Text = recnum
    IncreaseProgressBar
    
    x = tabMonsters.Fields("Number")

    If chkMonstersAll.Value = 0 Then
        If x < Val(txtMonstersFrom.Text) Then
            If optErrorsOnly.Value = True Then GoTo SkipRecord
            ts.WriteLine ("Monster #" & tabMonsters.Fields("Number") & " [" & ClipNull(tabMonsters.Fields("Name")) & "] -- Skipped, out of import range.")
            GoTo SkipRecord:
        ElseIf x > Val(txtMonstersTo.Text) Then
            Exit Sub
        End If
    End If
    
    ExistingRecord = True
    nStatus = BTRCALL(BGETEQUAL, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), x, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nStatus = 4 Then
            ExistingRecord = False
        Else
            ts.WriteLine ("Monster #" & tabMonsters.Fields("Number") & " [" & ClipNull(tabMonsters.Fields("Name")) & "] -- Error " & nStatus & " determining if record exists.")
            GoTo SkipRecord:
        End If
    Else
        If optSkip.Value = True Then
            If optErrorsOnly.Value = True Then GoTo SkipRecord
            ts.WriteLine ("Monster #" & tabMonsters.Fields("Number") & " [" & ClipNull(tabMonsters.Fields("Name")) & "] -- Existing record. Skipped due to setting.")
            GoTo SkipRecord:
        End If
    End If
    
    MonsterRowToStruct Monsterdatabuf.buf
    
        Monsterrec.Number = tabMonsters.Fields("Number")
        Monsterrec.Name = Trim(tabMonsters.Fields("Name"))
        Monsterrec.Group = tabMonsters.Fields("Group")
        Monsterrec.Index = tabMonsters.Fields("Index")
        Monsterrec.WeaponNumber = tabMonsters.Fields("Weapon Number")
        Monsterrec.AC = tabMonsters.Fields("AC")
        Monsterrec.DR = tabMonsters.Fields("DR")
        Monsterrec.Follow = tabMonsters.Fields("Follow")
        Monsterrec.MR = tabMonsters.Fields("MR")
        
        Monsterrec.Experience = ULong2SLong(tabMonsters.Fields("Experience"))
        If eDatFileVersion >= v111j Then
            If ExpMulti1 = True Then
                Monsterrec.ExpMulti = 1
            Else
                Monsterrec.ExpMulti = ULong2SLong(tabMonsters.Fields("Exp Multiplier"))
            End If
            If CDbl(SLong2ULong(Monsterrec.Experience)) * CDbl(SLong2ULong(Monsterrec.ExpMulti)) > 2147483646 Then
                Monsterrec.Experience = 65538
                Monsterrec.ExpMulti = 32767
            End If
        End If
        
        Monsterrec.Hitpoints = tabMonsters.Fields("Hit Points")
        Monsterrec.Energy = tabMonsters.Fields("Energy")
        Monsterrec.HPRegen = tabMonsters.Fields("HP Regen")
        Monsterrec.GameLimit = tabMonsters.Fields("Game Limit")
        Monsterrec.CharmLvL = tabMonsters.Fields("Charm LvL")
        Monsterrec.CharmRes = tabMonsters.Fields("Charm RES")
        Monsterrec.BSDefence = tabMonsters.Fields("BS Defense")
        Monsterrec.Active = tabMonsters.Fields("Active")
        Monsterrec.Type = tabMonsters.Fields("Type")
        Monsterrec.Undead = tabMonsters.Fields("Undead")
        Monsterrec.Alignment = tabMonsters.Fields("Alignment")
        Monsterrec.RegenTime = tabMonsters.Fields("Regen Time")
        Monsterrec.DateKilled = tabMonsters.Fields("Date Killed")
        Monsterrec.TimeKilled = tabMonsters.Fields("Time Killed")
        Monsterrec.MoveMsg = tabMonsters.Fields("Move Msg")
        Monsterrec.DeathMsg = tabMonsters.Fields("Death Msg")
        Monsterrec.Runic = tabMonsters.Fields("Runic")
        Monsterrec.Platinum = tabMonsters.Fields("Platinum")
        Monsterrec.Gold = tabMonsters.Fields("Gold")
        Monsterrec.Silver = tabMonsters.Fields("Silver")
        Monsterrec.Copper = tabMonsters.Fields("Copper")
        Monsterrec.GreetTxt = tabMonsters.Fields("Greet Txt")
        Monsterrec.DescTxt = tabMonsters.Fields("Desc Txt")
        Monsterrec.TalkTxt = tabMonsters.Fields("Talk Txt")
        Monsterrec.DeathSpellNumber = tabMonsters.Fields("Death Spell")
        Monsterrec.CreateSpellNumber = tabMonsters.Fields("Create Spell")
        Monsterrec.DescLine1 = Trim(tabMonsters.Fields("Desc 1"))
        Monsterrec.DescLine2 = Trim(tabMonsters.Fields("Desc 2"))
        Monsterrec.DescLine3 = Trim(tabMonsters.Fields("Desc 3"))
        Monsterrec.DescLine4 = Trim(tabMonsters.Fields("Desc 4"))
        Monsterrec.Gender = tabMonsters.Fields("Gender")

        For x = 0 To 4
            Monsterrec.AttackType(x) = tabMonsters.Fields("Attack Type " & x)
            Monsterrec.AttackAccuSpell(x) = tabMonsters.Fields("Attack Accu/Spell " & x)
            Monsterrec.AttackPer(x) = tabMonsters.Fields("Attack % " & x)
            Monsterrec.AttackMinHCastPer(x) = tabMonsters.Fields("Attack Min Hit/Cast % " & x)
            Monsterrec.AttackMaxHCastLvl(x) = tabMonsters.Fields("Attack Max Hit/Cast LVL " & x)
            Monsterrec.AttackHitMsg(x) = tabMonsters.Fields("Attack Hit Msg " & x)
            Monsterrec.AttackDodgeMsg(x) = tabMonsters.Fields("Attack Dodge Msg " & x)
            Monsterrec.AttackMissMsg(x) = tabMonsters.Fields("Attack Miss Msg " & x)
            Monsterrec.AttackEnergy(x) = tabMonsters.Fields("Attack Energy " & x)
            Monsterrec.AttackHitSpell(x) = tabMonsters.Fields("Attack Hit Spell " & x)
            Monsterrec.SpellNumber(x) = tabMonsters.Fields("Spell Number " & x)
            Monsterrec.SpellCastPer(x) = tabMonsters.Fields("Spell Cast % " & x)
            Monsterrec.SpellCastLvl(x) = tabMonsters.Fields("Spell Cast LVL " & x)
        Next

        For x = 0 To 9
            Monsterrec.ItemNumber(x) = tabMonsters.Fields("Item Number " & x)
            Monsterrec.ItemUses(x) = tabMonsters.Fields("Item Uses " & x)
            Monsterrec.ItemDropPer(x) = tabMonsters.Fields("Item Drop % " & x)
        Next

        For x = 0 To 9
            Monsterrec.AbilityA(x) = tabMonsters.Fields("Ability " & x)
            Monsterrec.AbilityB(x) = tabMonsters.Fields("Ability Value " & x)
        Next


        If ExistingRecord = True Then
            iUpdateMonster
        Else

            MonsterStructToRow Monsterdatabuf.buf
            
            If bPreview Then
                ts.WriteLine ("Monster #" & tabMonsters.Fields("Number") & " [" & ClipNull(tabMonsters.Fields("Name")) & "] -- Non-Existing Record, would be imported.")
            Else
                nStatus = BTRCALL(BINSERT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
                If Not nStatus = 0 Then
                    ts.WriteLine ("Monster #" & tabMonsters.Fields("Number") & " [" & ClipNull(tabMonsters.Fields("Name")) & "] -- Insert Error: " & nStatus)
                Else
                    If optErrorsNSkips.Value = True Then GoTo SkipRecord:
                    If optErrorsOnly.Value = True Then GoTo SkipRecord:
                    ts.WriteLine ("Monster #" & tabMonsters.Fields("Number") & " [" & ClipNull(tabMonsters.Fields("Name")) & "] -- Insert Successful.")
                End If
            End If
        End If
        
SkipRecord:
        tabMonsters.MoveNext
        If Not bUseCPU Then DoEvents
Loop

Exit Sub
error:
If CheckError = True Then Resume Next
End Sub
Private Sub iUpdateMonster()
Dim nStatus As Integer

MonsterStructToRow Monsterdatabuf.buf

If bPreview Then
    ts.WriteLine ("Monster #" & Monsterrec.Number & " [" & ClipNull(Monsterrec.Name) & "] -- Existing Record, would be updated.")
Else
    nStatus = BTRCALL(BUPDATE, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        ts.WriteLine ("Monster #" & Monsterrec.Number & " [" & ClipNull(Monsterrec.Name) & "] -- Update Error: " & nStatus)
    Else
        If optErrorsOnly.Value = True Or optErrorsNSkips.Value = True Then Exit Sub
        ts.WriteLine ("Monster #" & Monsterrec.Number & " [" & ClipNull(Monsterrec.Name) & "] -- Update Successful.")
    End If
End If
End Sub

Private Sub CreateLogFile(ByVal fil As String)

'Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(fil) = True Then Call fso.DeleteFile(fil, True)

Call fso.CreateTextFile(fil)

End Sub
'Private Function StripSpaces(ByVal sStr As String, ByVal nLen As Integer) As String
'StripSpaces = sStr
'If Not Right(StripSpaces, 1) = Chr(0) Then StripSpaces = RTrim(StripSpaces) & Chr(0)
'If Len(StripSpaces) < nLen Then
'    StripSpaces = StripSpaces & Chr(0)
'ElseIf Len(StripSpaces) > nLen Then
'    StripSpaces = Left(StripSpaces, nLen)
'End If
'End Function
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    If cmdGo.Enabled = False Then
        Cancel = 1
        Exit Sub
    End If
    
    Call WriteINI("Options", "ImportPath", Dir1.Path)
    Call WriteINI("Options", "ImportRoomsAll", chkRoomsAll.Value)
    Call WriteINI("Options", "ImportRoomsFrom", Val(txtRoomsFrom.Text))
    Call WriteINI("Options", "ImportRoomsTo", Val(txtRoomsTo.Text))
    Call WriteINI("Options", "ImportRoomsMap", Val(txtRoomsMap.Text))
    Call WriteINI("Options", "ImportItemsAll", chkItemsAll.Value)
    Call WriteINI("Options", "ImportItemsFrom", Val(txtItemsFrom.Text))
    Call WriteINI("Options", "ImportItemsTo", Val(txtItemsTo.Text))
    Call WriteINI("Options", "ImportSpellsAll", chkSpellsAll.Value)
    Call WriteINI("Options", "ImportSpellsFrom", Val(txtSpellsFrom.Text))
    Call WriteINI("Options", "ImportSpellsTo", Val(txtSpellsTo.Text))
    Call WriteINI("Options", "ImportMonstersAll", chkMonstersAll.Value)
    Call WriteINI("Options", "ImportMonstersFrom", Val(txtMonstersFrom.Text))
    Call WriteINI("Options", "ImportMonstersTo", Val(txtMonstersTo.Text))
    Call WriteINI("Options", "ImportShopsAll", chkShopsAll.Value)
    Call WriteINI("Options", "ImportShopsFrom", Val(txtShopsFrom.Text))
    Call WriteINI("Options", "ImportShopsTo", Val(txtShopsTo.Text))
    Call WriteINI("Options", "ImportTextblocksAll", chkTextblocksAll.Value)
    Call WriteINI("Options", "ImportTextblocksFrom", Val(txtTextblocksFrom.Text))
    Call WriteINI("Options", "ImportTextblocksTo", Val(txtTextblocksTo.Text))
    Call WriteINI("Options", "ImportRacesAll", chkRacesAll.Value)
    Call WriteINI("Options", "ImportRacesFrom", Val(txtRacesFrom.Text))
    Call WriteINI("Options", "ImportRacesTo", Val(txtRacesTo.Text))
    Call WriteINI("Options", "ImportClassesAll", chkClassesAll.Value)
    Call WriteINI("Options", "ImportClassesFrom", Val(txtClassesFrom.Text))
    Call WriteINI("Options", "ImportClassesTo", Val(txtClassesTo.Text))
    Call WriteINI("Options", "ImportMessagesAll", chkMessagesAll.Value)
    Call WriteINI("Options", "ImportMessagesFrom", Val(txtMessagesFrom.Text))
    Call WriteINI("Options", "ImportMessagesTo", Val(txtMessagesTo.Text))
    
    If Not Me.WindowState = vbMinimized Then
        Call WriteINI("Windows", "ImportTop", Me.Top)
        Call WriteINI("Windows", "ImportLeft", Me.Left)
    End If
    If optUpdate = True Then Call WriteINI("Options", "ImportUpdate", 1)
    If optErrorsNSkips.Value = True Then
        Call WriteINI("Options", "ImportErrorsSkips", 1)
    Else
        Call WriteINI("Options", "ImportErrorsSkips", 0)
    End If
    If optErrorsOnly.Value = True Then
        Call WriteINI("Options", "ImportErrorsOnly", 1)
    Else
        Call WriteINI("Options", "ImportErrorsOnly", 0)
    End If
    
    Call CloseAll
    
    Set ts = Nothing
    Set fso = Nothing
    
End Sub
Private Function CalcTotalRecords() As Long
On Error GoTo error:
Dim nStatus As Integer ', tabTemp As Recordset  '', adoConnect As Database

'Set adoConnect = OpenDatabase(sDataSource)

CalcTotalRecords = 0

If chkItems.Value = 1 Then
    If chkItemsAll.Value = 1 Then
        CalcTotalRecords = CalcTotalRecords + tabItems.RecordCount
    Else
        CalcTotalRecords = CalcTotalRecords + Val(txtItemsTo.Text) - Val(txtItemsFrom.Text) + 1
    End If
End If

If chkSpells.Value = 1 Then
    If chkSpellsAll.Value = 1 Then
        CalcTotalRecords = CalcTotalRecords + tabSpells.RecordCount
    Else
        CalcTotalRecords = CalcTotalRecords + Val(txtSpellsTo.Text) - Val(txtSpellsFrom.Text) + 1
    End If
End If

If chkShops.Value = 1 Then
    If chkShopsAll.Value = 1 Then
        CalcTotalRecords = CalcTotalRecords + tabShops.RecordCount
    Else
        CalcTotalRecords = CalcTotalRecords + Val(txtShopsTo.Text) - Val(txtShopsFrom.Text) + 1
    End If
End If

If chkMonsters.Value = 1 Then
    If chkMonstersAll.Value = 1 Then
        CalcTotalRecords = CalcTotalRecords + tabMonsters.RecordCount
    Else
        CalcTotalRecords = CalcTotalRecords + Val(txtMonstersTo.Text) - Val(txtMonstersFrom.Text) + 1
    End If
End If

If chkRooms.Value = 1 Then
    If chkRoomsAll.Value = 1 Then
        CalcTotalRecords = CalcTotalRecords + tabRooms.RecordCount
    Else
        CalcTotalRecords = CalcTotalRecords + Val(txtRoomsTo.Text) - Val(txtRoomsFrom.Text) + 1
    End If
End If

If chkMessages.Value = 1 Then
    If chkMessagesAll.Value = 1 Then
        CalcTotalRecords = CalcTotalRecords + tabMessages.RecordCount
    Else
        CalcTotalRecords = CalcTotalRecords + Val(txtMessagesTo.Text) - Val(txtMessagesFrom.Text) + 1
    End If
End If

If chkRaces.Value = 1 Then
    If chkRacesAll.Value = 1 Then
        CalcTotalRecords = CalcTotalRecords + tabRaces.RecordCount
    Else
        CalcTotalRecords = CalcTotalRecords + Val(txtRacesTo.Text) - Val(txtRacesFrom.Text) + 1
    End If
End If

If chkClasses.Value = 1 Then
    If chkClassesAll.Value = 1 Then
        CalcTotalRecords = CalcTotalRecords + tabClasses.RecordCount
    Else
        CalcTotalRecords = CalcTotalRecords + Val(txtClassesTo.Text) - Val(txtClassesFrom.Text) + 1
    End If
End If

If chkActions.Value = 1 Then
    CalcTotalRecords = CalcTotalRecords + tabActions.RecordCount
End If

If chkTextblocks.Value = 1 Then
    If chkTextblocksAll.Value = 1 Then
        CalcTotalRecords = CalcTotalRecords + tabTextblocks.RecordCount
    Else
        CalcTotalRecords = CalcTotalRecords + Val(txtTextblocksTo.Text) - Val(txtTextblocksFrom.Text) + 1
    End If
End If

If CalcTotalRecords <= 0 Then CalcTotalRecords = 1
'If CalcTotalRecords > 32767 Then CalcTotalRecords = 32767

Exit Function

error:
Call HandleError
'Set tabTemp = Nothing
'Set adoConnect = Nothing
End Function
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

Private Sub Label15_Click()

End Sub

Private Sub lblAll_Click()

If lblAll.Tag = "1" Then
    chkItemsAll.Value = 1
    chkMonstersAll.Value = 1
    chkTextblocksAll.Value = 1
    chkMessagesAll.Value = 1
    chkClassesAll.Value = 1
    chkRacesAll.Value = 1
    chkSpellsAll.Value = 1
    chkShopsAll.Value = 1
    chkRoomsAll.Value = 1
    lblAll.Tag = 0
Else
    chkItemsAll.Value = 0
    chkMonstersAll.Value = 0
    chkTextblocksAll.Value = 0
    chkMessagesAll.Value = 0
    chkClassesAll.Value = 0
    chkRacesAll.Value = 0
    chkSpellsAll.Value = 0
    chkShopsAll.Value = 0
    chkRoomsAll.Value = 0
    lblAll.Tag = 1
End If

End Sub
Private Function CheckError() As Boolean
Dim nYesNo As Integer

CheckError = False
'Debug.Print Err.Description
If Err.Number = 3265 Then
    If bSkipMissing = True Then
        CheckError = True
        Exit Function
    End If
    
    nYesNo = MsgBox("A table in this export file (for " & stsStatusBar.Panels(1).Text & ") is missing some of the required fields that NMR now handles." _
        & vbCrLf & "Do you want to continue importing anyway and import what is there?", vbYesNo + vbQuestion)
    
    If nYesNo = vbYes Then
        bSkipMissing = True
        CheckError = True
    End If
Else
    Call HandleError
End If
End Function

Private Sub txtClassesFrom_GotFocus()
Call SelectAll(txtClassesFrom)

End Sub

Private Sub txtClassesTo_GotFocus()
Call SelectAll(txtClassesTo)

End Sub

Private Sub txtItemsFrom_GotFocus()
Call SelectAll(txtItemsFrom)

End Sub

Private Sub txtItemsTo_GotFocus()
Call SelectAll(txtItemsTo)

End Sub

Private Sub txtMessagesFrom_GotFocus()
Call SelectAll(txtMessagesFrom)

End Sub

Private Sub txtMessagesTo_GotFocus()
Call SelectAll(txtMessagesTo)

End Sub

Private Sub txtMonstersFrom_GotFocus()
Call SelectAll(txtMonstersFrom)

End Sub

Private Sub txtMonstersTo_GotFocus()
Call SelectAll(txtMonstersTo)

End Sub

Private Sub txtRacesFrom_GotFocus()
Call SelectAll(txtRacesFrom)

End Sub

Private Sub txtRacesTo_GotFocus()
Call SelectAll(txtRacesTo)

End Sub

Private Sub txtRoomsFrom_GotFocus()
Call SelectAll(txtRoomsFrom)

End Sub

Private Sub txtRoomsMap_GotFocus()
Call SelectAll(txtRoomsMap)

End Sub

Private Sub txtRoomsTo_GotFocus()
Call SelectAll(txtRoomsTo)

End Sub

Private Sub txtShopsFrom_GotFocus()
Call SelectAll(txtShopsFrom)

End Sub

Private Sub txtShopsTo_GotFocus()
Call SelectAll(txtShopsTo)

End Sub

Private Sub txtSpellsFrom_GotFocus()
Call SelectAll(txtSpellsFrom)

End Sub

Private Sub txtSpellsTo_GotFocus()
Call SelectAll(txtSpellsTo)

End Sub

Private Sub txtTextblocksFrom_GotFocus()
Call SelectAll(txtTextblocksFrom)

End Sub

Private Sub txtTextblocksTo_GotFocus()
Call SelectAll(txtTextblocksTo)

End Sub
