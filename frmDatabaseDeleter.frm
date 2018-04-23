VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDatabaseDeleter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Record Deleter"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   Icon            =   "frmDatabaseDeleter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6015
   ScaleWidth      =   4230
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel / Close"
      Height          =   375
      Left            =   2520
      TabIndex        =   52
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Delete Selected"
      Height          =   375
      Left            =   120
      TabIndex        =   51
      Top             =   5280
      Width           =   1515
   End
   Begin VB.Frame fra1 
      Height          =   4755
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   3975
      Begin VB.TextBox txtRoomMap 
         Height          =   285
         Left            =   2820
         TabIndex        =   39
         Text            =   "0"
         Top             =   3720
         Width           =   435
      End
      Begin VB.TextBox txtRoomsFrom 
         Height          =   285
         Left            =   2220
         TabIndex        =   40
         Text            =   "0"
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox txtRoomsTo 
         Height          =   285
         Left            =   3060
         TabIndex        =   41
         Text            =   "9999"
         Top             =   3360
         Width           =   735
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1740
         TabIndex        =   47
         Top             =   4365
         Value           =   2  'Grayed
         Width           =   195
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1740
         TabIndex        =   45
         Top             =   4050
         Value           =   2  'Grayed
         Width           =   195
      End
      Begin VB.CheckBox chkTextblocksAll 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1740
         TabIndex        =   34
         Top             =   3030
         Width           =   195
      End
      Begin VB.CheckBox chkMessagesAll 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1740
         TabIndex        =   30
         Top             =   2670
         Width           =   195
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1740
         TabIndex        =   43
         Top             =   3720
         Value           =   2  'Grayed
         Width           =   195
      End
      Begin VB.CheckBox chkRoomsAll 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1740
         TabIndex        =   38
         Top             =   3390
         Width           =   195
      End
      Begin VB.CheckBox chkClassesAll 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1740
         TabIndex        =   26
         Top             =   2310
         Width           =   195
      End
      Begin VB.CheckBox chkRacesAll 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1740
         TabIndex        =   22
         Top             =   1965
         Width           =   195
      End
      Begin VB.CheckBox chkShopsAll 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1740
         TabIndex        =   18
         Top             =   1620
         Width           =   195
      End
      Begin VB.CheckBox chkSpellsAll 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1740
         TabIndex        =   14
         Top             =   1260
         Width           =   195
      End
      Begin VB.CheckBox chkMonstersAll 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1740
         TabIndex        =   10
         Top             =   900
         Width           =   195
      End
      Begin VB.TextBox txtTextblocksTo 
         Height          =   285
         Left            =   3060
         TabIndex        =   36
         Text            =   "9999"
         Top             =   2985
         Width           =   735
      End
      Begin VB.TextBox txtTextblocksFrom 
         Height          =   285
         Left            =   2220
         TabIndex        =   35
         Text            =   "0"
         Top             =   2985
         Width           =   735
      End
      Begin VB.TextBox txtMessagesTo 
         Height          =   285
         Left            =   3060
         TabIndex        =   32
         Text            =   "9999"
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtMessagesFrom 
         Height          =   285
         Left            =   2220
         TabIndex        =   31
         Text            =   "0"
         Top             =   2625
         Width           =   735
      End
      Begin VB.TextBox txtClassesTo 
         Height          =   285
         Left            =   3060
         TabIndex        =   28
         Text            =   "9999"
         Top             =   2265
         Width           =   735
      End
      Begin VB.TextBox txtClassesFrom 
         Height          =   285
         Left            =   2220
         TabIndex        =   27
         Text            =   "0"
         Top             =   2265
         Width           =   735
      End
      Begin VB.TextBox txtRacesTo 
         Height          =   285
         Left            =   3060
         TabIndex        =   24
         Text            =   "9999"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtRacesFrom 
         Height          =   285
         Left            =   2220
         TabIndex        =   23
         Text            =   "0"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtShopsTo 
         Height          =   285
         Left            =   3060
         TabIndex        =   20
         Text            =   "9999"
         Top             =   1575
         Width           =   735
      End
      Begin VB.TextBox txtShopsFrom 
         Height          =   285
         Left            =   2220
         TabIndex        =   19
         Text            =   "0"
         Top             =   1575
         Width           =   735
      End
      Begin VB.TextBox txtSpellsTo 
         Height          =   285
         Left            =   3060
         TabIndex        =   16
         Text            =   "9999"
         Top             =   1215
         Width           =   735
      End
      Begin VB.TextBox txtSpellsFrom 
         Height          =   285
         Left            =   2220
         TabIndex        =   15
         Text            =   "0"
         Top             =   1215
         Width           =   735
      End
      Begin VB.TextBox txtMonstersTo 
         Height          =   285
         Left            =   3060
         TabIndex        =   12
         Text            =   "9999"
         Top             =   855
         Width           =   735
      End
      Begin VB.TextBox txtMonstersFrom 
         Height          =   285
         Left            =   2220
         TabIndex        =   11
         Text            =   "0"
         Top             =   855
         Width           =   735
      End
      Begin VB.CheckBox chkItemsAll 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1740
         TabIndex        =   6
         Top             =   540
         Width           =   195
      End
      Begin VB.TextBox txtItemsTo 
         Height          =   285
         Left            =   3060
         TabIndex        =   8
         Text            =   "9999"
         Top             =   495
         Width           =   735
      End
      Begin VB.TextBox txtItemsFrom 
         Height          =   285
         Left            =   2220
         TabIndex        =   7
         Text            =   "0"
         Top             =   495
         Width           =   735
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
         TabIndex        =   5
         Top             =   510
         Width           =   1515
      End
      Begin VB.CheckBox chkBankbooks 
         Caption         =   "Bankbooks"
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
         TabIndex        =   46
         Top             =   4335
         Width           =   1515
      End
      Begin VB.CheckBox chkUsers 
         Caption         =   "Users"
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
         TabIndex        =   44
         Top             =   4020
         Width           =   1515
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
         TabIndex        =   33
         Top             =   3000
         Width           =   1515
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
         TabIndex        =   29
         Top             =   2640
         Width           =   1515
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
         Height          =   195
         Left            =   180
         TabIndex        =   42
         Top             =   3720
         Width           =   1515
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
         TabIndex        =   37
         Top             =   3360
         Width           =   1515
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
         TabIndex        =   25
         Top             =   2280
         Width           =   1515
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
         TabIndex        =   21
         Top             =   1935
         Width           =   1515
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
         TabIndex        =   17
         Top             =   1590
         Width           =   1515
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
         TabIndex        =   13
         Top             =   1230
         Width           =   1515
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
         TabIndex        =   9
         Top             =   870
         Width           =   1515
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
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
         Left            =   2820
         TabIndex        =   50
         Top             =   3960
         Width           =   435
      End
      Begin VB.Label Label16 
         Caption         =   "Delete:"
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
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "All"
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
         Left            =   1620
         TabIndex        =   2
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "To"
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
         Left            =   3060
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "From"
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
         Left            =   2220
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
   End
   Begin MSComctlLib.StatusBar stsStatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   49
      Top             =   5760
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4842
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   120
      TabIndex        =   48
      Top             =   4920
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmDatabaseDeleter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit
Dim nScale As Integer
Dim nScaleCount As Long
Dim bStopDelete As Boolean

Private Sub cmdClose_Click()

End Sub

Private Sub chkRoomsAll_Click()
If chkRoomsAll.Value = 1 Then
    txtRoomsFrom.Enabled = False
    txtRoomsTo.Enabled = False
    txtRoomMap.Enabled = False
Else
    txtRoomsFrom.Enabled = True
    txtRoomsTo.Enabled = True
    txtRoomMap.Enabled = True
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
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
Private Sub cmdGo_Click()
On Error GoTo error:
Dim nYesNo As Integer, nMaxValue As Long
Dim CheckboxArray(1 To 12) As Object, nFilesToDelete As Long, x As Integer

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

nYesNo = MsgBox("Are you sure you want to delete the selected records?", vbYesNo, "Confirm Deletion")
If nYesNo <> 6 Then Exit Sub

Call UnloadForms(Me.Name)
Call LockMenus
cmdGo.Enabled = False
fra1.Enabled = False
'cmdCancel.Enabled = False
'frmMain.Enabled = False
nFilesToDelete = 0

Set CheckboxArray(1) = chkMessages
Set CheckboxArray(2) = chkItems
Set CheckboxArray(3) = chkSpells
Set CheckboxArray(4) = chkClasses
Set CheckboxArray(5) = chkRaces
Set CheckboxArray(6) = chkShops
Set CheckboxArray(7) = chkRooms
Set CheckboxArray(8) = chkActions
Set CheckboxArray(9) = chkMonsters
Set CheckboxArray(10) = chkUsers
Set CheckboxArray(11) = chkBankbooks
Set CheckboxArray(12) = chkTextblocks

nMaxValue = CalcTotalRecords
Call SetRange(nMaxValue)
ProgressBar.Visible = True

DoEvents
bStopDelete = False
For x = 1 To UBound(CheckboxArray())
    If bStopDelete Then Exit For
    If CheckboxArray(x).Value = 1 Then
        If x = 1 Then Call DeleteMessages
        If x = 2 Then Call DeleteItems
        If x = 3 Then Call DeleteSpells
        If x = 4 Then Call DeleteClasses
        If x = 5 Then Call DeleteRaces
        If x = 6 Then Call DeleteShops
        If x = 7 Then Call DeleteRooms
        If x = 8 Then Call DeleteActions
        If x = 9 Then Call DeleteMonsters
        If x = 10 Then Call DeleteUsers
        If x = 11 Then Call DeleteBankbooks
        If x = 12 Then Call DeleteTextblocks
        DoEvents
    End If
Next

If bStopDelete Then GoTo ReEnable:

ProgressBar.Value = ProgressBar.Max
MsgBox "Deletion Complete.", vbInformation

ReEnable:
On Error Resume Next
ProgressBar.Visible = False
stsStatusBar.Panels(1).Text = ""
stsStatusBar.Panels(2).Text = ""
frmMain.Enabled = True
cmdGo.Enabled = True
fra1.Enabled = True
cmdCancel.Enabled = True
Call UnLockMenus

Exit Sub
error:
Call HandleError
Resume ReEnable:
End Sub

Private Sub cmdCancel_Click()
Dim nYesNo As Integer
If cmdGo.Enabled = False Then
    nYesNo = MsgBox("Are you sure you want to cancel?", vbYesNo + vbQuestion + vbDefaultButton2)
    If Not nYesNo = vbYes Then Exit Sub
    cmdCancel.Enabled = False
    bStopDelete = True
    DoEvents
Else
    Unload Me
End If
End Sub

Private Sub DeleteBankbooks()
Dim nStatus As Integer, recnum As Long

recnum = 1
stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_BANKS
stsStatusBar.Panels(2).Text = recnum

nStatus = BTRCALL(BGETFIRST, BankPosBlock, BankDatabuf, Len(BankDatabuf), ByVal BankKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Bank, BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
End If
    
Do While nStatus = 0
        If bStopDelete Then Exit Sub
        nStatus = BTRCALL(BDELETE, BankPosBlock, BankDatabuf, Len(BankDatabuf), ByVal BankKeyBuffer, KEY_BUF_LEN, 0)
        nStatus = BTRCALL(BGETNEXT, BankPosBlock, BankDatabuf, Len(BankDatabuf), ByVal BankKeyBuffer, KEY_BUF_LEN, 0)
        
        IncreaseProgressBar
        recnum = recnum + 1
        stsStatusBar.Panels(2).Text = recnum
        If Not bUseCPU Then DoEvents

Loop

End Sub

Private Sub DeleteTextblocks()
Dim nStatus As Integer, TextblocksTo As Long, x As Long
Dim recnum As Long, TextblocksFrom As Long, y As Integer


recnum = 1
stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_TEXT
stsStatusBar.Panels(2).Text = recnum

    nStatus = BTRCALL(BGETFIRST, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "GETFIRST Textblock, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
    
    If chkTextblocksAll.Value <> 1 Then GoTo range:
    Do While nStatus = 0
        If bStopDelete Then Exit Sub
        nStatus = BTRCALL(BDELETE, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
        nStatus = BTRCALL(BGETNEXT, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
        
        IncreaseProgressBar
        recnum = recnum + 1
        stsStatusBar.Panels(2).Text = recnum
        If Not bUseCPU Then DoEvents
    Loop

Exit Sub

range:
TextblocksFrom = Val(txtTextblocksFrom.Text)
TextblocksTo = Val(txtTextblocksTo.Text)
TextblockKey.PartNum = 0
TextblockKey.Number = TextblocksFrom
nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, Len(TextblockDataBuf), TextblockKeyStructToRow(), KEY_BUF_LEN, 0)

'Do While nStatus = 0
'    If bStopDelete Then Exit Do
'    nStatus = BTRCALL(BDELETE, TextblockPosBlock, TextblockDataBuf, Len(TextblockDataBuf), ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
'    nStatus = BTRCALL(BGETNEXT, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
'
'    If nStatus = 0 Then
'        TextblockRowToStruct TextblockDataBuf.buf
'        If TextblockRec.Number > TextblocksTo Then Exit Do
'
'        Call IncreaseProgressBar
'        stsStatusBar.Panels(2).Text = TextblockRec.Number
'    End If
'
'    If Not bUseCPU Then DoEvents
'Loop

For x = TextblocksFrom To TextblocksTo
    TextblockKey.Number = x
    For y = 0 To 20
        If bStopDelete Then Exit Sub
        TextblockKey.PartNum = y
        nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, Len(TextblockDataBuf), TextblockKeyStructToRow(), KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            nStatus = BTRCALL(BDELETE, TextblockPosBlock, TextblockDataBuf, Len(TextblockDataBuf), ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
        Else
            'Exit For
        End If
    Next y
    
    Call IncreaseProgressBar
    stsStatusBar.Panels(2).Text = x
    If Not bUseCPU Then DoEvents
Next x

End Sub

Private Sub DeleteMessages()
Dim nStatus As Integer, recnum As Long, MessagesTo As Long, x As Long, MessagesFrom As Long

recnum = 1
stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_MSG
stsStatusBar.Panels(2).Text = recnum

nStatus = BTRCALL(BGETFIRST, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "GETFIRST Messages, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
    
    If chkMessagesAll.Value <> 1 Then GoTo range:
    Do While nStatus = 0
        If bStopDelete Then Exit Sub
        nStatus = BTRCALL(BDELETE, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
        nStatus = BTRCALL(BGETNEXT, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
        
        IncreaseProgressBar
        recnum = recnum + 1
        stsStatusBar.Panels(2).Text = recnum
        If Not bUseCPU Then DoEvents
    Loop

Exit Sub
range:
MessagesFrom = Val(txtMessagesFrom.Text)
MessagesTo = Val(txtMessagesTo.Text)

For x = MessagesFrom To MessagesTo
        If bStopDelete Then Exit Sub
        nStatus = BTRCALL(BGETEQUAL, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), x, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            nStatus = BTRCALL(BDELETE, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
        End If
        IncreaseProgressBar
        stsStatusBar.Panels(2).Text = x
        If Not bUseCPU Then DoEvents
Next x

End Sub

Private Sub DeleteItems()
Dim nStatus As Integer, recnum As Long, ItemsFrom As Long, ItemsTo As Long, x As Long

recnum = 1
stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_ITEMS
stsStatusBar.Panels(2).Text = recnum

nStatus = BTRCALL(BGETFIRST, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "GETFIRST Items, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
    
    If chkItemsAll.Value <> 1 Then GoTo range:
    Do While nStatus = 0
        If bStopDelete Then Exit Sub
        nStatus = BTRCALL(BDELETE, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
        nStatus = BTRCALL(BGETNEXT, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
        
        IncreaseProgressBar
        recnum = recnum + 1
        stsStatusBar.Panels(2).Text = recnum
        If Not bUseCPU Then DoEvents
    Loop
Exit Sub
range:
ItemsFrom = Val(txtItemsFrom.Text)
ItemsTo = Val(txtItemsTo.Text)

For x = ItemsFrom To ItemsTo
        If bStopDelete Then Exit Sub
        nStatus = BTRCALL(BGETEQUAL, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), x, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            nStatus = BTRCALL(BDELETE, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
        End If
        IncreaseProgressBar
        stsStatusBar.Panels(2).Text = x
        If Not bUseCPU Then DoEvents
Next x

End Sub
Private Sub DeleteRooms()
Dim nStatus As Integer, recnum As Long, x As Long

recnum = 1
stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_MP
stsStatusBar.Panels(2).Text = recnum

If chkRoomsAll.Value = 1 Then
    nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "GETFIRST Rooms, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
Else
    If Val(txtRoomsFrom.Text) > Val(txtRoomsTo.Text) Then
        MsgBox "Illegal Room Range Entered!", vbExclamation
        Exit Sub
    End If
    
    RoomKeyStruct.MapNum = Val(txtRoomMap.Text)
    RoomKeyStruct.RoomNum = Val(txtRoomsFrom.Text)
    
    nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), RoomKeyStruct, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Couldn't get first room, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
End If

If chkRoomsAll.Value = 1 Then
    Do While nStatus = 0
        If bStopDelete Then Exit Sub
        nStatus = BTRCALL(BDELETE, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
        nStatus = BTRCALL(BGETNEXT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
        
        Call IncreaseProgressBar
        recnum = recnum + 1
        stsStatusBar.Panels(2).Text = recnum
        If Not bUseCPU Then DoEvents
    Loop
Else
    For x = Val(txtRoomsFrom.Text) To Val(txtRoomsTo.Text)
        If bStopDelete Then Exit Sub
        RoomKeyStruct.RoomNum = x
    
        nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), RoomKeyStruct, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            nStatus = BTRCALL(BDELETE, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
        End If
        
        Call IncreaseProgressBar
        recnum = recnum + 1
        stsStatusBar.Panels(2).Text = recnum
        If Not bUseCPU = True Then DoEvents
    Next
End If

End Sub
Private Sub DeleteSpells()
Dim nStatus As Integer, recnum As Long, SpellsTo As Integer, x As Long, SpellsFrom As Integer

recnum = 1
stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_SPELS
stsStatusBar.Panels(2).Text = recnum

nStatus = BTRCALL(BGETFIRST, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "GETFIRST Spells, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
    
    If chkSpellsAll.Value <> 1 Then GoTo range:
    Do While nStatus = 0
        If bStopDelete Then Exit Sub
        nStatus = BTRCALL(BDELETE, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
        nStatus = BTRCALL(BGETNEXT, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
        
        IncreaseProgressBar
        recnum = recnum + 1
        stsStatusBar.Panels(2).Text = recnum
        If Not bUseCPU Then DoEvents
    Loop

Exit Sub
range:
SpellsFrom = Val(txtSpellsFrom.Text)
SpellsTo = Val(txtSpellsTo.Text)

For x = SpellsFrom To SpellsTo
        If bStopDelete Then Exit Sub
        nStatus = BTRCALL(BGETEQUAL, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), x, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            nStatus = BTRCALL(BDELETE, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
        End If
        IncreaseProgressBar
        stsStatusBar.Panels(2).Text = x
        If Not bUseCPU Then DoEvents
Next x

End Sub

Private Sub DeleteActions()
Dim nStatus As Integer, recnum As Long

recnum = 1
stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_ACTS
stsStatusBar.Panels(2).Text = recnum

nStatus = BTRCALL(BGETFIRST, ActionPosBlock, ActionDatabuf, Len(ActionDatabuf), ByVal ActionKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "GETFIRST actions, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    Else
    End If


Do While nStatus = 0
        If bStopDelete Then Exit Sub
        nStatus = BTRCALL(BDELETE, ActionPosBlock, ActionDatabuf, Len(ActionDatabuf), ByVal ActionKeyBuffer, KEY_BUF_LEN, 0)
        nStatus = BTRCALL(BGETNEXT, ActionPosBlock, ActionDatabuf, Len(ActionDatabuf), ByVal ActionKeyBuffer, KEY_BUF_LEN, 0)
        
        IncreaseProgressBar
        recnum = recnum + 1
        stsStatusBar.Panels(2).Text = recnum
        If Not bUseCPU Then DoEvents
    Loop

End Sub
Private Sub DeleteClasses()
Dim nStatus As Integer, recnum As Long, ClassesTo As Integer, x As Long, ClassesFrom As Integer

recnum = 1
stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_CLASS
stsStatusBar.Panels(2).Text = recnum

nStatus = BTRCALL(BGETFIRST, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "GETFIRST classes, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
    
    If chkClassesAll.Value <> 1 Then GoTo range:
    Do While nStatus = 0
        If bStopDelete Then Exit Sub
        nStatus = BTRCALL(BDELETE, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
        nStatus = BTRCALL(BGETNEXT, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
        
        IncreaseProgressBar
        recnum = recnum + 1
        stsStatusBar.Panels(2).Text = recnum
        If Not bUseCPU Then DoEvents
Loop

Exit Sub
range:
ClassesFrom = Val(txtClassesFrom.Text)
ClassesTo = Val(txtClassesTo.Text)

For x = ClassesFrom To ClassesTo
    If bStopDelete Then Exit Sub
    nStatus = BTRCALL(BGETEQUAL, ClassPosBlock, Classdatabuf, Len(Classdatabuf), x, KEY_BUF_LEN, 0)
    If nStatus = 0 Then
        nStatus = BTRCALL(BDELETE, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
    End If
    IncreaseProgressBar
    stsStatusBar.Panels(2).Text = x
    If Not bUseCPU Then DoEvents
Next x

End Sub
Private Sub DeleteRaces()
Dim nStatus As Integer, recnum As Long, RacesTo As Integer, x As Integer, RacesFrom As Integer

recnum = 1
stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_RACE
stsStatusBar.Panels(2).Text = recnum

nStatus = BTRCALL(BGETFIRST, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "GETFIRST races, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
    
    If chkRacesAll.Value <> 1 Then GoTo range:
    Do While nStatus = 0
        If bStopDelete Then Exit Sub
        nStatus = BTRCALL(BDELETE, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
        nStatus = BTRCALL(BGETNEXT, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
        
        IncreaseProgressBar
        recnum = recnum + 1
        stsStatusBar.Panels(2).Text = recnum
        If Not bUseCPU Then DoEvents
    Loop

Exit Sub
range:
RacesFrom = Val(txtRacesFrom.Text)
RacesTo = Val(txtRacesTo.Text)

For x = RacesFrom To RacesTo
        If bStopDelete Then Exit Sub
        nStatus = BTRCALL(BGETEQUAL, RacePosBlock, Racedatabuf, Len(Racedatabuf), x, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            nStatus = BTRCALL(BDELETE, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
        End If
        IncreaseProgressBar
        stsStatusBar.Panels(2).Text = x
        If Not bUseCPU Then DoEvents
Next x

End Sub
Private Sub DeleteShops()
Dim nStatus As Integer, recnum As Long, ShopsTo As Long, x As Long, ShopsFrom As Long

recnum = 1
stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_SHOPS
stsStatusBar.Panels(2).Text = recnum

nStatus = BTRCALL(BGETFIRST, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "GETFIRST shops, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If

    If chkShopsAll.Value <> 1 Then GoTo range:
    Do While nStatus = 0
        If bStopDelete Then Exit Sub
        nStatus = BTRCALL(BDELETE, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
        nStatus = BTRCALL(BGETNEXT, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
        
        IncreaseProgressBar
        recnum = recnum + 1
        stsStatusBar.Panels(2).Text = recnum
        If Not bUseCPU Then DoEvents
    Loop


Exit Sub
range:
ShopsFrom = Val(txtShopsFrom.Text)
ShopsTo = Val(txtShopsTo.Text)

For x = ShopsFrom To ShopsTo
        If bStopDelete Then Exit Sub
        nStatus = BTRCALL(BGETEQUAL, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), x, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            nStatus = BTRCALL(BDELETE, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
        End If
        IncreaseProgressBar
        stsStatusBar.Panels(2).Text = x
        If Not bUseCPU Then DoEvents
Next x

End Sub
Private Sub DeleteMonsters()
Dim nStatus As Integer, recnum As Long, MonstersTo As Long, x As Long, MonstersFrom As Long

recnum = 1
stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_KNMSR
stsStatusBar.Panels(2).Text = recnum

nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "GETFIRST shops, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
    
    If chkMonstersAll.Value <> 1 Then GoTo range:
    Do While nStatus = 0
        If bStopDelete Then Exit Sub
        nStatus = BTRCALL(BDELETE, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
        nStatus = BTRCALL(BGETNEXT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
        
        IncreaseProgressBar
        recnum = recnum + 1
        stsStatusBar.Panels(2).Text = recnum
        If Not bUseCPU Then DoEvents
    Loop

Exit Sub
range:
MonstersFrom = Val(txtMonstersFrom.Text)
MonstersTo = Val(txtMonstersTo.Text)

For x = MonstersFrom To MonstersTo
        If bStopDelete Then Exit Sub
        nStatus = BTRCALL(BGETEQUAL, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), x, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            nStatus = BTRCALL(BDELETE, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
        End If
        IncreaseProgressBar
        stsStatusBar.Panels(2).Text = x
        If Not bUseCPU Then DoEvents
Next x

End Sub
Private Sub DeleteUsers()
Dim nStatus As Integer, recnum As Long

recnum = 1
stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_USERS
stsStatusBar.Panels(2).Text = recnum

nStatus = BTRCALL(BGETFIRST, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "User, BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
End If
    
Do While nStatus = 0
        If bStopDelete Then Exit Sub
        nStatus = BTRCALL(BDELETE, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0)
        nStatus = BTRCALL(BGETNEXT, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0)
        
        IncreaseProgressBar
        recnum = recnum + 1
        stsStatusBar.Panels(2).Text = recnum
        If Not bUseCPU Then DoEvents

Loop

End Sub

Private Function CalcTotalRecords() As Long
On Error GoTo error:
Dim nStatus As Integer

CalcTotalRecords = 0

If chkItems.Value = 1 Then
    If chkItemsAll.Value = 1 Then
        nStatus = BTRCALL(BSTAT, ItemPosBlock, DBStatDatabuf, Len(Itemdatabuf), 0, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            CalcTotalRecords = CalcTotalRecords + 1800
        Else
            DBStatRowToStruct DBStatDatabuf.buf
            CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
        End If
    Else
        CalcTotalRecords = CalcTotalRecords + Val(txtItemsTo.Text) - Val(txtItemsFrom.Text) + 1
    End If
End If

If chkSpells.Value = 1 Then
    If chkSpellsAll.Value = 1 Then
        nStatus = BTRCALL(BSTAT, SpellPosBlock, DBStatDatabuf, Len(Spelldatabuf), 0, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            CalcTotalRecords = CalcTotalRecords + 1300
        Else
            DBStatRowToStruct DBStatDatabuf.buf
            CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
        End If
    Else
        CalcTotalRecords = CalcTotalRecords + Val(txtSpellsTo.Text) - Val(txtSpellsFrom.Text) + 1
    End If
End If

If chkShops.Value = 1 Then
    If chkShopsAll.Value = 1 Then
        nStatus = BTRCALL(BSTAT, ShopPosBlock, DBStatDatabuf, Len(Shopdatabuf), 0, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            CalcTotalRecords = CalcTotalRecords + 200
        Else
            DBStatRowToStruct DBStatDatabuf.buf
            CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
        End If
    Else
        CalcTotalRecords = CalcTotalRecords + Val(txtShopsTo.Text) - Val(txtShopsFrom.Text) + 1
    End If
End If

If chkActions.Value = 1 Then
    nStatus = BTRCALL(BSTAT, ActionPosBlock, DBStatDatabuf, Len(ActionDatabuf), 0, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        CalcTotalRecords = CalcTotalRecords + 100
    Else
        DBStatRowToStruct DBStatDatabuf.buf
        CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
    End If
End If

If chkMonsters.Value = 1 Then
    If chkMonstersAll.Value = 1 Then
        nStatus = BTRCALL(BSTAT, MonsterPosBlock, DBStatDatabuf, Len(Monsterdatabuf), 0, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            CalcTotalRecords = CalcTotalRecords + 1100
        Else
            DBStatRowToStruct DBStatDatabuf.buf
            CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
        End If
    Else
        CalcTotalRecords = CalcTotalRecords + Val(txtMonstersTo.Text) - Val(txtMonstersFrom.Text) + 1
    End If
End If

If chkTextblocks.Value = 1 Then
    If chkTextblocksAll.Value = 1 Then
        nStatus = BTRCALL(BSTAT, TextblockPosBlock, DBStatDatabuf, Len(TextblockDataBuf), 0, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            CalcTotalRecords = CalcTotalRecords + 2600
        Else
            DBStatRowToStruct DBStatDatabuf.buf
            CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
        End If
    Else
        CalcTotalRecords = CalcTotalRecords + Val(txtTextblocksTo.Text) - Val(txtTextblocksFrom.Text) + 1
    End If
End If

If chkMessages.Value = 1 Then
    If chkMessagesAll.Value = 1 Then
        nStatus = BTRCALL(BSTAT, MessagePosBlock, DBStatDatabuf, Len(Messagedatabuf), 0, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            CalcTotalRecords = CalcTotalRecords + 3700
        Else
            DBStatRowToStruct DBStatDatabuf.buf
            CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
        End If
    Else
        CalcTotalRecords = CalcTotalRecords + Val(txtMessagesTo.Text) - Val(txtMessagesFrom.Text) + 1
    End If
End If

If chkRaces.Value = 1 Then
    If chkRacesAll.Value = 1 Then
        nStatus = BTRCALL(BSTAT, RacePosBlock, DBStatDatabuf, Len(Racedatabuf), 0, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            CalcTotalRecords = CalcTotalRecords + 30
        Else
            DBStatRowToStruct DBStatDatabuf.buf
            CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
        End If
    Else
        CalcTotalRecords = CalcTotalRecords + Val(txtRacesTo.Text) - Val(txtRacesFrom.Text) + 1
    End If
End If

If chkClasses.Value = 1 Then
    If chkClassesAll.Value = 1 Then
        nStatus = BTRCALL(BSTAT, ClassPosBlock, DBStatDatabuf, Len(Classdatabuf), 0, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            CalcTotalRecords = CalcTotalRecords + 30
        Else
            DBStatRowToStruct DBStatDatabuf.buf
            CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
        End If
    Else
        CalcTotalRecords = CalcTotalRecords + Val(txtClassesTo.Text) - Val(txtClassesFrom.Text) + 1
    End If
End If

If chkRooms.Value = 1 Then
    If chkRoomsAll.Value = 1 Then
        nStatus = BTRCALL(BSTAT, RoomPosBlock, DBStatDatabuf, Len(Roomdatabuf), 0, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            CalcTotalRecords = CalcTotalRecords + 30000
        Else
            DBStatRowToStruct DBStatDatabuf.buf
            CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
        End If
    Else
        CalcTotalRecords = CalcTotalRecords + Val(txtRoomsTo.Text) - Val(txtRoomsFrom.Text) + 1
    End If
End If

If chkUsers.Value = 1 Then
    nStatus = BTRCALL(BSTAT, UserPosBlock, DBStatDatabuf, Len(Userdatabuf), 0, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        CalcTotalRecords = CalcTotalRecords + 100
    Else
        DBStatRowToStruct DBStatDatabuf.buf
        CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
    End If
End If

If chkBankbooks.Value = 1 Then
    nStatus = BTRCALL(BSTAT, BankPosBlock, DBStatDatabuf, Len(BankDatabuf), 0, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        CalcTotalRecords = CalcTotalRecords + 100
    Else
        DBStatRowToStruct DBStatDatabuf.buf
        CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
    End If
End If

If CalcTotalRecords <= 0 Then CalcTotalRecords = 1
If CalcTotalRecords > 32767 Then CalcTotalRecords = 32767

Exit Function

error:
Call HandleError
End Function
Private Sub IncreaseProgressBar()
On Error Resume Next
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

Private Sub Form_Unload(Cancel As Integer)
If cmdGo.Enabled = False Then
    Cancel = 1
    Exit Sub
End If
End Sub

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

Private Sub txtRoomMap_GotFocus()
Call SelectAll(txtRoomMap)

End Sub

Private Sub txtRoomsFrom_GotFocus()
Call SelectAll(txtRoomsFrom)

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
