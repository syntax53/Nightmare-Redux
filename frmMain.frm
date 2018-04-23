VERSION 5.00
Object = "{20D5284F-7B23-4F0A-B8B1-6C9D18B64F1C}#1.0#0"; "exlimiter.ocx"
Object = "{AA61DC5D-A4D1-4F73-AF2B-208862262908}#3.0#0"; "NMRTaskBar.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Nightmare Redux"
   ClientHeight    =   8250
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10635
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":08CA
   StartUpPosition =   2  'CenterScreen
   Begin NMRTaskBar.ctlTaskBar tbTaskBar 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   661
      ButtonHeight    =   22
   End
   Begin MSComctlLib.StatusBar stsStatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7995
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2302
            MinWidth        =   2293
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2461
            MinWidth        =   2470
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin exlimiter.EL EL1 
      Left            =   9780
      Top             =   7140
      _ExtentX        =   1270
      _ExtentY        =   1270
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuDisableWrite 
         Caption         =   "&Disable DB Writing"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Editors"
      Begin VB.Menu mnuEditAction 
         Caption         =   "Actio&ns"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuEditBankbooks 
         Caption         =   "&Bankbooks"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuEditClass 
         Caption         =   "C&lasses"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuEditGangs 
         Caption         =   "&Gangs"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Items"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuEditMessage 
         Caption         =   "M&essages"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuEditMonster 
         Caption         =   "&Monsters"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditRace 
         Caption         =   "&Races"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEditRoom 
         Caption         =   "R&ooms"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuEditShop 
         Caption         =   "S&hops"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuEditSpell 
         Caption         =   "&Spells"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuEditTextblock 
         Caption         =   "&Textblocks"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuEditUser 
         Caption         =   "&Users"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuDatabase 
         Caption         =   "&Databases"
         Begin VB.Menu mnuakjhfw 
            Caption         =   "External Database Actions--"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuDatabaseExporter 
            Caption         =   "&Export Records"
            Shortcut        =   {F12}
         End
         Begin VB.Menu mnuDatabaseIndexChange 
            Caption         =   "Monster Group/Index Changer"
         End
         Begin VB.Menu mnuRecordChanger 
            Caption         =   "&Record Number Changer"
            Shortcut        =   +{F12}
         End
         Begin VB.Menu mnuMMUDExplorer 
            Caption         =   "Export to &MMUD Explorer"
            Shortcut        =   ^{F12}
         End
         Begin VB.Menu mnuaslqkjrg 
            Caption         =   "Internal Database Actions--"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuDatabaseImporter 
            Caption         =   "&Import Records"
            Shortcut        =   {F11}
         End
         Begin VB.Menu mnuDatabaseDeleter 
            Caption         =   "&Delete Records"
         End
      End
      Begin VB.Menu mnuItem 
         Caption         =   "&Items"
         Begin VB.Menu mnuLimitedItemList 
            Caption         =   "&Build Limited Item List"
            Shortcut        =   ^{F8}
         End
         Begin VB.Menu mnuFindItem 
            Caption         =   "&Find an Item"
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu mnuItemsFixUses 
            Caption         =   "Fix Number of Uses"
         End
         Begin VB.Menu mnuNoLimited 
            Caption         =   "No Limited &Items"
         End
         Begin VB.Menu mnuNoLevelRestrictions 
            Caption         =   "No Level &Restrictions"
         End
      End
      Begin VB.Menu mnuMonster 
         Caption         =   "&Monsters"
         Begin VB.Menu mnuMonsterAttackSim 
            Caption         =   "&Attack Simulator"
         End
         Begin VB.Menu mnuDivideExp 
            Caption         =   "&Divide All Exp"
         End
         Begin VB.Menu mnuExp 
            Caption         =   "&Multiply All Exp"
         End
         Begin VB.Menu mnuMultiplyBossExp 
            Caption         =   "Multiply &Boss Exp"
         End
         Begin VB.Menu mnuMonstersCombineExp 
            Caption         =   "Combine Exp and Exp Multi Fields"
         End
         Begin VB.Menu mnudash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuResetMonsterKillsToTime 
            Caption         =   "&Set Last Killed Times to Date"
         End
         Begin VB.Menu mnuResetMonsterKills 
            Caption         =   "&Reset Last Killed Times to Zero"
         End
         Begin VB.Menu mnudash2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFixMonsterUses 
            Caption         =   "Fix Number of &Uses on Item Drops"
         End
      End
      Begin VB.Menu mnuRooms 
         Caption         =   "R&ooms"
         Begin VB.Menu mnuChangeRoomCallLetters 
            Caption         =   "&Change dat call letters"
         End
         Begin VB.Menu mnuMassRoomEditor 
            Caption         =   "&Mass Room Editor"
            Shortcut        =   ^{F5}
         End
         Begin VB.Menu mnufds 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRoomPad 
            Caption         =   "&Insert Buffer Rooms"
         End
         Begin VB.Menu mnuDeleteBufferRooms 
            Caption         =   "&Delete Buffer Rooms"
         End
         Begin VB.Menu mnudash3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRoomsCombineItems 
            Caption         =   "Combine like items on floor"
         End
      End
      Begin VB.Menu mnuShops 
         Caption         =   "S&hops"
         Begin VB.Menu mnuShopRestock 
            Caption         =   "&Restock All Shops"
         End
      End
      Begin VB.Menu mnuTextblocks 
         Caption         =   "&Textblocks"
         Begin VB.Menu mnuStripChars 
            Caption         =   "&Strip characters off the end"
         End
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "&Users"
         Begin VB.Menu mnuDatabaseMerge 
            Caption         =   "&Merge Users"
         End
         Begin VB.Menu mnuUserModifyGang 
            Caption         =   "Change &Gang on All Users"
         End
         Begin VB.Menu mnuRetrainUsers 
            Caption         =   "&Retrain All Users"
         End
      End
      Begin VB.Menu mnusep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbilityEditor 
         Caption         =   "&Ability List Editor"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuControlRoomList 
         Caption         =   "Create Control Room List"
      End
      Begin VB.Menu mnuBuildMonsterIndex 
         Caption         =   "Create Monster/&Index List"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuMonsterNPCList 
         Caption         =   "Create N&PC/Room List"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuExpCalculator 
         Caption         =   "E&xp Calculator"
      End
      Begin VB.Menu mnuSwingCalculator 
         Caption         =   "S&wing Calculator"
      End
      Begin VB.Menu mnuUniversalModifier 
         Caption         =   "&Universal Modifier"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuQuestOrganizer 
         Caption         =   "&Quest Organizer"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnusep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompileUpdateMenu 
         Caption         =   "&Compile Update File"
         Begin VB.Menu mnuCompileUpdate 
            Caption         =   "Compile Full Update File"
            Shortcut        =   {F8}
         End
         Begin VB.Menu mnuCompileBlank 
            Caption         =   "Compile Blank Update File"
            Shortcut        =   +{F8}
         End
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCascade 
         Caption         =   "C&ascade Windows"
      End
      Begin VB.Menu mnuCloseAll 
         Caption         =   "&Close All Windows"
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuMinimizeWindows 
         Caption         =   "&Minimize All Windows"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuRestoreWindows 
         Caption         =   "&Restore All Windows"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnusep4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUnDock 
         Caption         =   "UnDock Current Window"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpGeneral 
         Caption         =   "&General Info"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpMessages 
         Caption         =   "&Messages"
      End
      Begin VB.Menu mnuHelpMonsters 
         Caption         =   "&Monsters"
      End
      Begin VB.Menu mnuHelpRooms 
         Caption         =   "&Rooms"
      End
      Begin VB.Menu mnuHelpTextblocks 
         Caption         =   "&Textblocks"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpChangeLog 
         Caption         =   "&Change Log"
      End
      Begin VB.Menu mnuHelpBug 
         Caption         =   "&Bug or Suggestion?"
      End
      Begin VB.Menu mnuAbout2 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuLimitedRightClick 
      Caption         =   "limited list right click"
      Visible         =   0   'False
      Begin VB.Menu mnuLimitedCopyLine 
         Caption         =   "Copy Line to clipboard"
      End
   End
   Begin VB.Menu mnuItemFindRightClick 
      Caption         =   "item find right click"
      Visible         =   0   'False
      Begin VB.Menu mnuItemFindCopyLine 
         Caption         =   "Copy Line to clipboard"
      End
   End
   Begin VB.Menu mnuMonsterIndexRightClick 
      Caption         =   "monster index right click"
      Visible         =   0   'False
      Begin VB.Menu mnuMonsterIndexCopyLine 
         Caption         =   "Copy Line to clipboard"
      End
   End
   Begin VB.Menu mnuMonsterNPCListRightClick 
      Caption         =   "monster npc lis right click"
      Visible         =   0   'False
      Begin VB.Menu mnuMonsterListCopyLine 
         Caption         =   "Copy Line to clipboard"
      End
   End
   Begin VB.Menu mnuMapUp 
      Caption         =   "MapUp"
      Visible         =   0   'False
      Begin VB.Menu mnuMapUpFollow 
         Caption         =   "Follow Up and Redraw"
      End
      Begin VB.Menu mnuMapUpRedraw 
         Caption         =   "Redraw from here"
      End
   End
   Begin VB.Menu mnuMapDown 
      Caption         =   "MapDown"
      Visible         =   0   'False
      Begin VB.Menu mnuMapDownFollow 
         Caption         =   "Follow Down and Redraw"
      End
      Begin VB.Menu mnuMapDownRedraw 
         Caption         =   "Redraw from here"
      End
   End
   Begin VB.Menu mnuMapUpDown 
      Caption         =   "MapUpDown"
      Visible         =   0   'False
      Begin VB.Menu mnuMapUpDownFollowUp 
         Caption         =   "Follow Up and Redraw"
      End
      Begin VB.Menu mnuMapUpDownFollowDown 
         Caption         =   "Follow Down and Redraw"
      End
      Begin VB.Menu mnuMapUpDownRedraw 
         Caption         =   "Redraw from here"
      End
   End
   Begin VB.Menu mnuMapEditorUp 
      Caption         =   "MapEditorUp"
      Visible         =   0   'False
      Begin VB.Menu mnuMapEditorUpFollow 
         Caption         =   "Follow Up and Redraw"
      End
      Begin VB.Menu mnuMapEditorUpRedraw 
         Caption         =   "Redraw from here"
      End
   End
   Begin VB.Menu mnuMapEditorDown 
      Caption         =   "MapEditorDown"
      Visible         =   0   'False
      Begin VB.Menu mnuMapEditorDownFollow 
         Caption         =   "Follow Down and Redraw"
      End
      Begin VB.Menu mnuMapEditorDownRedraw 
         Caption         =   "Redraw from here"
      End
   End
   Begin VB.Menu mnuMapEditorUpDown 
      Caption         =   "MapEditorUpDown"
      Visible         =   0   'False
      Begin VB.Menu mnuMapEditorUpDownFollowUp 
         Caption         =   "Follow Up and Redraw"
      End
      Begin VB.Menu mnuMapEditorUpDownFollowDown 
         Caption         =   "Follow Down and Redraw"
      End
      Begin VB.Menu mnuMapEditorUpDownRedraw 
         Caption         =   "Redraw from here"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit
Dim bStopProcess As Boolean
Dim bWarnedAboutCopy As Boolean

Private Sub MDIForm_Load()
On Error GoTo error:
Dim nTmp As Integer
Dim fso As FileSystemObject

sAppVersion = "v" & App.Major & "." & App.Minor _
    & IIf(App.Revision > 0, "." & App.Revision, "") _
    & IIf(WorksWithN = True, "n", "") _
    & IIf(WorksWithWG = True, "-WG", "")
sMenuCaption = App.Title & " " & sAppVersion

Load frmSplash
frmSplash.lblStatus.Caption = "Loading ..."
'frmSplash.Left = frmMain.Width / 4
'frmSplash.Top = frmMain.Height / 4
frmSplash.Show
DoEvents

Set fso = CreateObject("Scripting.FileSystemObject")

Call LockMenus

' & "." & App.Revision & " BETA"

Me.Caption = sMenuCaption

With EL1
    .FormInQuestion = Me
    .EnableLimiter = True
    .MINHEIGHT = 470
    .MINWIDTH = 580
    .CenterOnLoad = True
End With

If Right(App.Path, 1) = "\" Then
    sINIFileName = App.Path & "settings.ini"
Else
    sINIFileName = App.Path & "\settings.ini"
End If

If fso.FileExists(sINIFileName) = False Then Call CreateSettings

If Val(ReadINI("Windows", "MainNoMax")) = 0 Then
    Me.WindowState = vbMaximized
Else
    Me.Width = ReadINI("Windows", "MainWidth")
    Me.Height = ReadINI("Windows", "MainHeight")
End If

Call Startup

If ReadINI("Settings", "DisableWriting") = "1" Then
    mnuDisableWrite.Checked = True
    bDisableWriting = True
    Me.Caption = sMenuCaption & " -- *DB WRITING DISABLED*"
    stsStatusBar.Panels(1).Text = "*WRITING DISABLED*"
Else
    mnuDisableWrite.Checked = False
    bDisableWriting = False
    Me.Caption = sMenuCaption
    stsStatusBar.Panels(1).Text = ""
End If

DoEvents
Unload frmSplash
DoEvents

Call UnLockMenus

Me.Show
Me.Enabled = True
Set fso = Nothing

DoEvents
Exit Sub
error:
Call HandleError("Main_Load")
Resume Next
End Sub


Private Sub MDIForm_Resize()
On Error Resume Next
'    frmLogo.Left = Me.Width - 8700
'    frmLogo.Top = Me.Height - 2200
End Sub


Private Sub mnuBuildMonsterIndex_Click()
On Error GoTo error:

If FormIsLoaded("frmMonsterIndex") Then
    frmMonsterIndex.SetFocus
Else
    frmMonsterIndex.Show
End If

Exit Sub
error:
Call HandleError
End Sub

Private Sub mnuCascade_Click()
Me.Arrange (vbCascade)
End Sub

Private Sub mnuChangeRoomCallLetters_Click()
On Error GoTo error:
Dim sLetters As String, nRoom As Long, nStatus As Integer, x As Integer, sFile As String

sLetters = InputBox("This will change every text string in the rooms database that refers to files " _
        & "with the specified call letters below to your current call letter setting." _
        & vbCrLf & vbCrLf & "So if you specified 'CC' and your current setting is 'BB', it would change occurances like " _
        & "'WCCMAP01.ANS' to 'WBBMAP01.ANS' and 'WCC89615.HSE' to 'WBB89615.HSE'. " _
        & vbCrLf & vbCrLf & "(It scanes room descriptions, room name, and ansi map)" _
        & vbCrLf & vbCrLf & "Enter call letters to change FROM: " _
        & vbCrLf & "(They will be changed to your current setting of " & strDatCallLetters & ")", "Change Call Letters", "cc")

If sLetters = "" Then Exit Sub

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then MsgBox "Error getting first room, error: " & BtrieveErrorCode(nStatus): Exit Sub

frmProgressBar.sCaption = "Changing room call letters"
frmProgressBar.lblCaption = "Changing room call letters ..."
frmProgressBar.cmdCancel.Enabled = True
frmProgressBar.ProgressBar.Value = 0
frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_MP
frmProgressBar.lblPanel(1).Caption = ""
frmProgressBar.Show
frmMain.Enabled = False
DoEvents

nStatus = BTRCALL(BSTAT, RoomPosBlock, DBStatDatabuf, Len(Roomdatabuf), 0, KEY_BUF_LEN, 0)
DBStatRowToStruct DBStatDatabuf.buf
Call frmProgressBar.SetRange(DBStat.nRecords)
nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)

bStopProcess = False
nRoom = 0
Do While nStatus = 0 And bStopProcess = False
    RoomRowToStruct Roomdatabuf.buf

    nRoom = nRoom + 1
    frmProgressBar.lblPanel(1).Caption = nRoom
    Call frmProgressBar.IncreaseProgress
    
    Roomrec.AnsiMap = ChangeCallLetters(sLetters, Roomrec.AnsiMap)
    Roomrec.Name = ChangeCallLetters(sLetters, Roomrec.Name)
    
    For x = 0 To 6
        Roomrec.Desc(x) = ChangeCallLetters(sLetters, Roomrec.Desc(x))
    Next x
    
    nStatus = UpdateRoom
    If Not nStatus = 0 Then Exit Do
    
    nStatus = BTRCALL(BGETNEXT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

If bStopProcess = True Then GoTo kill:

If Not nStatus = 9 And Not nStatus = 0 Then
    MsgBox "Abnormal Exit: " & BtrieveErrorCode(nStatus), vbOKOnly + vbExclamation
Else
    frmProgressBar.ProgressBar.Value = frmProgressBar.ProgressBar.Max
    DoEvents
    MsgBox "Complete!", vbInformation
End If

kill:
frmMain.Enabled = True
Unload frmProgressBar
If Me.Visible Then Me.SetFocus
Exit Sub
error:
Call HandleError
On Error Resume Next
frmMain.Enabled = True
Unload frmProgressBar

End Sub

Private Sub mnuCloseAll_Click()
UnloadForms (Me.Name)
End Sub

Private Sub mnuCompileBlank_Click()
If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
Call CompileUpdatefile(True)
End Sub

Private Sub mnuControlRoomList_Click()
On Error GoTo error:

If FormIsLoaded("frmRoomControlRoomList") Then
    frmRoomControlRoomList.SetFocus
Else
    frmRoomControlRoomList.Show
End If

Exit Sub
error:
Call HandleError
End Sub

Private Sub mnuDatabaseIndexChange_Click()
Unload frmMonsterIndexChanger
Load frmMonsterIndexChanger
End Sub

Private Sub mnuDatabaseMerge_Click()
Unload frmDatabaseMerge
Load frmDatabaseMerge
End Sub

Private Sub mnuDeleteBufferRooms_Click()
On Error GoTo error:
Dim nMap As Long, nRoom As Long, nStatus As Integer, x As Integer, sFile As String
Dim fso As FileSystemObject, ts As TextStream, nPrevRoom As Long, nPrevMap As Long
Dim nLastRoom As Long, nYesNo As Integer, bAll As Boolean, nMaxRooms As Long

nMap = Val(InputBox("This will delete any rooms on the chosen map which name begins with ""Buffer Room""." _
        & vbCrLf & "Enter a value of -1 to delete buffer rooms on all maps." _
        & vbCrLf & vbCrLf & "Enter map number to delete buffer rooms on:", "Delete buffer rooms", 1))

If nMap = 0 Then Exit Sub
If nMap = -1 Then bAll = True
If nMap < -1 Then Exit Sub

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then MsgBox "Error getting first room, error: " & BtrieveErrorCode(nStatus): Exit Sub

Set fso = CreateObject("Scripting.FileSystemObject")

If Right(App.Path, 1) = "\" Then
    sFile = App.Path & "NMR-Log_RoomBufDelete.txt"
Else
    sFile = App.Path & "\NMR-Log_RoomBufDelete.txt"
End If

If fso.FileExists(sFile) Then Call fso.DeleteFile(sFile, True)
Set ts = fso.OpenTextFile(sFile, ForWriting, True)

ts.WriteLine ("Buffer room delete job started " & Date & " @ " & Time)
ts.WriteBlankLines (1)

frmProgressBar.sCaption = "Deleting buffer rooms"
frmProgressBar.lblCaption = "Deleting buffer rooms ..."
frmProgressBar.cmdCancel.Enabled = True
frmProgressBar.ProgressBar.Value = 0
frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_MP
frmProgressBar.lblPanel(1).Caption = ""
frmProgressBar.Show
frmMain.Enabled = False
DoEvents

bStopProcess = False

nStatus = BTRCALL(BSTAT, RoomPosBlock, DBStatDatabuf, Len(Roomdatabuf), 0, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    nMaxRooms = 30000
Else
    DBStatRowToStruct DBStatDatabuf.buf
    nMaxRooms = DBStat.nRecords
End If
Call frmProgressBar.SetRange(nMaxRooms)

nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
bStopProcess = False
Do While nStatus = 0 And bStopProcess = False
    RoomRowToStruct Roomdatabuf.buf
    
    If Roomrec.MapNumber = nMap Or bAll Then
        If UCase(Left(Roomrec.Name, 11)) = "BUFFER ROOM" Then
            nStatus = BTRCALL(BDELETE, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
            If Not nStatus = 0 Then
                MsgBox "Error deleting Map " & Roomrec.MapNumber & " Room " & Roomrec.RoomNumber & ", Error: " & BtrieveErrorCode(nStatus)
                GoTo kill:
            End If
            
            ts.WriteLine "Deleted Room: Map " & Roomrec.MapNumber & " Room " & Roomrec.RoomNumber
        End If
        
        Call frmProgressBar.IncreaseProgress
        frmProgressBar.lblPanel(1).Caption = Roomrec.MapNumber & "/" & Roomrec.RoomNumber
    End If
    
    nStatus = BTRCALL(BGETNEXT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

If bStopProcess = True Then
    ts.WriteLine "...canceled by user"
    GoTo kill:
End If

If Not nStatus = 0 And Not nStatus = 9 Then
    ts.WriteLine "Exited because of btrieve error: " & BtrieveErrorCode(nStatus)
    MsgBox "Abnormal Exit: " & BtrieveErrorCode(nStatus), vbOKOnly + vbExclamation
Else
    ts.WriteBlankLines (1)
    ts.WriteLine ("Complete: " & Date & " @ " & Time)
    frmProgressBar.ProgressBar.Value = frmProgressBar.ProgressBar.Max
    DoEvents
    nYesNo = MsgBox("Complete, view log?", vbYesNo + vbQuestion, "View?")
    If nYesNo = vbYes Then Call ShellExecute(0&, "open", sFile, vbNullString, vbNullString, vbNormalFocus)
    DoEvents
End If

kill:
On Error Resume Next
ts.Close
Set ts = Nothing
Set fso = Nothing
frmMain.Enabled = True
Unload frmProgressBar
If Me.Visible Then Me.SetFocus
Exit Sub
error:
Call HandleError
Resume kill:
End Sub

Private Sub mnuDisableWrite_Click()
If mnuDisableWrite.Checked = False Then
    mnuDisableWrite.Checked = True
    bDisableWriting = True
    Me.Caption = sMenuCaption & " -- *DB WRITING DISABLED*"
    stsStatusBar.Panels(1).Text = "*WRITING DISABLED*"
Else
    mnuDisableWrite.Checked = False
    bDisableWriting = False
    Me.Caption = sMenuCaption
    stsStatusBar.Panels(1).Text = ""
End If
End Sub

Private Sub mnuDivideExp_Click()
On Error GoTo error:
Dim nStatus As Integer, nYesNo As Long, nTemp As Double, nDivisor As Integer
Dim nBase As Double, nMulti As Double, nMultiMax As Double, x As Double

nDivisor = Val(InputBox("Divide monster exp by how many times (2-20)?" & vbCrLf & "(Minimum exp will be 1)", "Monster EXP Divider", "2"))
If nDivisor <= 0 Then Exit Sub
If nDivisor > 20 Then nDivisor = 20
If nDivisor < 2 Then nDivisor = 2

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

nYesNo = MsgBox("Are you sure you want to DIVIDE monster EXP by " & nDivisor & "?", vbYesNo + vbQuestion + vbDefaultButton2)
If nYesNo <> vbYes Then Exit Sub
    
UnloadForms (Me.Name)
frmProgressBar.sCaption = "Dividing Monster EXP"
frmProgressBar.lblCaption.Caption = "Dividing Monster EXP ..."
frmProgressBar.cmdCancel.Enabled = True
frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_KNMSR
frmProgressBar.Show
frmMain.Enabled = False
DoEvents

nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
    GoTo finish:
Else
    nStatus = BTRCALL(BSTAT, MonsterPosBlock, DBStatDatabuf, Len(Monsterdatabuf), 0, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        Call frmProgressBar.SetRange(30000)
    Else
        DBStatRowToStruct DBStatDatabuf.buf
        Call frmProgressBar.SetRange(DBStat.nRecords)
    End If
    nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
End If
                    
bStopProcess = False
Do While nStatus = 0 And bStopProcess = False
    MonsterRowToStruct Monsterdatabuf.buf
    frmProgressBar.lblPanel(1).Caption = Monsterrec.Number
    Call frmProgressBar.IncreaseProgress
    
    If eDatFileVersion >= v111j Then
        If Monsterrec.ExpMulti = 1 Or Monsterrec.ExpMulti = 0 Then
            nBase = SLong2ULong(Monsterrec.Experience)
            nBase = Round(nBase / nDivisor)
            If nBase <= 0 Then nBase = 1
            If nBase > 2147483646 Then nBase = 2147483646
            Monsterrec.Experience = ULong2SLong(nBase)
        Else
            nBase = SLong2ULong(Monsterrec.Experience) * SLong2ULong(Monsterrec.ExpMulti)
            nBase = Round(nBase / nDivisor)
            If nBase > 2147483646 Then nBase = 2147483646
            
tryagain:
            If nBase > 100000 Then
                nMultiMax = 20
                For x = 20 To 32767
                    If x * 65538 >= nBase Then
                        nMultiMax = x
                        Exit For
                    End If
                Next x
                
                nMulti = 1
                For x = 3 To nMultiMax
                    nTemp = nBase Mod x
                    If nTemp = 0 Then nMulti = x
                Next x
                
                If nMulti = 1 Then
                    nBase = nBase - 1
                    GoTo tryagain:
                End If
                
                nBase = nBase / nMulti
            Else
                nMulti = 1
            End If
            
            If nBase <= 0 Then nBase = 1
            
            Monsterrec.Experience = ULong2SLong(nBase)
            Monsterrec.ExpMulti = ULong2SLong(nMulti)
        End If
    Else
        nBase = SLong2ULong(Monsterrec.Experience)
        nBase = Round(nBase / nDivisor)
        If nBase <= 0 Then nBase = 1
        Monsterrec.Experience = ULong2SLong(nBase)
    End If
    
    nStatus = UpdateMonster
    If Not nStatus = 0 Then
        MsgBox "Update record Error, " & BtrieveErrorCode(nStatus)
        GoTo finish:
    End If
    nStatus = BTRCALL(BGETNEXT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
    DoEvents
Loop

If bStopProcess = True Then GoTo finish:

If Not nStatus = 9 And Not nStatus = 0 Then
    MsgBox "Abnormal Exit: " & BtrieveErrorCode(nStatus), vbOKOnly + vbExclamation
Else
    frmProgressBar.ProgressBar.Value = frmProgressBar.ProgressBar.Max
    DoEvents
    MsgBox "Complete!", vbInformation
End If

finish:
frmMain.Enabled = True
Unload frmProgressBar
If Me.Visible Then Me.SetFocus
Exit Sub
error:
Call HandleError
frmMain.Enabled = True
Unload frmProgressBar
End Sub
Private Sub mnuAbilityEditor_Click()
On Error GoTo error:

If bAbilityDBOpen = False Then
    MsgBox "The ability database was never opened successfully, reload program.", vbOKOnly + vbExclamation
    Exit Sub
End If

Unload frmAbilityEdit
Load frmAbilityEdit

Exit Sub
error:
Call HandleError
End Sub

Public Sub mnuEditAction_Click()
On Error GoTo error:

If FormIsLoaded("frmAction") = True Then
    Call CopyActionForm
Else
    Unload frmAction
    Load frmAction
End If

Exit Sub
error:
Call HandleError
End Sub

Public Sub mnuEditBankbooks_Click()

Unload frmBank
Load frmBank

End Sub

Private Sub mnuDatabaseDeleter_Click()
Unload frmDatabaseDeleter
Load frmDatabaseDeleter
End Sub

Private Sub mnuDatabaseImporter_Click()
Unload frmDatabaseImport
Load frmDatabaseImport
End Sub

Private Sub mnuDatabaseExporter_Click()
Unload frmDatabaseExport
Load frmDatabaseExport
End Sub

Private Sub mnuEditGangs_Click()
frmGang.Show
frmGang.SetFocus
End Sub

Public Sub mnuEditRace_Click()
On Error GoTo error:

If FormIsLoaded("frmRace") = True Then
    Call CopyRaceForm
Else
    Unload frmRace
    Load frmRace
End If

Exit Sub
error:
Call HandleError
End Sub
Public Sub mnuEditClass_Click()
On Error GoTo error:

If FormIsLoaded("frmClass") = True Then
    Call CopyClassForm
Else
    Unload frmClass
    Load frmClass
End If

Exit Sub
error:
Call HandleError
End Sub
Public Sub mnuEditSpell_Click()
On Error GoTo error:

If FormIsLoaded("frmSpell") = True Then
    Call CopySpellForm
Else
    Unload frmSpell
    Load frmSpell
End If

Exit Sub
error:
Call HandleError
End Sub
Public Sub mnuEditMonster_Click()
On Error GoTo error:

If FormIsLoaded("frmMonster") = True Then
    Call CopyMonsterForm
Else
    Unload frmMonster
    Load frmMonster
End If

Exit Sub
error:
Call HandleError
End Sub
Public Sub mnuEditItem_Click()
On Error GoTo error:

If FormIsLoaded("frmItem") = True Then
    Call CopyItemForm
Else
    Unload frmItem
    Load frmItem
End If

Exit Sub
error:
Call HandleError
End Sub
Public Sub mnuEditTextblock_Click()
On Error GoTo error:

If FormIsLoaded("frmTextblock") = True Then
    Call CopyTextblockForm
Else
    Unload frmTextblock
    Load frmTextblock
End If

Exit Sub
error:
Call HandleError
End Sub
Public Sub mnuEditUser_Click()
On Error GoTo error:

If FormIsLoaded("frmUser") = True Then
    Call CopyUserForm
Else
    Unload frmUser
    Load frmUser
End If

Exit Sub
error:
Call HandleError
End Sub
Public Sub mnuEditShop_Click()
On Error GoTo error:

If FormIsLoaded("frmShop") = True Then
    Call CopyShopForm
Else
    Unload frmShop
    Load frmShop
End If

Exit Sub
error:
Call HandleError
End Sub
Public Sub mnuEditRoom_Click()
On Error GoTo error:

If FormIsLoaded("frmRoom") = True Then
    If bWarnedAboutCopy = False Then
        MsgBox "NOTE: The Map and Map Editor will only interact with the original (non-copied) room editor.", vbInformation + vbOKOnly
        bWarnedAboutCopy = True
    End If
    Call CopyRoomForm
Else
    Unload frmRoom
    Load frmRoom
End If

Exit Sub
error:
Call HandleError
End Sub
Public Sub mnuEditMessage_Click()
On Error GoTo error:

If FormIsLoaded("frmMessage") = True Then
    Call CopyMessageForm
Else
    Unload frmMessage
    Load frmMessage
End If

Exit Sub
error:
Call HandleError
End Sub

Private Sub mnuExpCalculator_Click()
On Error GoTo error:

frmExpCalc.Show
frmExpCalc.SetFocus

Exit Sub
error:
Call HandleError
End Sub

Private Sub mnuFindItem_Click()
'On Error GoTo error:
'Dim nStatus As Integer
'Dim fso As FileSystemObject, nItem As Integer, x As Integer
'Dim sFile As String, oFile As TextStream, nRecNum As Long
'
'nItem = Val(InputBox("This will search through every room for an item you specify and dump to a file (NMR-ItemFind.txt) all the rooms it finds the item in." & vbCrLf & vbCrLf & "Enter Item Number to Search for:", "Enter Item Number to Search for:", 0))
'If nItem = 0 Then Exit Sub
'
'UnloadForms (Me.Name)
'
'Set fso = CreateObject("Scripting.FileSystemObject")
'
'If Right(App.Path, 1) = "\" Then
'    sFile = App.Path & "NMR-Log_FindItem.txt"
'Else
'    sFile = App.Path & "\NMR-Log_FindItem.txt"
'End If
'
'If fso.FileExists(sFile) Then fso.DeleteFile sFile, True
'
'nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
'If Not nStatus = 0 Then MsgBox "Could not get first room, Error: " & BtrieveErrorCode(nStatus), vbExclamation: Exit Sub
'
'nStatus = BTRCALL(BSTAT, RoomPosBlock, DBStatDatabuf, Len(Roomdatabuf), 0, KEY_BUF_LEN, 0)
'If Not nStatus = 0 Then
'    Call frmProgressBar.SetRange(25000)
'Else
'    DBStatRowToStruct DBStatDatabuf.buf
'    Call frmProgressBar.SetRange(DBStat.nRecords)
'End If
'
'frmProgressBar.sCaption = "Find Room Item"
'frmProgressBar.lblCaption.Caption = "Searching ..."
'frmProgressBar.cmdCancel.Enabled = True
'frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_MP
'frmProgressBar.Show
'frmMain.Enabled = False
'DoEvents
'
'Set oFile = fso.OpenTextFile(sFile, ForWriting, True)
'oFile.WriteLine "Find Item Started @ " & Time & " on " & Date
'oFile.WriteBlankLines (1)
'
'bStopProcess = False
'nRecNum = 0
'Do While nStatus = 0 And bStopProcess = False
'
'    RoomRowToStruct Roomdatabuf.buf
'    nRecNum = nRecNum + 1
'    frmProgressBar.lblPanel(1).Caption = nRecNum
'    Call frmProgressBar.IncreaseProgress
'
'    For x = 0 To 14
'        If Roomrec.RoomItems(x) = nItem Then GoTo found:
'        If Roomrec.InvisItems(x) = nItem Then GoTo found:
'    Next
'    For x = 15 To 16
'        If Roomrec.RoomItems(x) = nItem Then GoTo found:
'    Next
'
'    GoTo Skip:
'found:
'    oFile.WriteLine "Item found in Room " & Roomrec.RoomNumber & ", Map " & Roomrec.MapNumber
'
'Skip:
'    nStatus = BTRCALL(BGETNEXT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
'    DoEvents
'Loop
'
'If bStopProcess = True Then GoTo kill:
'
'If Not nStatus = 9 And Not nStatus = 0 Then
'    MsgBox "Abnormal Exit: " & BtrieveErrorCode(nStatus), vbOKOnly + vbExclamation
'Else
'    oFile.WriteBlankLines (1)
'    oFile.WriteLine "Find Item Finished @ " & Time & " on " & Date
'    oFile.Close
'    frmProgressBar.ProgressBar.Value = frmProgressBar.ProgressBar.Max
'    DoEvents
'    nItem = MsgBox("Complete, view results?", vbYesNo + vbQuestion + vbDefaultButton1)
'    If nItem = vbYes Then Call ShellExecute(0&, "open", sFile, vbNullString, vbNullString, vbNormalFocus)
'End If
'
'kill:
'On Error Resume Next
'oFile.Close
'frmMain.Enabled = True
'Unload frmProgressBar
'Set oFile = Nothing
'Set fso = Nothing
'If Me.Visible Then Me.SetFocus
'Exit Sub
'error:
'Call HandleError
'Resume kill:

frmItemFind.Show

End Sub
Public Sub CancelProcess()
    bStopProcess = True
    DoEvents
End Sub

Private Sub mnuFixMonsterUses_Click()
Dim nStatus As Integer, sFile As String, bUpdateRec As Boolean
Dim nYesNo As Integer, x As Integer, y As Integer, nItem As Long
Dim fso As FileSystemObject, ts As TextStream

On Error GoTo error:

nYesNo = MsgBox("This will cross reference every dropped item's " _
    & "actual number of uses and set the number of uses on drop to that number. " _
    & "If the item has 0 or less max uses, the uses on drop will be set to -1.  It would " _
    & "be wise to run the similar tool under items before running this one.  " _
    & "This will fix item stacking problems for items that don't have uses." _
    & vbCrLf & vbCrLf _
    & "Continue?", vbYesNo + vbQuestion + vbDefaultButton2)
If nYesNo <> vbYes Then Exit Sub

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

UnloadForms (Me.Name)
frmProgressBar.sCaption = "Fixing Number of Uses on Monster Item Drops"
frmProgressBar.lblCaption.Caption = "Working ..."
frmProgressBar.cmdCancel.Enabled = True
frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_KNMSR
frmProgressBar.Show
frmMain.Enabled = False
DoEvents

nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
    GoTo out:
Else
    nStatus = BTRCALL(BSTAT, MonsterPosBlock, DBStatDatabuf, Len(Monsterdatabuf), 0, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        Call frmProgressBar.SetRange(30000)
    Else
        DBStatRowToStruct DBStatDatabuf.buf
        Call frmProgressBar.SetRange(DBStat.nRecords)
    End If
    nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
End If

Set fso = CreateObject("Scripting.FileSystemObject")

sFile = IIf(Right(App.Path, 1) = "\", _
    App.Path & "NMR-Log_MonsterDrops.txt", _
    App.Path & "\NMR-Log_MonsterDrops.txt")

If fso.FileExists(sFile) Then fso.DeleteFile sFile, True

Set ts = fso.OpenTextFile(sFile, ForWriting, True)

ts.WriteLine ("Job started " & Date & " @ " & Time)
ts.WriteBlankLines (1)

bStopProcess = False
Do While nStatus = 0 And bStopProcess = False
    bUpdateRec = False
    MonsterRowToStruct Monsterdatabuf.buf
    frmProgressBar.lblPanel(1).Caption = Monsterrec.Number
    Call frmProgressBar.IncreaseProgress
    
    For x = 0 To 9
        If Monsterrec.ItemNumber(x) > 0 Then
            nItem = Monsterrec.ItemNumber(x)
            nStatus = BTRCALL(BGETEQUAL, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), nItem, KEY_BUF_LEN, 0)
            If nStatus = 0 Then
                ItemRowToStruct Itemdatabuf.buf
                If Monsterrec.ItemUses(x) <> Itemrec.Uses Then
                    bUpdateRec = True
                    ts.WriteLine ("Monster #" & Monsterrec.Number _
                        & " [" & ClipNull(Monsterrec.Name) & "] -- Item #" _
                        & Monsterrec.ItemNumber(x) _
                        & " [" & ClipNull(Itemrec.Name) & "] -- Set to " _
                        & Itemrec.Uses & " uses on drop (was " & Monsterrec.ItemUses(x) & ")")
                    Monsterrec.ItemUses(x) = Itemrec.Uses
                End If
            Else
                ts.WriteLine ("Monster #" & Monsterrec.Number _
                    & " [" & ClipNull(Monsterrec.Name) & "] -- Item #" _
                    & Monsterrec.ItemNumber(x) & " not found")
            End If
        End If
    Next x
    
    If bUpdateRec Then
        nStatus = UpdateMonster
        If Not nStatus = 0 Then
            MsgBox "Update record Error, " & BtrieveErrorCode(nStatus)
            GoTo out:
        End If
    End If
    nStatus = BTRCALL(BGETNEXT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
    DoEvents
Loop

If bStopProcess = True Then GoTo out:

ts.WriteBlankLines (1)
ts.WriteLine ("Job finished " & Date & " @ " & Time)

If Not nStatus = 9 And Not nStatus = 0 Then
    MsgBox "Abnormal Exit: " & BtrieveErrorCode(nStatus), vbOKOnly + vbExclamation
Else
    frmProgressBar.ProgressBar.Value = frmProgressBar.ProgressBar.Max
    DoEvents
    nYesNo = MsgBox("Complete, view log?", vbYesNo + vbQuestion + vbDefaultButton1, "View Log?")
    If nYesNo = vbYes Then Call ShellExecute(0&, "open", sFile, vbNullString, vbNullString, vbNormalFocus)
End If

out:
On Error Resume Next
ts.Close
Set ts = Nothing
Set fso = Nothing
frmMain.Enabled = True
Unload frmProgressBar
If Me.Visible Then Me.SetFocus
Exit Sub
error:
Call HandleError("mnuFixMonsterUses_Click")
Resume out:
End Sub

Private Sub mnuHelpBug_Click()

MsgBox "Chances are that I DO NOT KNOW about a bug you are encountering." & vbCrLf _
    & "Please email ANY feature request, bug report, or suggestion" & vbCrLf _
    & "that you may have: syntax53@mudinfo.net", vbInformation
    
End Sub

Private Sub mnuHelpChangeLog_Click()
Unload frmHelpChangeLog
Load frmHelpChangeLog
End Sub

Private Sub mnuItemFindCopyLine_Click()
Call frmItemFind.CopyLine
End Sub

Private Sub mnuItemsFixUses_Click()
Dim nStatus As Integer, sFile As String, bUpdateRec As Boolean
Dim nYesNo As Integer, x As Integer, y As Integer
Dim fso As FileSystemObject, ts As TextStream
On Error GoTo error:

nYesNo = MsgBox("This will go through ever item and verify that the uses are set " _
    & "correctly.  If an item has 0 uses it will be set to -1.  If an item has >0 " _
    & "uses and is not set to destroy after the uses expire, the recharge ability value " _
    & "will be verified against the max uses.  If an item in this case has no recharge ability, " _
    & "a line will be logged in the log file for further attention.  " _
    & "This will fix item stacking problems for items that don't have uses and should " _
    & "be run along with the item drop fix listed under monsters in the tools." _
    & vbCrLf & vbCrLf _
    & "Continue?", vbYesNo + vbQuestion + vbDefaultButton2)
If nYesNo <> vbYes Then Exit Sub

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

UnloadForms (Me.Name)
frmProgressBar.sCaption = "Fixing Number of Uses on Items"
frmProgressBar.lblCaption.Caption = "Working ..."
frmProgressBar.cmdCancel.Enabled = True
frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_ITEMS
frmProgressBar.Show
frmMain.Enabled = False
DoEvents

Set fso = CreateObject("Scripting.FileSystemObject")

sFile = IIf(Right(App.Path, 1) = "\", _
    App.Path & "NMR-Log_ItemUseFix.txt", _
    App.Path & "\NMR-Log_ItemUseFix.txt")

If fso.FileExists(sFile) Then fso.DeleteFile sFile, True

Set ts = fso.OpenTextFile(sFile, ForWriting, True)

ts.WriteLine ("Job started " & Date & " @ " & Time)
ts.WriteBlankLines (1)

nStatus = BTRCALL(BGETFIRST, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
    GoTo out:
Else
    nStatus = BTRCALL(BSTAT, ItemPosBlock, DBStatDatabuf, Len(Itemdatabuf), 0, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        Call frmProgressBar.SetRange(30000)
    Else
        DBStatRowToStruct DBStatDatabuf.buf
        Call frmProgressBar.SetRange(DBStat.nRecords)
    End If
    nStatus = BTRCALL(BGETFIRST, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
End If

bStopProcess = False
Do While nStatus = 0 And bStopProcess = False
    bUpdateRec = False
    ItemRowToStruct Itemdatabuf.buf
    frmProgressBar.lblPanel(1).Caption = Itemrec.Number
    Call frmProgressBar.IncreaseProgress
    
    If Itemrec.Uses = 0 Then
        bUpdateRec = True
        Itemrec.Uses = -1
        ts.WriteLine ("Item #" & Itemrec.Number _
            & " [" & ClipNull(Itemrec.Name) & "] -- Set to -1 uses (was zero)")
    ElseIf Itemrec.Uses > 0 And Itemrec.RetainAfterUses = 1 Then
        For x = 0 To 19
            If Itemrec.AbilityA(x) = 121 Then
                If Itemrec.AbilityB(x) <> Itemrec.Uses Then
                    bUpdateRec = True
                    ts.WriteLine ("Item #" & Itemrec.Number _
                        & " [" & ClipNull(Itemrec.Name) & "] -- Recharge ability set to " _
                        & Itemrec.Uses & " uses (was " & Itemrec.AbilityB(x) & ")")
                    Itemrec.AbilityB(x) = Itemrec.Uses
                End If
                GoTo cont:
            End If
        Next x
        ts.WriteLine (">> Item #" & Itemrec.Number _
            & " [" & ClipNull(Itemrec.Name) & "] -- Item has " & Itemrec.Uses & " uses, but no recharge ability!")
    End If
cont:
    If bUpdateRec Then
        nStatus = UpdateItem
        If Not nStatus = 0 Then
            MsgBox "Update record Error, " & BtrieveErrorCode(nStatus)
            GoTo out:
        End If
    End If
    nStatus = BTRCALL(BGETNEXT, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
    DoEvents
Loop

If bStopProcess = True Then GoTo out:

ts.WriteBlankLines (1)
ts.WriteLine ("Job finished " & Date & " @ " & Time)

If Not nStatus = 9 And Not nStatus = 0 Then
    MsgBox "Abnormal Exit: " & BtrieveErrorCode(nStatus), vbOKOnly + vbExclamation
Else
    frmProgressBar.ProgressBar.Value = frmProgressBar.ProgressBar.Max
    DoEvents
    nYesNo = MsgBox("Complete, view log?", vbYesNo + vbQuestion + vbDefaultButton1, "View Log?")
    If nYesNo = vbYes Then Call ShellExecute(0&, "open", sFile, vbNullString, vbNullString, vbNormalFocus)
End If

out:
On Error Resume Next
ts.Close
Set ts = Nothing
Set fso = Nothing
frmMain.Enabled = True
Unload frmProgressBar
If Me.Visible Then Me.SetFocus
Exit Sub
error:
Call HandleError("mnuItemsFixUses_Click")
Resume out:
End Sub

Private Sub mnuLimitedCopyLine_Click()
Call frmLimitedList.CopyLine
End Sub

Private Sub mnuLimitedItemList_Click()
Unload frmLimitedList
frmLimitedList.Show
End Sub




Private Sub mnuMapUpDownFollowDown_Click()
Call mnuMapDownFollow_Click
End Sub

Private Sub mnuMapUpDownFollowUp_Click()
Call mnuMapUpFollow_Click
End Sub

Private Sub mnuMapUpDownRedraw_Click()
If FormIsLoaded("frmRoom") = False Then frmRoom.Show
Call frmRoom.GotoRoom(frmMap.nLastMapClick, frmMap.nLastRoomClick, True)
Call frmMap.StartMapping
End Sub

Private Sub mnuMapUpRedraw_Click()
If FormIsLoaded("frmRoom") = False Then frmRoom.Show
Call frmRoom.GotoRoom(frmMap.nLastMapClick, frmMap.nLastRoomClick, True)
Call frmMap.StartMapping
End Sub

Private Sub mnuMapDownRedraw_Click()
If FormIsLoaded("frmRoom") = False Then frmRoom.Show
Call frmRoom.GotoRoom(frmMap.nLastMapClick, frmMap.nLastRoomClick, True)
Call frmMap.StartMapping
End Sub
Private Sub mnuMapUpFollow_Click()
On Error GoTo error:
Dim nStatus As Integer, nMap As Long, nRoom As Long

RoomKeyStruct.MapNum = frmMap.nLastMapClick
RoomKeyStruct.RoomNum = frmMap.nLastRoomClick

nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then MsgBox "BGETEQUAL, Room, Error: " & BtrieveErrorCode(nStatus): Exit Sub

RoomRowToStruct Roomdatabuf.buf

If Roomrec.RoomType(8) = 8 Then
    nMap = Roomrec.Para1(8)
Else
    nMap = Roomrec.MapNumber
End If

nRoom = Roomrec.RoomExit(8)

If FormIsLoaded("frmRoom") = False Then frmRoom.Show
Call frmRoom.GotoRoom(nMap, nRoom, True)
Call frmMap.StartMapping

Exit Sub
error:
Call HandleError
End Sub

Private Sub mnuMapDownFollow_Click()
On Error GoTo error:
Dim nStatus As Integer, nMap As Long, nRoom As Long

RoomKeyStruct.MapNum = frmMap.nLastMapClick
RoomKeyStruct.RoomNum = frmMap.nLastRoomClick

nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then MsgBox "BGETEQUAL, Room, Error: " & BtrieveErrorCode(nStatus): Exit Sub

RoomRowToStruct Roomdatabuf.buf
    
If Roomrec.RoomType(9) = 8 Then
    nMap = Roomrec.Para1(9)
Else
    nMap = Roomrec.MapNumber
End If

nRoom = Roomrec.RoomExit(9)

If FormIsLoaded("frmRoom") = False Then frmRoom.Show
Call frmRoom.GotoRoom(nMap, nRoom, True)
Call frmMap.StartMapping
    
Exit Sub
error:
Call HandleError
End Sub


Private Sub mnuMapEditorDownFollow_Click()
On Error GoTo error:
Dim nStatus As Integer, nMap As Long, nRoom As Long

RoomKeyStruct.MapNum = frmMapEditor.nLastMapClick
RoomKeyStruct.RoomNum = frmMapEditor.nLastRoomClick

nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then MsgBox "BGETEQUAL, Room, Error: " & BtrieveErrorCode(nStatus): Exit Sub

RoomRowToStruct Roomdatabuf.buf
    
If Roomrec.RoomType(9) = 8 Then
    nMap = Roomrec.Para1(9)
Else
    nMap = Roomrec.MapNumber
End If

nRoom = Roomrec.RoomExit(9)

If FormIsLoaded("frmRoom") = False Then frmRoom.Show
Call frmRoom.GotoRoom(nMap, nRoom, True)
Call frmMapEditor.StartMapping
    
Exit Sub
error:
Call HandleError
End Sub

Private Sub mnuMapEditorDownRedraw_Click()

If FormIsLoaded("frmRoom") = False Then frmRoom.Show
Call frmRoom.GotoRoom(frmMapEditor.nLastMapClick, frmMapEditor.nLastRoomClick, True)
Call frmMapEditor.StartMapping

End Sub

Private Sub mnuMapEditorUpDownFollowDown_Click()
Call mnuMapEditorDownFollow_Click
End Sub

Private Sub mnuMapEditorUpDownFollowUp_Click()
Call mnuMapEditorUpFollow_Click
End Sub

Private Sub mnuMapEditorUpRedraw_Click()
If FormIsLoaded("frmRoom") = False Then frmRoom.Show
Call frmRoom.GotoRoom(frmMapEditor.nLastMapClick, frmMapEditor.nLastRoomClick, True)
Call frmMapEditor.StartMapping
End Sub

Private Sub mnuMapEditorUpDownRedraw_Click()
If FormIsLoaded("frmRoom") = False Then frmRoom.Show
Call frmRoom.GotoRoom(frmMapEditor.nLastMapClick, frmMapEditor.nLastRoomClick, True)
Call frmMapEditor.StartMapping
End Sub

Private Sub mnuMapEditorUpFollow_Click()
On Error GoTo error:
Dim nStatus As Integer, nMap As Long, nRoom As Long

RoomKeyStruct.MapNum = frmMapEditor.nLastMapClick
RoomKeyStruct.RoomNum = frmMapEditor.nLastRoomClick

nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then MsgBox "BGETEQUAL, Room, Error: " & BtrieveErrorCode(nStatus): Exit Sub

RoomRowToStruct Roomdatabuf.buf

If Roomrec.RoomType(8) = 8 Then
    nMap = Roomrec.Para1(8)
Else
    nMap = Roomrec.MapNumber
End If

nRoom = Roomrec.RoomExit(8)

If FormIsLoaded("frmRoom") = False Then frmRoom.Show
Call frmRoom.GotoRoom(nMap, nRoom, True)
Call frmMapEditor.StartMapping

Exit Sub
error:
Call HandleError
End Sub


Private Sub mnuMassRoomEditor_Click()
Unload frmMassRoomEditor
Load frmMassRoomEditor
End Sub

Private Sub mnuMinimizeWindows_Click()
Dim frmForm As Form
On Error Resume Next

For Each frmForm In Forms
    If Not frmForm Is Me Then frmForm.WindowState = vbMinimized
Next

Set frmForm = Nothing
End Sub

Private Sub mnuMMUDExplorer_Click()

Unload frmMME_Export
frmMME_Export.Show

End Sub

Private Sub mnuMonsterAttackSim_Click()
Unload frmMonsterAttackSim
Load frmMonsterAttackSim
End Sub

Private Sub mnuMonsterIndexCopyLine_Click()
Call frmMonsterIndex.CopyLine
End Sub

Private Sub mnuMonsterListCopyLine_Click()
Call frmMonsterNPC_List.CopyLine
End Sub

Private Sub mnuMonsterNPCList_Click()
Unload frmMonsterNPC_List
Load frmMonsterNPC_List
End Sub

Private Sub mnuMonstersCombineExp_Click()
Dim nStatus As Integer, sFile As String, bUpdateRec As Boolean
Dim nYesNo As Integer, nExp As Currency, nExpMulti As Currency
On Error GoTo error:

nYesNo = MsgBox("This will combine the Experience and Expierence Multiplier fields " _
    & "into one field (intended to fix the experience bug when monsters die from DOT spells)." _
    & "If a monster has an exp multi value <= 0 it will be set to 1." & vbCrLf & vbCrLf _
    & "Continue?", vbYesNo + vbQuestion + vbDefaultButton2)
If nYesNo <> vbYes Then Exit Sub

If eDatFileVersion < v111j Then MsgBox "This is only available for MajorMUD v1.11j and greater.", vbInformation: Exit Sub
If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

UnloadForms (Me.Name)
frmProgressBar.sCaption = "Combining Monster Exp"
frmProgressBar.lblCaption.Caption = "Working ..."
frmProgressBar.cmdCancel.Enabled = True
frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_KNMSR
frmProgressBar.Show
frmMain.Enabled = False
DoEvents

nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
    GoTo out:
Else
    nStatus = BTRCALL(BSTAT, MonsterPosBlock, DBStatDatabuf, Len(Monsterdatabuf), 0, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        Call frmProgressBar.SetRange(30000)
    Else
        DBStatRowToStruct DBStatDatabuf.buf
        Call frmProgressBar.SetRange(DBStat.nRecords)
    End If
    nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
End If

bStopProcess = False
Do While nStatus = 0 And bStopProcess = False
    bUpdateRec = False
    MonsterRowToStruct Monsterdatabuf.buf
    frmProgressBar.lblPanel(1).Caption = Monsterrec.Number
    Call frmProgressBar.IncreaseProgress
    
    nExp = SLong2ULong(Monsterrec.Experience)
    nExpMulti = SLong2ULong(Monsterrec.ExpMulti)
    
    If nExpMulti = 1 Then GoTo Skip:
    If nExpMulti <= 0 Then nExpMulti = 1
    
    Monsterrec.Experience = ULong2SLong(nExp * nExpMulti)
    Monsterrec.ExpMulti = 1
    
    nStatus = UpdateMonster
    If Not nStatus = 0 Then
        MsgBox "Update record Error, " & BtrieveErrorCode(nStatus)
        GoTo out:
    End If
Skip:
    nStatus = BTRCALL(BGETNEXT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
    DoEvents
Loop

If bStopProcess = True Then GoTo out:

If Not nStatus = 9 And Not nStatus = 0 Then
    MsgBox "Abnormal Exit: " & BtrieveErrorCode(nStatus), vbOKOnly + vbExclamation
Else
    frmProgressBar.ProgressBar.Value = frmProgressBar.ProgressBar.Max
    DoEvents
    MsgBox "Complete!", vbInformation
End If

out:
On Error Resume Next
frmMain.Enabled = True
Unload frmProgressBar
If Me.Visible Then Me.SetFocus
Exit Sub
error:
Call HandleError("mnuMonstersCombineExp_Click")
Resume out:
End Sub

Private Sub mnuMultiplyBossExp_Click()
Unload frmUniversalModifier
Load frmUniversalModifier
frmUniversalModifier.cmbEditor.ListIndex = 2 'monster

If eDatFileVersion >= v111j Then
    frmUniversalModifier.cmbField.ListIndex = 1 'exp multiplier
Else
    frmUniversalModifier.cmbField.ListIndex = 0 'exp
End If

frmUniversalModifier.cmbModifier.ListIndex = 2 '*
frmUniversalModifier.txtValue.Text = 2 '2 to multiply by 2

frmUniversalModifier.chkOnlyIfOn(0).Value = 1 'turn on "only if"
frmUniversalModifier.cmbOnlyIf(0).ListIndex = 0 'game limit
frmUniversalModifier.cmbOnlyIfModifier(0).ListIndex = 0 '=
frmUniversalModifier.txtOnlyIfValue(0).Text = 1 '1 for 1 game limit

MsgBox "The Universal Modifier Fields have been set up for a Boss Exp Multiply." _
    & vbCrLf & "You can modify the settings if need be.  The '2' in the value field is what the exp will be multiplied by." _
    & vbCrLf & vbCrLf & "NOTE: The key here is the 'Only if Game Limit = 1' ... that signifies that the monster is _probably_ a boss.", vbInformation

End Sub

Private Sub mnuQuestOrganizer_Click()
Unload frmQuests
Load frmQuests

End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuExp_Click()
On Error GoTo error:
Dim nStatus As Integer, nYesNo As Long, nMultiplier As Integer, nCap As Double, nMaxCap As Double
Dim nBase As Double, nMulti As Double, nMultiMax As Double, x As Double, nTemp As Double

nMultiplier = Val(InputBox("Multiply monster exp by how many times (2-20)?", "Monster EXP nMultiplier", "2"))
If nMultiplier <= 0 Then Exit Sub
If nMultiplier > 20 Then nMultiplier = 20
If nMultiplier < 2 Then nMultiplier = 2

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

If eDatFileVersion >= v111j Then
    nMaxCap = 2147483646
Else
    nMaxCap = 4294967295#
End If

nCap = Val(InputBox("Do you want to place a cap on the monster exp?" & vbCrLf & "(enter 0 for a max of " & nMaxCap & ")", "Monster EXP nMultiplier", "0"))
If nCap <= 0 Then nCap = nMaxCap
If nCap > nMaxCap Then nCap = nMaxCap

nYesNo = MsgBox("Are you sure you want to increase monster EXP by " & nMultiplier & "x, and cap the experience at " & nCap & "?", vbYesNo + vbQuestion + vbDefaultButton2)
If nYesNo <> vbYes Then Exit Sub
    
UnloadForms (Me.Name)
frmProgressBar.sCaption = "Multiplying Monster EXP"
frmProgressBar.lblCaption.Caption = "Multiplying Monster EXP ..."
frmProgressBar.cmdCancel.Enabled = True
frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_KNMSR
frmProgressBar.Show
frmMain.Enabled = False
DoEvents

nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
    GoTo finish:
Else
    nStatus = BTRCALL(BSTAT, MonsterPosBlock, DBStatDatabuf, Len(Monsterdatabuf), 0, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        Call frmProgressBar.SetRange(30000)
    Else
        DBStatRowToStruct DBStatDatabuf.buf
        Call frmProgressBar.SetRange(DBStat.nRecords)
    End If
    nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
End If

bStopProcess = False
Do While nStatus = 0 And bStopProcess = False
    MonsterRowToStruct Monsterdatabuf.buf
    frmProgressBar.lblPanel(1).Caption = Monsterrec.Number
    Call frmProgressBar.IncreaseProgress
    
    If eDatFileVersion >= v111j Then
    
        nBase = SLong2ULong(Monsterrec.Experience) * SLong2ULong(Monsterrec.ExpMulti)
        nBase = nBase * nMultiplier
        If nBase > nCap Then nBase = nCap
        
tryagain:
        If nBase > 100000 Then
            nMultiMax = 20
            For x = 20 To 32767
                If x * 65538 >= nBase Then
                    nMultiMax = x
                    Exit For
                End If
            Next x
            
            nMulti = 1
            For x = 3 To nMultiMax
                nTemp = nBase Mod x
                If nTemp = 0 Then nMulti = x
            Next x
            
            If nMulti = 1 Then
                nBase = nBase - 1
                GoTo tryagain:
            End If
            
            nBase = nBase / nMulti
        Else
            nMulti = 1
        End If
        
        If nBase <= 0 Then nBase = 1
        
        Monsterrec.Experience = ULong2SLong(nBase)
        Monsterrec.ExpMulti = ULong2SLong(nMulti)
    
'        nBase = SLong2ULong(Monsterrec.Experience)
'        nMulti = SLong2ULong(Monsterrec.ExpMulti)
'
'        nBase = nBase * nMultiplier
'
'        If nBase * nMulti > 2147483646 Then
'            If nCap = 2147483646 Then
'                Monsterrec.Experience = 65538
'                Monsterrec.ExpMulti = 32767
'            Else
'                Monsterrec.Experience = ULong2SLong(nCap)
'                Monsterrec.ExpMulti = 1
'            End If
'        Else
'            If nBase * nMulti < nCap Then
'                Monsterrec.Experience = ULong2SLong(nBase)
'            Else
'                Monsterrec.Experience = ULong2SLong(nCap)
'                Monsterrec.ExpMulti = 1
'            End If
'        End If

    Else
        nBase = SLong2ULong(Monsterrec.Experience)
        nBase = nBase * nMultiplier
        If nBase > nCap Then nBase = nCap
        Monsterrec.Experience = ULong2SLong(nBase)
    End If
    
    nStatus = UpdateMonster
    If Not nStatus = 0 Then
        MsgBox "Update monster Error: " & BtrieveErrorCode(nStatus)
        GoTo finish:
    End If
    nStatus = BTRCALL(BGETNEXT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
    DoEvents
Loop

If bStopProcess = True Then GoTo finish:

If Not nStatus = 9 And Not nStatus = 0 Then
    MsgBox "Abnormal Exit: " & BtrieveErrorCode(nStatus), vbOKOnly + vbExclamation
Else
    frmProgressBar.ProgressBar.Value = frmProgressBar.ProgressBar.Max
    DoEvents
    MsgBox "Complete!", vbInformation
End If

finish:
Unload frmProgressBar
frmMain.Enabled = True
If Me.Visible Then Me.SetFocus
Exit Sub
error:
Call HandleError
End Sub

Private Sub mnuNoLevelRestrictions_Click()
On Error GoTo error:
Dim nStatus As Integer, nYesNo As Long, i As Integer

nYesNo = MsgBox("Are you sure you remove the level restrictions from all items?", vbYesNo + vbQuestion + vbDefaultButton2)
If nYesNo <> 6 Then Exit Sub
    
If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

UnloadForms (Me.Name)
frmProgressBar.sCaption = "Removing Level Restrictions on Items"
frmProgressBar.lblCaption.Caption = "Removing Level Restrictions on Items..."
frmProgressBar.cmdCancel.Enabled = True
frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_ITEMS
frmProgressBar.Show
frmMain.Enabled = False
DoEvents

nStatus = BTRCALL(BGETFIRST, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
    GoTo finish:
Else
    nStatus = BTRCALL(BSTAT, ItemPosBlock, DBStatDatabuf, Len(Itemdatabuf), 0, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        Call frmProgressBar.SetRange(30000)
    Else
        DBStatRowToStruct DBStatDatabuf.buf
        Call frmProgressBar.SetRange(DBStat.nRecords)
    End If
    nStatus = BTRCALL(BGETFIRST, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
End If

bStopProcess = False
Do While nStatus = 0 And bStopProcess = False
    ItemRowToStruct Itemdatabuf.buf
    frmProgressBar.lblPanel(1).Caption = Itemrec.Number
    Call frmProgressBar.IncreaseProgress
    
    For i = 0 To 19
        If Itemrec.AbilityA(i) = 135 Or Itemrec.AbilityA(i) = 136 Then
            Itemrec.AbilityA(i) = 0
            Itemrec.AbilityB(i) = 0
        End If
    Next i
    nStatus = UpdateItem
    If Not nStatus = 0 Then
        MsgBox "mnuNoLimited BUPDATE, Error: " & BtrieveErrorCode(nStatus)
        GoTo finish:
    End If
    nStatus = BTRCALL(BGETNEXT, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
    DoEvents
Loop

If bStopProcess = True Then GoTo finish:

If Not nStatus = 9 And Not nStatus = 0 Then
    MsgBox "Abnormal Exit: " & BtrieveErrorCode(nStatus), vbOKOnly + vbExclamation
Else
    frmProgressBar.ProgressBar.Value = frmProgressBar.ProgressBar.Max
    DoEvents
    MsgBox "Item level restrictions removed.", vbInformation
End If

finish:
Unload frmProgressBar
frmMain.Enabled = True
If Me.Visible Then Me.SetFocus
Exit Sub
error:
Call HandleError
End Sub


Private Sub mnuNoLimited_Click()
On Error GoTo error:
Dim nStatus As Integer, nYesNo As Long, x As Integer

nYesNo = MsgBox("Are you sure you want to remove the game limits from all items (except those listed in the general help)?", vbYesNo + vbQuestion + vbDefaultButton2)
If nYesNo <> 6 Then Exit Sub
    
If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

UnloadForms (Me.Name)
frmProgressBar.sCaption = "Removing Limits on Items"
frmProgressBar.lblCaption.Caption = "Removing Limits on Items..."
frmProgressBar.cmdCancel.Enabled = True
frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_ITEMS
frmProgressBar.Show
frmMain.Enabled = False
DoEvents
    
nStatus = BTRCALL(BGETFIRST, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
    GoTo finish:
Else
    nStatus = BTRCALL(BSTAT, ItemPosBlock, DBStatDatabuf, Len(Itemdatabuf), 0, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        Call frmProgressBar.SetRange(30000)
    Else
        DBStatRowToStruct DBStatDatabuf.buf
        Call frmProgressBar.SetRange(DBStat.nRecords)
    End If
    nStatus = BTRCALL(BGETFIRST, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
End If

bStopProcess = False
Do While nStatus = 0 And bStopProcess = False
    ItemRowToStruct Itemdatabuf.buf
    frmProgressBar.lblPanel(1).Caption = Itemrec.Number
    Call frmProgressBar.IncreaseProgress
    
    Select Case Itemrec.Number
        Case 1078: GoTo skipitem:     'flaming portal
        Case 1326: GoTo skipitem:     'shimmering portal
        Case 1531: GoTo skipitem:     'shadowy portal
        Case 1643: GoTo skipitem:     'chaotic vortex
        Case 1686: GoTo skipitem:     'hideous face
        Case 1749: GoTo skipitem:     'yellow bone portal
        Case 839: GoTo skipitem:      'nightblack portal
    End Select
    
    For x = 0 To 19 'gang abilities
        If Itemrec.AbilityA(x) >= 181 And Itemrec.AbilityA(x) <= 184 Then GoTo skipitem:
    Next
    
    Itemrec.GameLimit = 0
        
    nStatus = UpdateItem
    If Not nStatus = 0 Then
        MsgBox "mnuNoLimited BUPDATE, Error: " & BtrieveErrorCode(nStatus)
        GoTo finish:
    End If
skipitem:
    nStatus = BTRCALL(BGETNEXT, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
    DoEvents
Loop

If bStopProcess = True Then GoTo finish:

If Not nStatus = 9 And Not nStatus = 0 Then
    MsgBox "Abnormal Exit: " & BtrieveErrorCode(nStatus), vbOKOnly + vbExclamation
Else
    frmProgressBar.ProgressBar.Value = frmProgressBar.ProgressBar.Max
    DoEvents
    MsgBox "Only the items listed in the general help are now limited.", vbInformation
End If

finish:
Unload frmProgressBar
frmMain.Enabled = True
If Me.Visible Then Me.SetFocus
Exit Sub
error:
Call HandleError
End Sub

Private Sub mnuRecordChanger_Click()
'MsgBox "Sorry, this isn't ready yet."
'Exit Sub
frmRecordChange.Show
End Sub

Private Sub mnuResetMonsterKills_Click()
On Error GoTo error:
Dim nStatus As Integer, nYesNo As Long

nYesNo = MsgBox("Are you sure you want to reset the monster's Last killed times?", vbYesNo + vbQuestion + vbDefaultButton2)
If nYesNo <> 6 Then Exit Sub
    
If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
    
UnloadForms (Me.Name)
frmProgressBar.sCaption = "Reseting Monster Last Killed Date/Time"
frmProgressBar.lblCaption.Caption = "Reseting Monster Last Killed Date/Time ..."
frmProgressBar.cmdCancel.Enabled = True
frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_KNMSR
frmProgressBar.Show
frmMain.Enabled = False
DoEvents

nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
    GoTo finish:
Else
    nStatus = BTRCALL(BSTAT, MonsterPosBlock, DBStatDatabuf, Len(Monsterdatabuf), 0, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        Call frmProgressBar.SetRange(30000)
    Else
        DBStatRowToStruct DBStatDatabuf.buf
        Call frmProgressBar.SetRange(DBStat.nRecords)
    End If
    nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
End If

bStopProcess = False
Do While nStatus = 0 And bStopProcess = False
    MonsterRowToStruct Monsterdatabuf.buf
    frmProgressBar.lblPanel(1).Caption = Monsterrec.Number
    Call frmProgressBar.IncreaseProgress
    
    Monsterrec.TimeKilled = 0
    Monsterrec.DateKilled = 0
    nStatus = UpdateMonster
        If Not nStatus = 0 Then
            MsgBox "Update record Error, " & BtrieveErrorCode(nStatus)
            GoTo finish:
        End If
    nStatus = BTRCALL(BGETNEXT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
    DoEvents
Loop

If bStopProcess = True Then GoTo finish:

If Not nStatus = 9 And Not nStatus = 0 Then
    MsgBox "Abnormal Exit: " & BtrieveErrorCode(nStatus), vbOKOnly + vbExclamation
Else
    frmProgressBar.ProgressBar.Value = frmProgressBar.ProgressBar.Max
    DoEvents
    MsgBox "Complete!", vbInformation
End If

finish:
frmMain.Enabled = True
Unload frmProgressBar
If Me.Visible Then Me.SetFocus
Exit Sub
error:
Call HandleError
End Sub

Private Sub mnuResetMonsterKillsToTime_Click()
On Error GoTo error:
Dim nStatus As Integer, sDate As String, nTime As Long, nDate As Long

sDate = InputBox("This will set all the monster kill times to the date specified." _
    & "  Enter the date as MM/DD/YYYY including the preceding zeros." _
    & "  The time will be set to 01:01:02.", "Set Monster Kill Time", "01/01/1990")
If sDate = "" Then Exit Sub

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

nDate = Date2DOSDate(sDate)
If nDate = -1 Then Exit Sub

nTime = Time2DOSTime("01:01:02")

UnloadForms (Me.Name)
frmProgressBar.sCaption = "Setting Monster Last Killed Date/Time"
frmProgressBar.lblCaption.Caption = "Setting Monster Last Killed Date/Time ..."
frmProgressBar.cmdCancel.Enabled = True
frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_KNMSR
frmProgressBar.Show
frmMain.Enabled = False
DoEvents

nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
    GoTo finish:
Else
    nStatus = BTRCALL(BSTAT, MonsterPosBlock, DBStatDatabuf, Len(Monsterdatabuf), 0, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        Call frmProgressBar.SetRange(30000)
    Else
        DBStatRowToStruct DBStatDatabuf.buf
        Call frmProgressBar.SetRange(DBStat.nRecords)
    End If
    nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
End If

bStopProcess = False
Do While nStatus = 0 And bStopProcess = False
    MonsterRowToStruct Monsterdatabuf.buf
    frmProgressBar.lblPanel(1).Caption = Monsterrec.Number
    Call frmProgressBar.IncreaseProgress
    
    Monsterrec.TimeKilled = nTime
    Monsterrec.DateKilled = nDate
    nStatus = UpdateMonster
        If Not nStatus = 0 Then
            MsgBox "Update record Error, " & BtrieveErrorCode(nStatus)
            GoTo finish:
        End If
    nStatus = BTRCALL(BGETNEXT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
    DoEvents
Loop

If bStopProcess = True Then GoTo finish:

If Not nStatus = 9 And Not nStatus = 0 Then
    MsgBox "Abnormal Exit: " & BtrieveErrorCode(nStatus), vbOKOnly + vbExclamation
Else
    frmProgressBar.ProgressBar.Value = frmProgressBar.ProgressBar.Max
    DoEvents
    MsgBox "Complete!", vbInformation
End If

finish:
frmMain.Enabled = True
Unload frmProgressBar
If Me.Visible Then Me.SetFocus
Exit Sub
error:
Call HandleError
End Sub

Private Sub mnuRestoreWindows_Click()
Dim frmForm As Form
On Error Resume Next

For Each frmForm In Forms
    If Not frmForm Is Me Then
        frmForm.WindowState = vbNormal
        frmForm.Show
    End If
Next

Set frmForm = Nothing
End Sub

Private Sub mnuRetrainUsers_Click()
On Error GoTo error:
Dim nStatus As Integer, nYesNo As Long, nRecord As Long

nRecord = 1
nYesNo = MsgBox("Are you sure you want to give all users a retrain?", vbYesNo + vbQuestion + vbDefaultButton2)
If nYesNo <> 6 Then Exit Sub

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

UnloadForms (Me.Name)
frmProgressBar.sCaption = "Retrain Users"
frmProgressBar.lblCaption.Caption = "Retraining Users ..."
frmProgressBar.cmdCancel.Enabled = True
frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_USERS
frmProgressBar.ProgressBar.Value = 0
frmProgressBar.Show
frmMain.Enabled = False
DoEvents

nStatus = BTRCALL(BGETFIRST, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
    GoTo finish:
Else
    nStatus = BTRCALL(BSTAT, UserPosBlock, DBStatDatabuf, Len(Userdatabuf), 0, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        Call frmProgressBar.SetRange(100)
    Else
        DBStatRowToStruct DBStatDatabuf.buf
        Call frmProgressBar.SetRange(DBStat.nRecords)
    End If
    nStatus = BTRCALL(BGETFIRST, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0)
End If

bStopProcess = False
Do While nStatus = 0 And bStopProcess = False
    UserRowToStruct Userdatabuf.buf
    frmProgressBar.lblPanel(1).Caption = nRecord
    Call frmProgressBar.IncreaseProgress
    
    Userrec.CPRemaining = -1
    
    nStatus = UpdateUser
    If Not nStatus = 0 Then
        MsgBox "Update record Error, " & BtrieveErrorCode(nStatus)
        GoTo finish:
    End If
    nStatus = BTRCALL(BGETNEXT, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0)
    nRecord = nRecord + 1
    DoEvents
Loop

If bStopProcess = True Then GoTo finish:

If Not nStatus = 9 And Not nStatus = 0 Then
    MsgBox "Abnormal Exit: " & BtrieveErrorCode(nStatus), vbOKOnly + vbExclamation
Else
    frmProgressBar.ProgressBar.Value = frmProgressBar.ProgressBar.Max
    DoEvents
    MsgBox "Complete!", vbInformation
End If

finish:
frmMain.Enabled = True
Unload frmProgressBar
If Me.Visible Then Me.SetFocus
Exit Sub
error:
Call HandleError
End Sub

Private Sub mnuRoomPad_Click()
On Error GoTo error:
Dim nMap As Long, nRoom As Long, nStatus As Integer, x As Integer, sFile As String
Dim fso As FileSystemObject, ts As TextStream, nPrevRoom As Long, nPrevMap As Long
Dim nLastRoom As Long, nYesNo As Integer, bAll As Boolean, nMaxRooms As Long

nMap = Val(InputBox("This will make sure that every room number is used up until the highest number created.  " _
        & "During bbs load, rooms stop loading shops (and possibly some other things) to " _
        & "memory for regen when 50 room numbers have been skipped.  This will create those " _
        & "rooms by inserting 'Buffer' rooms.  This also fixes offline room item sweeps (new " _
        & "rooms are normally skipped during a sweep because of this same reason)." _
        & vbCrLf & "Enter a value of -1 to pad rooms on maps 1-18." _
        & vbCrLf & vbCrLf & "Enter map number to pad rooms on:", "Pad Rooms?", 1))

If nMap = 0 Then Exit Sub
If nMap = -1 Then bAll = True
If nMap < -1 Then Exit Sub

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then MsgBox "Error getting first room, error: " & BtrieveErrorCode(nStatus): Exit Sub

Set fso = CreateObject("Scripting.FileSystemObject")

If Right(App.Path, 1) = "\" Then
    sFile = App.Path & "NMR-Log_RoomPad.txt"
Else
    sFile = App.Path & "\NMR-Log_RoomPad.txt"
End If

If fso.FileExists(sFile) Then Call fso.DeleteFile(sFile, True)
Set ts = fso.OpenTextFile(sFile, ForWriting, True)

ts.WriteLine ("Room pad job started " & Date & " @ " & Time)
ts.WriteBlankLines (1)

frmProgressBar.sCaption = "Padding Room Numbers"
frmProgressBar.lblCaption = ""
frmProgressBar.cmdCancel.Enabled = True
frmProgressBar.ProgressBar.Value = 0
frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_MP
frmProgressBar.lblPanel(1).Caption = ""
frmProgressBar.Show
frmMain.Enabled = False
DoEvents

bStopProcess = False

nStatus = BTRCALL(BSTAT, RoomPosBlock, DBStatDatabuf, Len(Roomdatabuf), 0, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    nMaxRooms = 30000
Else
    DBStatRowToStruct DBStatDatabuf.buf
    nMaxRooms = DBStat.nRecords
End If

For nMap = IIf(bAll, 1, nMap) To IIf(bAll, 18, nMap)
    If bStopProcess Then Exit For
    'determine last room number of specified map
    nLastRoom = -1
    frmProgressBar.lblCaption = "Padding Room Numbers on Map " & nMap & "..."
    frmProgressBar.lblPanel(1).Caption = "determining last room on map " & nMap & "..."
    nStatus = BTRCALL(BGETLAST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    RoomRowToStruct Roomdatabuf.buf
    Call frmProgressBar.SetRange(nMaxRooms)
    DoEvents
    
    Do While nStatus = 0 And bStopProcess = False
        RoomRowToStruct Roomdatabuf.buf
        Call frmProgressBar.IncreaseProgress
        If Roomrec.MapNumber = nMap Then
            nLastRoom = Roomrec.RoomNumber
            Exit Do
        End If
        nStatus = BTRCALL(BGETPREVIOUS, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
        If Not bUseCPU Then DoEvents
    Loop
    
    If bStopProcess Then Exit For
    
    If nLastRoom = -1 Then
        MsgBox "Unable to determine last room on map " & nMap, vbOKOnly + vbExclamation
        ts.WriteLine "Unable to determine last room on map " & nMap
        GoTo next_map:
    Else
        ts.WriteLine "Last room of map " & nMap & " determined as room #" & nLastRoom
        ts.WriteBlankLines (1)
    End If
    
    'do operation
    Call frmProgressBar.SetRange(nLastRoom)
    
    nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    nRoom = 0
    bStopProcess = False
    
    Do While nStatus = 0 And bStopProcess = False
        RoomRowToStruct Roomdatabuf.buf
    
        If nRoom + 1 = nLastRoom Then Exit Do
        
        If Roomrec.RoomNumber > nRoom + 1 Then
            
            'make the room all null
            For x = 1 To Len(Roomdatabuf)
                Roomdatabuf.buf(x) = &H0
            Next
            
            RoomRowToStruct Roomdatabuf.buf
            
            'set the necessary properties
            Roomrec.MapNumber = nMap
            Roomrec.RoomNumber = nRoom + 1
            Roomrec.Name = "Buffer Room " & (nRoom + 1) & Chr(0)
            Roomrec.AnsiMap = "W" & strDatCallLetters & "MAP01.ANS" & Chr(0)
            Roomrec.Desc(0) = "This is just a buffer room to keep the rooms scanning on load." & Chr(0)
            Roomrec.RoomExit(0) = 1
            Roomrec.RoomType(0) = 8
            Roomrec.Para1(0) = 1
            
            'cunstruct row
            RoomStructToRow Roomdatabuf.buf
            
            'insert new room
            nStatus = BTRCALL(BINSERT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
            If Not nStatus = 0 Then
                MsgBox "Error inserting Map " & nMap & " Room " & Roomrec.RoomNumber & ", Error: " & BtrieveErrorCode(nStatus)
                GoTo kill:
            End If
            
            ts.WriteLine "Inserted Room: Map " & nMap & " Room " & Roomrec.RoomNumber
            nRoom = Roomrec.RoomNumber
            
            'reretrieve last room that should be before the room we just inserted
            RoomKeyStruct.MapNum = nPrevMap
            RoomKeyStruct.RoomNum = nPrevRoom
            nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
            If Not nStatus = 0 Then
                MsgBox "Error re-retrieving Map " & nPrevMap & " Room " & nPrevRoom & ", Error: " & BtrieveErrorCode(nStatus)
                GoTo kill:
            End If
        Else
            If Roomrec.MapNumber = nMap Then
                nRoom = Roomrec.RoomNumber
                Call frmProgressBar.IncreaseProgress
                frmProgressBar.lblPanel(1).Caption = Roomrec.RoomNumber
            End If
            
            nPrevRoom = Roomrec.RoomNumber
            nPrevMap = Roomrec.MapNumber
        End If
        
        nStatus = BTRCALL(BGETNEXT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
        If Not bUseCPU Then DoEvents
    Loop
next_map:
    ts.WriteBlankLines (1)
    ts.WriteBlankLines (1)
Next nMap

If bStopProcess = True Then
    ts.WriteLine "...canceled by user"
    GoTo kill:
End If

If Not nStatus = 9 And Not nStatus = 0 Then
    MsgBox "Abnormal Exit: " & BtrieveErrorCode(nStatus), vbOKOnly + vbExclamation
    ts.WriteLine "Exited because of btrieve error: " & BtrieveErrorCode(nStatus)
Else
    frmProgressBar.ProgressBar.Value = frmProgressBar.ProgressBar.Max
    DoEvents
    ts.WriteBlankLines (1)
    ts.WriteLine ("Complete: " & Date & " @ " & Time)
    ts.Close
    
    nYesNo = MsgBox("Complete, view log?", vbYesNo + vbQuestion, "View?")
    If nYesNo = vbYes Then Call ShellExecute(0&, "open", sFile, vbNullString, vbNullString, vbNormalFocus)
    DoEvents
End If

kill:
On Error Resume Next
ts.Close
Set ts = Nothing
Set fso = Nothing
frmMain.Enabled = True
Unload frmProgressBar
If Me.Visible Then Me.SetFocus
Exit Sub
error:
Call HandleError
Resume kill:
End Sub

Private Sub mnuRoomsCombineItems_Click()
Dim nStatus As Integer, bUpdateRec As Boolean
Dim nYesNo As Integer, x As Integer, y As Integer, nItem As Long
On Error GoTo error:

nYesNo = MsgBox("This will go through all the rooms and combine items with the same number " _
    & "of uses into one slot.  Continue?", vbYesNo + vbQuestion + vbDefaultButton2)
If nYesNo <> vbYes Then Exit Sub

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

UnloadForms (Me.Name)
frmProgressBar.sCaption = "Combining like items in rooms"
frmProgressBar.lblCaption.Caption = "Working ..."
frmProgressBar.cmdCancel.Enabled = True
frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_KNMSR
frmProgressBar.Show
frmMain.Enabled = False
DoEvents

nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
    GoTo out:
End If

nStatus = BTRCALL(BSTAT, RoomPosBlock, DBStatDatabuf, Len(Roomdatabuf), 0, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    Call frmProgressBar.SetRange(30000)
Else
    DBStatRowToStruct DBStatDatabuf.buf
    Call frmProgressBar.SetRange(DBStat.nRecords)
End If

nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
    GoTo out:
End If

bStopProcess = False
Do While nStatus = 0 And bStopProcess = False
    bUpdateRec = False
    RoomRowToStruct Roomdatabuf.buf
    frmProgressBar.lblPanel(1).Caption = Roomrec.MapNumber & "/" & Roomrec.RoomNumber
    Call frmProgressBar.IncreaseProgress
    
recheck:
    For x = 0 To 16
        If Roomrec.RoomItems(x) > 0 Then
            For y = 0 To 16
                If (y <> x) And (Roomrec.RoomItems(y) = Roomrec.RoomItems(x)) _
                    And (Roomrec.RoomItemUses(x) = Roomrec.RoomItemUses(y)) Then
                    
                    bUpdateRec = True
                    Roomrec.RoomItemQty(x) = Roomrec.RoomItemQty(x) + Roomrec.RoomItemQty(y) + 1
                    Roomrec.RoomItems(y) = 0
                    Roomrec.RoomItemUses(y) = 0
                    Roomrec.RoomItemQty(y) = 0
                    GoTo recheck:
                End If
            Next y
        End If
    Next x
    
recheck2:
    For x = 0 To 14
        If Roomrec.InvisItems(x) > 0 Then
            For y = 0 To 14
                If (y <> x) And (Roomrec.InvisItems(y) = Roomrec.InvisItems(x)) _
                    And (Roomrec.InvisItemUses(x) = Roomrec.InvisItemUses(y)) Then
                    
                    bUpdateRec = True
                    Roomrec.InvisItemQty(x) = Roomrec.InvisItemQty(x) + Roomrec.InvisItemQty(y) + 1
                    Roomrec.InvisItems(y) = 0
                    Roomrec.InvisItemUses(y) = 0
                    Roomrec.InvisItemQty(y) = 0
                    GoTo recheck2:
                End If
            Next y
        End If
    Next x
    
    If bUpdateRec Then
        nStatus = UpdateRoom
        If Not nStatus = 0 Then
            MsgBox "Update record Error, " & BtrieveErrorCode(nStatus)
            GoTo out:
        End If
    End If
    nStatus = BTRCALL(BGETNEXT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    DoEvents
Loop

If bStopProcess = True Then GoTo out:

If Not nStatus = 9 And Not nStatus = 0 Then
    MsgBox "Abnormal Exit: " & BtrieveErrorCode(nStatus), vbOKOnly + vbExclamation
Else
    frmProgressBar.ProgressBar.Value = frmProgressBar.ProgressBar.Max
    DoEvents
    MsgBox "Complete!", vbInformation
End If

out:
On Error Resume Next
frmMain.Enabled = True
Unload frmProgressBar
If Me.Visible Then Me.SetFocus
Exit Sub
error:
Call HandleError("mnuRoomsCombineItems_Click")
Resume out:
End Sub

Private Sub mnuSettings_Click()
Load frmSettings
End Sub

Private Sub mnuAbout2_Click()
Unload frmAbout
Load frmAbout
End Sub

Private Sub mnuCompileUpdate_Click()
If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
Call CompileUpdatefile
End Sub

Public Sub mnuHelpTextblocks_Click()
Unload frmHelpTextblocks
Load frmHelpTextblocks
End Sub
Private Sub mnuHelpGeneral_Click()
Unload frmHelpGeneral
Load frmHelpGeneral
End Sub

Private Sub mnuHelpMessages_Click()
Unload frmHelpMessages
Load frmHelpMessages
End Sub

Private Sub mnuHelpMonsters_Click()
Unload frmHelpMonsters
Load frmHelpMonsters
End Sub

Public Sub mnuHelpRooms_Click()
Unload frmHelpRooms
Load frmHelpRooms
End Sub

Private Sub mnuShopRestock_Click()
On Error GoTo error:
Dim nStatus As Integer, x As Integer ', bMod As Boolean

x = MsgBox("This will set the current stock of the items in all the shops to whatever " & vbCrLf _
            & "the max stock is for each item.  If an item has a zero regen time, regen %, " & vbCrLf _
            & "OR regen #, it will be skipped.  Gangshops will be skipped completely." _
            & vbCrLf & vbCrLf & "Note: Doing this on live dats is pointless as the buffers will overwrite it when saved." _
            & vbCrLf & vbCrLf _
            & "Continue?", vbYesNo + vbQuestion + vbDefaultButton2, "Restock Shops?")
If Not x = vbYes Then Exit Sub

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

Unload frmShop
frmProgressBar.sCaption = "Restocking Shops"
frmProgressBar.lblCaption = "Restocking Shops ..."
frmProgressBar.cmdCancel.Enabled = True
frmProgressBar.ProgressBar.Value = 0
frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_SHOPS
frmProgressBar.lblPanel(1).Caption = ""
frmProgressBar.Show
frmMain.Enabled = False
DoEvents

nStatus = BTRCALL(BGETFIRST, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
    GoTo out:
Else
    nStatus = BTRCALL(BSTAT, ShopPosBlock, DBStatDatabuf, Len(Shopdatabuf), 0, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        Call frmProgressBar.SetRange(30000)
    Else
        DBStatRowToStruct DBStatDatabuf.buf
        Call frmProgressBar.SetRange(DBStat.nRecords)
    End If
    nStatus = BTRCALL(BGETFIRST, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
End If

bStopProcess = False
Do While nStatus = 0 And bStopProcess = False
    ShopRowToStruct Shopdatabuf.buf
    frmProgressBar.lblPanel(1).Caption = Shoprec.Number
    frmProgressBar.IncreaseProgress
    
    If Not Shoprec.ShopType = 11 Then 'gang shop
        For x = 0 To 19
            'bMod = False
            If Shoprec.ShopItemNumber(x) > 0 Then
                If Shoprec.ShopRgnNumber(x) > 0 Then
                    If Shoprec.ShopRgnPercentage(x) > 0 Then
                        If Shoprec.ShopRgnTime(x) > 0 Then
                            'bMod = True
                            Shoprec.ShopNow(x) = Shoprec.ShopMax(x)
                        End If
                    End If
                End If
            End If
            'If bMod = False Then Shoprec.ShopNow(x) = 0
        Next x
    End If
    
    nStatus = UpdateShop
    If Not nStatus = 0 Then Exit Do
    
    nStatus = BTRCALL(BGETNEXT, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

If bStopProcess = True Then GoTo out:

If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Exited at shop #" & Shoprec.Number & " because of btrieve error: " & BtrieveErrorCode(nStatus), vbExclamation
Else
    MsgBox "Restock complete.", vbInformation
End If

out:
frmMain.Enabled = True
Unload frmProgressBar
If Me.Visible Then Me.SetFocus
Exit Sub
error:
Call HandleError
On Error Resume Next
frmMain.Enabled = True
Unload frmProgressBar
End Sub

Private Sub mnuStripChars_Click()
Dim sStrip As String, nResult As Integer

sStrip = InputBox("This will strip the specified string of characters off the end of every textblock (if it's there).  " _
            & "Use this if you see a bunch of strange characters at the end of every textblock." & vbCrLf _
            & vbCrLf _
            & "Paste the text to be stripped:", "Textblock String Stripper")

If sStrip = "" Then Exit Sub

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

nResult = StripTextblocks(sStrip)
If nResult = 0 Then
    MsgBox "Textblock stripping Complete.", vbInformation
Else
    If nResult > 0 Then
        MsgBox "Textblock stripper did not complete -- Btrieve error: " & BtrieveErrorCode(nResult), vbExclamation
    End If
End If

If Me.Visible Then Me.SetFocus
End Sub
Private Function StripTextblocks(sStrip As String) As Integer
On Error GoTo error:
Dim nStatus As Integer
Dim x As Integer, decrypted As String

StripTextblocks = 0

nStatus = BTRCALL(BGETFIRST, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    StripTextblocks = nStatus
    Exit Function
End If

UnloadForms (Me.Name)
DoEvents

nStatus = BTRCALL(BSTAT, TextblockPosBlock, DBStatDatabuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
DBStatRowToStruct DBStatDatabuf.buf
Call frmProgressBar.SetRange(DBStat.nRecords + 1)

nStatus = BTRCALL(BGETFIRST, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    StripTextblocks = nStatus
    Exit Function
Else
    TextblockRowToStruct TextblockDataBuf.buf
End If

frmProgressBar.sCaption = "Stripping TB Chars"
frmProgressBar.lblCaption = "Stripping TB Chars ..."
frmProgressBar.cmdCancel.Enabled = True
frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_TEXT
frmProgressBar.lblPanel(1).Caption = TextblockRec.Number
frmProgressBar.Show
frmMain.Enabled = False
DoEvents

bStopProcess = False
Do While nStatus = 0 And bStopProcess = False
    
    frmProgressBar.lblPanel(1).Caption = TextblockRec.Number
    Call frmProgressBar.IncreaseProgress
    
    TextblockRowToStruct TextblockDataBuf.buf

    decrypted = DecryptTextblock(TextblockRec.Data)
    
    If Not Len(decrypted) < Len(sStrip) Then
        If Right(decrypted, Len(sStrip)) = sStrip Then
            decrypted = Left(decrypted, Len(decrypted) - Len(sStrip))
            
            TextblockRec.Data = EncryptTextblock(decrypted)
            
            nStatus = UpdateTextblock
            If Not nStatus = 0 Then Exit Do
        End If
    End If
    
    nStatus = BTRCALL(BGETNEXT, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

If bStopProcess = True Then
    StripTextblocks = -1
    GoTo kill:
End If

If Not nStatus = 0 And Not nStatus = 9 Then
    StripTextblocks = nStatus
Else
    StripTextblocks = 0
End If

kill:
frmMain.Enabled = True
Unload frmProgressBar
Exit Function

error:
Call HandleError
frmMain.Enabled = True
Unload frmProgressBar
End Function


Private Sub mnuSwingCalculator_Click()
frmSwingCalc.Show
frmSwingCalc.SetFocus
End Sub

Private Sub mnuUnDock_Click()
On Error GoTo error:
Dim frm As Form

If Me.ActiveForm Is Nothing Then Exit Sub
Set frm = Me.ActiveForm
SetParent frm.hwnd, GetDesktopWindow
Set frm = Nothing

Exit Sub

error:
Call HandleError
Set frm = Nothing
End Sub

Private Sub mnuUniversalModifier_Click()
Load frmUniversalModifier
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next

If Not Me.WindowState = vbMinimized Then
    If Me.WindowState = vbMaximized Then
        Call WriteINI("Windows", "MainNoMax", 0)
    Else
        Call WriteINI("Windows", "MainNoMax", 1)
        Call WriteINI("Windows", "MainWidth", Me.Width)
        Call WriteINI("Windows", "MainHeight", Me.Height)
    End If
End If
If Val(ReadINI("Settings", "AutoCompile")) = 1 Then Call CompileUpdatefile
If mnuDisableWrite.Checked = True Then
    Call WriteINI("Settings", "DisableWriting", 1)
Else
    Call WriteINI("Settings", "DisableWriting", 0)
End If

Call UnloadForms("frmMain")
DoEvents

Call StopBtrieve

Erase MGIL()

rsAbilities.Close
dbAbilities.Close
Set rsAbilities = Nothing
Set dbAbilities = Nothing

DoEvents

End Sub

Private Sub mnuUserModifyGang_Click()
Dim nStatus As Integer, strOldName As String, strNewName As String, nRecord As Long
On Error GoTo error:

strOldName = InputBox("This will change the gang name specified on each user from" _
    & " one gang name to another." & vbCrLf & "Enter gangname to change FROM:", "Change Users Gang Names")
If strOldName = "" Then Exit Sub

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

retry:
strNewName = InputBox("Enter gangname to change TO (20 chars max):", "Change Users Gang Names", strNewName)
If strNewName = "" Then Exit Sub
If Len(strNewName) > 20 Then GoTo retry:

UnloadForms (Me.Name)
frmProgressBar.sCaption = "Change User's Gang Names"
frmProgressBar.lblCaption.Caption = "Changing Names ..."
frmProgressBar.cmdCancel.Enabled = True
frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_USERS
frmProgressBar.ProgressBar.Value = 0
frmProgressBar.Show
frmMain.Enabled = False
DoEvents

nStatus = BTRCALL(BGETFIRST, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
    GoTo finish:
Else
    nStatus = BTRCALL(BSTAT, UserPosBlock, DBStatDatabuf, Len(Userdatabuf), 0, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        Call frmProgressBar.SetRange(100)
    Else
        DBStatRowToStruct DBStatDatabuf.buf
        Call frmProgressBar.SetRange(DBStat.nRecords)
    End If
    nStatus = BTRCALL(BGETFIRST, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0)
End If

bStopProcess = False
Do While nStatus = 0 And bStopProcess = False
    UserRowToStruct Userdatabuf.buf
    frmProgressBar.lblPanel(1).Caption = nRecord
    Call frmProgressBar.IncreaseProgress
    
    If LCase(ClipNull(Userrec.GangName)) = LCase(strOldName) Then
        Userrec.GangName = strNewName & String(20 - Len(strNewName), Chr(0))
        
        nStatus = UpdateUser
        If Not nStatus = 0 Then
            MsgBox "Update record Error, " & BtrieveErrorCode(nStatus)
            GoTo finish:
        End If
    End If
    
    nStatus = BTRCALL(BGETNEXT, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0)
    nRecord = nRecord + 1
    DoEvents
Loop

If bStopProcess = True Then GoTo finish:

If Not nStatus = 9 And Not nStatus = 0 Then
    MsgBox "Abnormal Exit: " & BtrieveErrorCode(nStatus), vbOKOnly + vbExclamation
Else
    frmProgressBar.ProgressBar.Value = frmProgressBar.ProgressBar.Max
    DoEvents
    MsgBox "Complete!", vbInformation
End If

finish:
frmMain.Enabled = True
Unload frmProgressBar
If Me.Visible Then Me.SetFocus
Exit Sub
error:
Call HandleError
End Sub

