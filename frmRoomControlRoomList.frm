VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{20D5284F-7B23-4F0A-B8B1-6C9D18B64F1C}#1.0#0"; "exlimiter.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRoomControlRoomList 
   Caption         =   "Control Rooms List"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7395
   Icon            =   "frmRoomControlRoomList.frx":0000
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
      Caption         =   "&Rebuild List"
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
   Begin MSComctlLib.ListView lvControlRoomList 
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
   Begin VB.Label lblLabel 
      Alignment       =   2  'Center
      Caption         =   "(Note: Double-Click to jump to room.  This list may be incomplete/inaccurate until you ""Rebuild List"")"
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
      Left            =   2820
      TabIndex        =   4
      Top             =   60
      Width           =   3615
   End
End
Attribute VB_Name = "frmRoomControlRoomList"
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

If ControlRoomList.Count = 0 Then
    Call cmdBuild_Click
Else
    Call CreateList
End If

End Sub
Private Sub cmdBuild_Click()
On Error GoTo error:

bCancel = False
Me.Enabled = False
frmMain.Enabled = False

Call BuildControlRoomList
'lblLabel.Caption = ""
Call CreateList

out:
On Error Resume Next
frmMain.Enabled = True
Me.Enabled = True
Me.SetFocus

Exit Sub
error:
Call HandleError
Resume out:
End Sub

Private Sub CreateList()
Dim oLI As ListItem, x As Long, y As Integer, z As Integer, sMonsters As String, sArr() As String
Dim nMap As Long, nRoom As Long, sKey As Variant
On Error GoTo error:

If ControlRoomList.Count = 0 Then
    MsgBox "No Control Rooms Found.", vbInformation
    Exit Sub
End If

lvControlRoomList.ListItems.clear
For Each sKey In ControlRoomList.Keys
    sArr = Split(sKey, "/")
    If UBound(sArr()) = 1 Then
        nMap = Val(sArr(0))
        nRoom = Val(sArr(1))
        Set oLI = lvControlRoomList.ListItems.add()
        oLI.Text = sKey
        oLI.ListSubItems.add 1, "Name", GetRoomName(nMap, nRoom)
        oLI.ListSubItems.add 2, "Refs", GetControlRoomListByRoom(nMap, nRoom, 30, True)
    End If
Next

Set oLI = Nothing
If lvControlRoomList.ListItems.Count > 1 Then
    lvControlRoomList.SortOrder = lvwDescending
    Call lvControlRoomList_ColumnClick(lvControlRoomList.ColumnHeaders(1))
End If

out:
Exit Sub
error:
Call HandleError("CreateList")
Resume out:
End Sub

Private Sub AddColumnHeaders()

lvControlRoomList.ColumnHeaders.clear
lvControlRoomList.ColumnHeaders.add 1, "ControlRoom", "Control Room", 1200, lvwColumnLeft
lvControlRoomList.ColumnHeaders.add 2, "Name", "Name", 2800, lvwColumnLeft
lvControlRoomList.ColumnHeaders.add 3, "Refs", "References", 7000, lvwColumnLeft
'lvControlRoomList.ColumnHeaders.add 3, "Name", "Name", 2400, lvwColumnCenter
'lvControlRoomList.ColumnHeaders.add 4, "MapRoom", "Map/Room Monster Regens In", 6000, lvwColumnLeft

End Sub


Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo error:
Dim oLI As ListItem, str As String, x As Integer
Dim fso As FileSystemObject, nYesNo As Integer, sFile As TextStream

Dim sArr() As String, nMap As Long, nRoom As Long

CommonDialog1.Filter = "TXT Files (*.txt)|*.txt"
CommonDialog1.DialogTitle = "Enter New File Name"
CommonDialog1.FileName = "NMR-ControlRoomList.txt"

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
sFile.WriteLine ("Control Room List -- " & Date & " @ " & Time)
sFile.WriteBlankLines (1)

For Each oLI In lvControlRoomList.ListItems
    
    str = """" & oLI.Text & ""","
    str = str & """" & oLI.SubItems(1) & ""","
    
    sArr = Split(oLI.Text, "/")
    If UBound(sArr()) = 1 Then
        nMap = Val(sArr(0))
        nRoom = Val(sArr(1))
        str = str & """" & GetControlRoomListByRoom(nMap, nRoom, 999, False) & """"
    Else
        str = str & """" & oLI.SubItems(2) & """"
    End If
    
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
lvControlRoomList.Width = Me.Width - 230
lvControlRoomList.Height = Me.Height - TITLEBAR_OFFSET - 860

If Not lvControlRoomList.ColumnHeaders.Count = 0 Then
    lvControlRoomList.ColumnHeaders(3).Width = lvControlRoomList.Width - 4500
End If

End Sub

Private Sub lvControlRoomList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If ColumnHeader.Index = 1 Or ColumnHeader.Index = 3 Then
    SortListView lvControlRoomList, ColumnHeader.Index, ldtNumber, lvControlRoomList.SortOrder
Else
    SortListView lvControlRoomList, ColumnHeader.Index, ldtString, lvControlRoomList.SortOrder
End If
End Sub


Public Sub CopyLine()

If lvControlRoomList.SelectedItem Is Nothing Then Exit Sub

Clipboard.clear
Clipboard.SetText lvControlRoomList.SelectedItem.Text & " -- " & lvControlRoomList.SelectedItem.SubItems(1)

End Sub

Private Sub lvControlRoomList_DblClick()
Dim sArr() As String, nMap As Long, nRoom As Long
On Error GoTo error:

If lvControlRoomList.SelectedItem Is Nothing Then Exit Sub

sArr = Split(lvControlRoomList.SelectedItem.Text, "/")
If UBound(sArr()) = 1 Then
    nMap = Val(sArr(0))
    nRoom = Val(sArr(1))
    frmRoom.GotoRoom nMap, nRoom
End If

out:
Exit Sub
error:
Call HandleError("CreateList")
Resume out:
End Sub

Private Sub lvControlRoomList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'If Button = 2 Then
'    PopupMenu frmMain.mnuMonsterIndexRightClick
'End If
End Sub
