VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20D5284F-7B23-4F0A-B8B1-6C9D18B64F1C}#1.0#0"; "exlimiter.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmItemFind 
   Caption         =   " Find Item"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7395
   Icon            =   "frmItemFind.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5415
   ScaleWidth      =   7395
   Begin VB.TextBox txtItemNumber 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox txtItemName 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   3540
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   3795
   End
   Begin exlimiter.EL EL1 
      Left            =   6480
      Top             =   4500
      _ExtentX        =   1270
      _ExtentY        =   1270
   End
   Begin VB.CheckBox chkShops 
      Caption         =   "Scan Shops"
      Height          =   195
      Left            =   6120
      TabIndex        =   3
      Top             =   420
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkUsers 
      Caption         =   "Scan Users"
      Height          =   195
      Left            =   3540
      TabIndex        =   1
      Top             =   420
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkRooms 
      Caption         =   "Scan Rooms"
      Height          =   195
      Left            =   4800
      TabIndex        =   2
      Top             =   420
      Value           =   1  'Checked
      Width           =   1275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save List"
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdBuild 
      Caption         =   "&Build List"
      Height          =   615
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   1155
   End
   Begin MSComctlLib.ListView lvItems 
      Height          =   4635
      Left            =   60
      TabIndex        =   6
      Top             =   720
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   8176
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
      Left            =   6600
      Top             =   300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Item to Find:"
      Height          =   195
      Left            =   1320
      TabIndex        =   8
      Top             =   60
      Width           =   975
   End
End
Attribute VB_Name = "frmItemFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim nItemToFind As Long
Dim bCancel As Boolean

Private Sub Form_Load()
On Error Resume Next

With EL1
    .FormInQuestion = Me
    .MINWIDTH = 500
    .MINHEIGHT = 200
    .EnableLimiter = True
End With

Me.Width = 7515
Me.Height = 5925

Call AddColumnHeaders

End Sub
Private Sub cmdBuild_Click()
On Error GoTo error:

bCancel = False
lvItems.ListItems.clear

nItemToFind = Val(txtItemNumber.Text)
If nItemToFind <= 0 Then Exit Sub

frmProgressBar.sCaption = "Find Item"
frmProgressBar.lblCaption.Caption = "Searching ..."
frmProgressBar.cmdCancel.Enabled = True
Call frmProgressBar.SetRange(CalcTotalRecords)
frmProgressBar.Show
frmMain.Enabled = False
DoEvents

LockWindowUpdate Me.hwnd

frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & "user2.dat"
If chkUsers.Value = 1 Then Call ScanUsers
If bCancel Then GoTo ReEnable:

frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & "shop2.dat"
If chkShops.Value = 1 Then Call ScanShops
If bCancel Then GoTo ReEnable:

frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & "mp002.dat"
If chkRooms.Value = 1 Then Call ScanRooms
If bCancel Then GoTo ReEnable:


ReEnable:
If lvItems.ListItems.Count > 0 Then Call SortListView(lvItems, 1, ldtString, True)

On Error Resume Next
Unload frmProgressBar
frmMain.Enabled = True
DoEvents
LockWindowUpdate 0&
Me.SetFocus
Me.refresh

Exit Sub
error:
Call HandleError
Resume ReEnable:
End Sub
'Private Sub CheckTotals()
'Dim oLI As ListItem, nLimit As Integer, nInUse As Integer
'Dim y1 As Integer, y2 As Integer, x As Integer
'
'For Each oLI In lvItems.ListItems
'    x = 1
'    If Not oLI.SubItems(2) = "" Then
'        If Not InStr(1, oLI.SubItems(1), "(") = 0 Then
'            y1 = InStr(1, oLI.SubItems(1), "(")
'            y2 = InStr(y1, oLI.SubItems(1), ")")
'            nLimit = Val(Mid(oLI.SubItems(1), y1 + 1, y2 - y1))
'        Else
'            nLimit = 1
'        End If
'
'        nInUse = 0
'checknext:
'        x = x + 1
'        If Not InStr(x, oLI.SubItems(2), "(") = 0 Then
'            y1 = InStr(x, oLI.SubItems(2), "(")
'            y2 = InStr(y1, oLI.SubItems(2), ")")
'            nInUse = nInUse + Val(Mid(oLI.SubItems(2), y1 + 1, y2 - y1))
'            x = y1 + 1
'            GoTo checknext:
'        End If
'
'        If nInUse > nLimit Then oLI.ForeColor = vbRed
'    End If
'Next
'End Sub

Private Sub ScanRooms()
Dim nStatus As Integer, oLI As ListItem, nRec As Long
Dim x As Integer, nCount As Integer

nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Rooms: Could not get first record, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

Do While nStatus = 0 And Not bCancel
    RoomRowToStruct Roomdatabuf.buf
    
    nRec = nRec + 1
    frmProgressBar.lblPanel(1).Caption = nRec
    Call frmProgressBar.IncreaseProgress
    
    nCount = 0
    For x = 0 To 16
        If Roomrec.RoomItems(x) = nItemToFind Then
            nCount = nCount + Roomrec.RoomItemQty(x) + 1
        End If
    Next x
    If nCount > 0 Then
        Set oLI = lvItems.ListItems.add(, , "Room(vis): " & ClipNull(Roomrec.Name) & " (" & Roomrec.MapNumber & "/" & Roomrec.RoomNumber & ")")
        oLI.SubItems(1) = nCount
        oLI.Tag = Roomrec.MapNumber & "/" & Roomrec.RoomNumber
    End If
    
    nCount = 0
    For x = 0 To 14
        If Roomrec.InvisItems(x) = nItemToFind Then
            nCount = nCount + Roomrec.InvisItemQty(x) + 1
        End If
    Next x
    If nCount > 0 Then
        Set oLI = lvItems.ListItems.add(, , "Room(hid): " & ClipNull(Roomrec.Name) & " (" & Roomrec.MapNumber & "/" & Roomrec.RoomNumber & ")")
        oLI.SubItems(1) = nCount
        oLI.Tag = Roomrec.MapNumber & "/" & Roomrec.RoomNumber
    End If
    
    nStatus = BTRCALL(BGETNEXT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

Set oLI = Nothing
End Sub
Private Sub ScanShops()
Dim nStatus As Integer, oLI As ListItem, x As Integer, nCount As Integer

nStatus = BTRCALL(BGETFIRST, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Shops: Could not get first record, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

Do While nStatus = 0 And Not bCancel
    ShopRowToStruct Shopdatabuf.buf

    frmProgressBar.lblPanel(1).Caption = Shoprec.Number
    Call frmProgressBar.IncreaseProgress
    
    nCount = 0
    For x = 0 To 19
        If Shoprec.ShopItemNumber(x) = nItemToFind And Not Shoprec.ShopNow(x) = 0 Then
            nCount = nCount + Shoprec.ShopNow(x)
        End If
    Next x
    
    If nCount > 0 Then
        Set oLI = lvItems.ListItems.add(, , "Shop: " & ClipNull(Shoprec.Name) & " (" & Shoprec.Number & ")")
        oLI.SubItems(1) = nCount
        oLI.Tag = Shoprec.Number
    End If
    
Skip:
    nStatus = BTRCALL(BGETNEXT, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

Set oLI = Nothing
End Sub
Private Sub ScanUsers()
Dim nStatus As Integer, oLI As ListItem, sTemp As String, nRec As Long
Dim x As Integer, nCount As Integer

nStatus = BTRCALL(BGETFIRST, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Users: Could not get first record, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

Do While nStatus = 0 And Not bCancel
    nCount = 0
    UserRowToStruct Userdatabuf.buf
    
    nRec = nRec + 1
    frmProgressBar.lblPanel(1).Caption = nRec
    Call frmProgressBar.IncreaseProgress
    
    For x = 0 To 99
        If Userrec.Item(x) = nItemToFind Then nCount = nCount + 1
    Next x
    
    'now fill the table
    If nCount > 0 Then
        sTemp = ClipNull(Userrec.FirstName)
        Set oLI = lvItems.ListItems.add(, , "User: " & sTemp & "/" & ClipNull(Userrec.BBSName))
        oLI.SubItems(1) = nCount
        oLI.Tag = sTemp
    End If

    nStatus = BTRCALL(BGETNEXT, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

Set oLI = Nothing
End Sub

Public Sub ToggleStopBuild()
bCancel = True
End Sub

Private Sub AddColumnHeaders()

lvItems.ColumnHeaders.clear
lvItems.ColumnHeaders.add 1, "Location", "Location", 3700, lvwColumnLeft
lvItems.ColumnHeaders.add 2, "QTY", "QTY", 1000, lvwColumnLeft

End Sub
Private Function CalcTotalRecords() As Long
On Error GoTo error:
Dim nStatus As Integer

CalcTotalRecords = 0

'nStatus = BTRCALL(BSTAT, ItemPosBlock, DBStatDatabuf, Len(Itemdatabuf), 0, KEY_BUF_LEN, 0)
'If Not nStatus = 0 Then
'    CalcTotalRecords = CalcTotalRecords + 1800
'Else
'    DBStatRowToStruct DBStatDatabuf.buf
'    CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
'End If

nStatus = BTRCALL(BSTAT, ShopPosBlock, DBStatDatabuf, Len(Shopdatabuf), 0, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    CalcTotalRecords = CalcTotalRecords + 200
Else
    DBStatRowToStruct DBStatDatabuf.buf
    CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
End If

If chkRooms.Value = 1 Then
    nStatus = BTRCALL(BSTAT, RoomPosBlock, DBStatDatabuf, Len(Roomdatabuf), 0, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        CalcTotalRecords = CalcTotalRecords + 30000
    Else
        DBStatRowToStruct DBStatDatabuf.buf
        CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
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


If CalcTotalRecords <= 0 Then CalcTotalRecords = 1
'If CalcTotalRecords > 32767 Then CalcTotalRecords = 32767

Exit Function

error:
Call HandleError
End Function

Private Sub cmdSave_Click()
On Error GoTo error:
Dim oLI As ListItem, str As String, x As Integer
Dim fso As FileSystemObject, nYesNo As Integer, sFile As TextStream

CommonDialog1.Filter = "TXT Files (*.txt)|*.txt"
CommonDialog1.DialogTitle = "Enter New File Name"
CommonDialog1.FileName = "NMR-ItemFind.txt"

On Error GoTo canceled:
CommonDialog1.ShowSave
If CommonDialog1.FileName = "" Then GoTo canceled:

On Error GoTo error:

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(CommonDialog1.FileName) Then
    nYesNo = MsgBox("File Exists, Overwrite?", vbYesNo, "Overwrite?")
    If nYesNo = vbYes Then
        fso.DeleteFile (CommonDialog1.FileName)
    Else
        GoTo canceled:
    End If
End If

Set sFile = fso.OpenTextFile(CommonDialog1.FileName, ForWriting, True)
sFile.WriteLine ("Item Find Results, " & Date & " @ " & Time)
sFile.WriteBlankLines (1)

For Each oLI In lvItems.ListItems
    str = oLI.Text
    str = str & " - " & oLI.SubItems(1)
    
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


lvItems.Width = Me.Width - 230
lvItems.Height = Me.Height - TITLEBAR_OFFSET - 1150

If Not lvItems.ColumnHeaders.Count = 0 Then
    lvItems.ColumnHeaders(1).Width = lvItems.Width - 1500
End If

End Sub

Private Sub lvItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If ColumnHeader.Index = 2 Then
    SortListView lvItems, ColumnHeader.Index, ldtNumber, lvItems.SortOrder
Else
    SortListView lvItems, ColumnHeader.Index, ldtString, lvItems.SortOrder
End If
End Sub


Public Sub CopyLine()

If lvItems.SelectedItem Is Nothing Then Exit Sub

Clipboard.clear
Clipboard.SetText lvItems.SelectedItem.Text & " -- " & lvItems.SelectedItem.SubItems(1)

End Sub

Private Sub lvItems_DblClick()
Dim oLI As ListItem, tExits As RoomExitType
On Error GoTo error:

If lvItems.SelectedItem Is Nothing Then Exit Sub

Select Case LCase(Left(lvItems.SelectedItem.Text, 4))
    Case "user":
        Load frmUser
        frmUser.SetFocus
        
        Set oLI = frmUser.lvDatabase.FindItem(lvItems.SelectedItem.Tag, lvwText)
        If oLI Is Nothing Then
            MsgBox "User not found in user editor.", vbInformation
        Else
            Set frmUser.lvDatabase.SelectedItem = oLI
            frmUser.lvDatabase.SelectedItem.EnsureVisible
            Call frmUser.lvDatabase_ItemClick(oLI)
        End If
        
    Case "shop":
        Load frmShop
        frmShop.SetFocus
        
        Set oLI = frmShop.lvDatabase.FindItem(lvItems.SelectedItem.Tag, lvwText)
        If oLI Is Nothing Then
            MsgBox "Shop not found in shop editor.", vbInformation
        Else
            Set frmShop.lvDatabase.SelectedItem = oLI
            frmShop.lvDatabase.SelectedItem.EnsureVisible
            Call frmShop.lvDatabase_ItemClick(oLI)
        End If
        
    Case "room":
        tExits = ExtractMapRoom(lvItems.SelectedItem.Tag)
        If tExits.Map > 0 And tExits.Room > 0 Then
            Call frmRoom.GotoRoom(tExits.Map, tExits.Room)
        End If
        
End Select

out:
Set oLI = Nothing
Exit Sub
error:
Call HandleError("lvItems_DblClick")
Resume out:

End Sub

Private Sub lvItems_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu frmMain.mnuItemFindRightClick
End If
End Sub

Private Sub Text2_Change()

End Sub

Private Sub txtItemNumber_Change()
txtItemName.Text = GetItemName(Val(txtItemNumber.Text))
End Sub

Private Sub txtItemNumber_GotFocus()
Call SelectAll(txtItemNumber)
End Sub
