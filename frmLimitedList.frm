VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20D5284F-7B23-4F0A-B8B1-6C9D18B64F1C}#1.0#0"; "exlimiter.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLimitedList 
   Caption         =   " Limited Items List"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7395
   Icon            =   "frmLimitedList.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5415
   ScaleWidth      =   7395
   Begin exlimiter.EL EL1 
      Left            =   6480
      Top             =   4500
      _ExtentX        =   1270
      _ExtentY        =   1270
   End
   Begin VB.CheckBox chkShops 
      Caption         =   "Scan Shops"
      Height          =   195
      Left            =   5160
      TabIndex        =   4
      Top             =   120
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkUsers 
      Caption         =   "Scan Users"
      Height          =   195
      Left            =   2580
      TabIndex        =   2
      Top             =   120
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkRooms 
      Caption         =   "Scan Rooms"
      Height          =   195
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Value           =   1  'Checked
      Width           =   1275
   End
   Begin VB.CommandButton cmdQ 
      Cancel          =   -1  'True
      Caption         =   "NOTE"
      Height          =   315
      Left            =   6480
      TabIndex        =   5
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save List"
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   60
      Width           =   1155
   End
   Begin VB.CommandButton cmdBuild 
      Caption         =   "&Build List"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1155
   End
   Begin MSComctlLib.ListView lvLimiteds 
      Height          =   4935
      Left            =   60
      TabIndex        =   6
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
      Left            =   6600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmLimitedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim LimitedItem() As Boolean
Dim bCancel As Boolean

Private Sub cmdQ_Click()
MsgBox "NOTE: Shop #109 is skipped in the search because this shop contains limited" & vbCrLf _
    & "items, but the shop iteself is not in the game.  If you've converted or started" & vbCrLf _
    & "using this shop, check it manually.", vbInformation
    
End Sub

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
lvLimiteds.ListItems.clear

frmProgressBar.sCaption = "Building Limited Item List"
frmProgressBar.lblCaption.Caption = "Building ..."
frmProgressBar.cmdCancel.Enabled = True
Call frmProgressBar.SetRange(CalcTotalRecords)
frmProgressBar.Show
frmMain.Enabled = False
DoEvents

LockWindowUpdate frmLimitedList.hwnd

ReDim LimitedItem(0 To 2500)

frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_ITEMS
Call AddItems
If bCancel Then GoTo ReEnable:
If lvLimiteds.ListItems.Count = 0 Then GoTo ReEnable:

frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_USERS
If chkUsers.Value = 1 Then Call ScanUsers
If bCancel Then GoTo ReEnable:

frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_SHOPS
If chkShops.Value = 1 Then Call ScanShops
If bCancel Then GoTo ReEnable:

frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_MP
If chkRooms.Value = 1 Then Call ScanRooms
If bCancel Then GoTo ReEnable:

Call CheckTotals

ReEnable:
Erase LimitedItem()
Unload frmProgressBar
frmMain.Enabled = True
LockWindowUpdate 0&
'Me.SetFocus

Exit Sub
error:
Call HandleError
Erase LimitedItem()
Unload frmProgressBar
frmMain.Enabled = True
LockWindowUpdate 0&
End Sub
Private Sub CheckTotals()
Dim oLI As ListItem, nLimit As Integer, nInUse As Integer
Dim y1 As Integer, y2 As Integer, x As Integer

For Each oLI In lvLimiteds.ListItems
    x = 1
    If Not oLI.SubItems(2) = "" Then
        If Not InStr(1, oLI.SubItems(1), "(") = 0 Then
            y1 = InStr(1, oLI.SubItems(1), "(")
            y2 = InStr(y1, oLI.SubItems(1), ")")
            nLimit = Val(Mid(oLI.SubItems(1), y1 + 1, y2 - y1))
        Else
            nLimit = 1
        End If
        
        nInUse = 0
checknext:
        x = x + 1
        If Not InStr(x, oLI.SubItems(2), "(") = 0 Then
            y1 = InStr(x, oLI.SubItems(2), "(")
            y2 = InStr(y1, oLI.SubItems(2), ")")
            nInUse = nInUse + Val(Mid(oLI.SubItems(2), y1 + 1, y2 - y1))
            x = y1 + 1
            GoTo checknext:
        End If
        
        If nInUse > nLimit Then oLI.ForeColor = vbRed
    End If
Next
End Sub
Private Function IsLimited(nNum As Long) As Boolean
    If nNum > UBound(LimitedItem()) Then ReDim Preserve LimitedItem(nNum)
    IsLimited = LimitedItem(nNum)
End Function
Private Sub ScanRooms()
Dim nStatus As Integer, oLI As ListItem, sTemp As String, nRec As Long
Dim y As Integer, x As Integer, ItemList() As Long, bMatch As Boolean, InvisList() As Long

nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Rooms: Could not get first record, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

ReDim ItemList(0 To 17, 1 To 2)
ReDim InvisList(0 To 15, 1 To 2)

Do While nStatus = 0 And Not bCancel
    RoomRowToStruct Roomdatabuf.buf
    
    nRec = nRec + 1
    frmProgressBar.lblPanel(1).Caption = nRec
    Call frmProgressBar.IncreaseProgress
    
    For x = 0 To 16
        If IsLimited(Roomrec.RoomItems(x)) Then
            bMatch = False
            For y = 0 To 16
                If ItemList(y, 1) = Roomrec.RoomItems(x) Then
                    bMatch = True
                    ItemList(y, 2) = ItemList(y, 2) + Roomrec.RoomItemQty(x) + 1
                    Exit For
                End If
            Next y
            
            If bMatch Then
                GoTo next1:
            Else
                y = 0
                Do Until ItemList(y, 1) = 0 Or y = 17
                    y = y + 1
                Loop
                ItemList(y, 1) = Roomrec.RoomItems(x)
                ItemList(y, 2) = Roomrec.RoomItemQty(x) + 1
            End If
        End If
next1:
    Next x
    
    For x = 0 To 14
        If IsLimited(Roomrec.InvisItems(x)) Then
            bMatch = False
            For y = 0 To 14
                If InvisList(y, 1) = Roomrec.InvisItems(x) Then
                    bMatch = True
                    InvisList(y, 2) = InvisList(y, 2) + Roomrec.InvisItemQty(x) + 1
                    Exit For
                End If
            Next y
            
            If bMatch Then
                GoTo next2:
            Else
                y = 0
                Do Until InvisList(y, 1) = 0 Or y = 15
                    y = y + 1
                Loop
                InvisList(y, 1) = Roomrec.InvisItems(x)
                InvisList(y, 2) = Roomrec.InvisItemQty(x) + 1
            End If
        End If
next2:
    Next x
    
    x = 0
    Do Until ItemList(x, 1) = 0 Or x = 17
        sTemp = ""
        Set oLI = lvLimiteds.FindItem(ItemList(x, 1), lvwText, , 0)
        If Not oLI Is Nothing Then
            sTemp = sTemp & "Room " & Roomrec.MapNumber & "/" & Roomrec.RoomNumber & " (" & ItemList(x, 2) & ")"
            If InStr(1, oLI.SubItems(2), sTemp) = 0 Then
                If Not oLI.SubItems(2) = "" Then sTemp = oLI.SubItems(2) & ", " & sTemp
                oLI.SubItems(2) = sTemp
            End If
        End If
        x = x + 1
    Loop
    
    x = 0
    Do Until InvisList(x, 1) = 0 Or x = 15
        sTemp = ""
        Set oLI = lvLimiteds.FindItem(InvisList(x, 1), lvwText, , 0)
        If Not oLI Is Nothing Then
            sTemp = sTemp & "Room: " & Roomrec.MapNumber & "/" & Roomrec.RoomNumber & " (" & InvisList(x, 2) & ")"
            If InStr(1, oLI.SubItems(2), sTemp) = 0 Then
                If Not oLI.SubItems(2) = "" Then sTemp = oLI.SubItems(2) & ", " & sTemp
                oLI.SubItems(2) = sTemp
            End If
        End If
        x = x + 1
    Loop
    
    Erase ItemList()
    Erase InvisList()
    ReDim ItemList(0 To 17, 1 To 2)
    ReDim InvisList(0 To 15, 1 To 2)
    nStatus = BTRCALL(BGETNEXT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

Erase ItemList()
Erase InvisList()
Set oLI = Nothing
End Sub
Private Sub ScanShops()
Dim nStatus As Integer, oLI As ListItem, x As Integer, sTemp As String

nStatus = BTRCALL(BGETFIRST, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Shops: Could not get first record, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

Do While nStatus = 0 And Not bCancel
    ShopRowToStruct Shopdatabuf.buf

    frmProgressBar.lblPanel(1).Caption = Shoprec.Number
    Call frmProgressBar.IncreaseProgress
    
    'shop 109 isn't in the game ... could screw up ppl if they've put this record # in the game
    If Shoprec.Number = 109 Then GoTo Skip:
    
    For x = 0 To 19
        If IsLimited(Shoprec.ShopItemNumber(x)) Then
            sTemp = ""
            If Not Shoprec.ShopItemNumber(x) = 0 And Not Shoprec.ShopNow(x) = 0 Then
                Set oLI = lvLimiteds.FindItem(Shoprec.ShopItemNumber(x), lvwText, , 0)
                If Not oLI Is Nothing Then
                    sTemp = sTemp & "Shop " & Shoprec.Number & " (" & Shoprec.ShopNow(x) & ")"
                    If InStr(1, oLI.SubItems(2), sTemp) = 0 Then
                        If Not oLI.SubItems(2) = "" Then sTemp = oLI.SubItems(2) & ", " & sTemp
                        oLI.SubItems(2) = sTemp
                    End If 'end instr
                End If 'end oli is nothing
            End If 'end shopitemnum
        End If
    Next

Skip:
    nStatus = BTRCALL(BGETNEXT, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

Set oLI = Nothing
End Sub
Private Sub ScanUsers()
Dim nStatus As Integer, oLI As ListItem, sTemp As String, nRec As Long
Dim y As Integer, x As Integer, ItemList() As Long, bMatch As Boolean

nStatus = BTRCALL(BGETFIRST, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Users: Could not get first record, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

ReDim ItemList(0 To 100, 1 To 2)

Do While nStatus = 0 And Not bCancel
    UserRowToStruct Userdatabuf.buf
    
    nRec = nRec + 1
    frmProgressBar.lblPanel(1).Caption = nRec
    Call frmProgressBar.IncreaseProgress
    
    'first create an index of all the items the user has
    For x = 0 To 99
        If IsLimited(Userrec.Item(x)) Then
            bMatch = False
            For y = 0 To 99
                'see if this item matches any of the items we've already seen, if so add +1
                If ItemList(y, 1) = Userrec.Item(x) Then
                    bMatch = True
                    ItemList(y, 2) = ItemList(y, 2) + 1
                    Exit For
                End If
            Next y
            
            If bMatch Then
                GoTo match:
            Else
                'figure out where the first unused spot in the array is
                y = 0
                Do Until ItemList(y, 1) = 0 Or y = 100
                    y = y + 1
                Loop
                
                'set the data
                ItemList(y, 1) = Userrec.Item(x)
                ItemList(y, 2) = 1
            End If
        End If
match:
    Next x
    
    'now fill the table
    x = 0
    Do Until ItemList(x, 1) = 0 Or x = 101
        sTemp = ""
        
        Set oLI = lvLimiteds.FindItem(ItemList(x, 1), lvwText, , 0)
        If Not oLI Is Nothing Then
            sTemp = sTemp & "User " & ClipNull(Userrec.FirstName) & "/" & ClipNull(Userrec.BBSName) & " (" & ItemList(x, 2) & ")"
            
            If Not oLI.SubItems(2) = "" Then sTemp = oLI.SubItems(2) & ", " & sTemp
            oLI.SubItems(2) = sTemp
        End If
        
        x = x + 1
    Loop
    
    Erase ItemList()
    ReDim ItemList(0 To 100, 1 To 2)
    nStatus = BTRCALL(BGETNEXT, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

Set oLI = Nothing
Erase ItemList()
End Sub
Private Sub AddItems()
Dim nStatus As Integer, oLI As ListItem, sName As String

nStatus = BTRCALL(BGETFIRST, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Items: Could not get first record, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

Do While nStatus = 0 And Not bCancel
    ItemRowToStruct Itemdatabuf.buf
    
    frmProgressBar.lblPanel(1).Caption = Itemrec.Number
    Call frmProgressBar.IncreaseProgress
    
    'no gamelimit, skip
    If Itemrec.GameLimit = 0 Then GoTo Skip:
    
    'no name, skip
    sName = ClipNull(Itemrec.Name)
    If sName = "" Then GoTo Skip:
    
    If UBound(LimitedItem()) < Itemrec.Number Then ReDim Preserve LimitedItem(Itemrec.Number)
    LimitedItem(Itemrec.Number) = True
    
    'add it
    Set oLI = lvLimiteds.ListItems.add()
    oLI.Text = Itemrec.Number
    
    If Itemrec.GameLimit > 1 Then sName = sName & " (" & Itemrec.GameLimit & ")"
    oLI.ListSubItems.add (1), "Name", sName

Skip:
    nStatus = BTRCALL(BGETNEXT, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

Set oLI = Nothing
End Sub
Public Sub ToggleStopBuild()
bCancel = True
End Sub

Private Sub AddColumnHeaders()

lvLimiteds.ColumnHeaders.clear
lvLimiteds.ColumnHeaders.add 1, "Number", "#", 600, lvwColumnLeft
lvLimiteds.ColumnHeaders.add 2, "Name", "Name (# limited)", 2600, lvwColumnCenter
lvLimiteds.ColumnHeaders.add 3, "Location", "Location (how many)", 3700, lvwColumnLeft

End Sub
Private Function CalcTotalRecords() As Long
On Error GoTo error:
Dim nStatus As Integer

CalcTotalRecords = 0

nStatus = BTRCALL(BSTAT, ItemPosBlock, DBStatDatabuf, Len(Itemdatabuf), 0, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    CalcTotalRecords = CalcTotalRecords + 1800
Else
    DBStatRowToStruct DBStatDatabuf.buf
    CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
End If

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
CommonDialog1.FileName = "NMR-Limiteds.txt"

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
sFile.WriteLine ("Limited list, " & Date & " @ " & Time)
sFile.WriteBlankLines (1)

For Each oLI In lvLimiteds.ListItems
    str = oLI.Text & " " & String(6 - Len(oLI.Text), ".") & " "
    str = str & oLI.SubItems(1) & " " & String(35 - Len(oLI.SubItems(1)), ".") & " "
    str = str & oLI.SubItems(2)
    
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


lvLimiteds.Width = Me.Width - 230
lvLimiteds.Height = Me.Height - TITLEBAR_OFFSET - 860

If Not lvLimiteds.ColumnHeaders.Count = 0 Then
    lvLimiteds.ColumnHeaders(3).Width = lvLimiteds.Width - 3700
End If

End Sub

Private Sub lvLimiteds_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If ColumnHeader.Index = 1 Then
    SortListView lvLimiteds, ColumnHeader.Index, ldtNumber, lvLimiteds.SortOrder
Else
    SortListView lvLimiteds, ColumnHeader.Index, ldtString, lvLimiteds.SortOrder
End If
End Sub


Public Sub CopyLine()

If lvLimiteds.SelectedItem Is Nothing Then Exit Sub

Clipboard.clear
Clipboard.SetText lvLimiteds.SelectedItem.Text & " -- " & lvLimiteds.SelectedItem.SubItems(1) & " -- " & lvLimiteds.SelectedItem.SubItems(2)

End Sub

Private Sub lvLimiteds_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu frmMain.mnuLimitedRightClick
End If
End Sub
