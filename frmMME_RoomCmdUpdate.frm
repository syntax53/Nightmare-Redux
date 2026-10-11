VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmMME_RoomCmdUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Exported MDB with Room Commands"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5415
   ScaleWidth      =   7230
   Begin VB.TextBox txtReport 
      Height          =   3735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1560
      Width           =   6990
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   435
      Left            =   5775
      TabIndex        =   5
      Top             =   1020
      Width           =   1335
   End
   Begin VB.CommandButton cmdFix 
      Caption         =   "&Fix"
      Default         =   -1  'True
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   1020
      Width           =   1335
   End
   Begin VB.CommandButton cmdBrowseDest 
      Caption         =   "Browse..."
      Height          =   315
      Left            =   6150
      TabIndex        =   3
      Top             =   555
      Width           =   960
   End
   Begin VB.TextBox txtDest 
      Height          =   285
      Left            =   1860
      TabIndex        =   2
      Top             =   570
      Width           =   4215
   End
   Begin VB.CommandButton cmdBrowseSource 
      Caption         =   "Browse..."
      Height          =   315
      Left            =   6150
      TabIndex        =   1
      Top             =   135
      Width           =   960
   End
   Begin VB.TextBox txtSource 
      Height          =   285
      Left            =   1860
      TabIndex        =   0
      Top             =   150
      Width           =   4215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5160
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label lblResult 
      Caption         =   "Successful"
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
      Left            =   1620
      TabIndex        =   9
      Top             =   1125
      Width           =   3435
   End
   Begin VB.Label lblDest 
      Caption         =   "Save fixed MDB as:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblSource 
      Caption         =   "Exported MDB:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   180
      Width           =   1695
   End
End
Attribute VB_Name = "frmMME_RoomCmdUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

'=====================================================================================
' Update Exported MDB with Room Commands
'
' Reads an MME export that already exists and writes a copy in which the items that
' room commands hand out name the commands and their rooms in "Obtained From":
'
'     Room(pry coffin) 17/278|17/287|17/2073-2096(18.6%)
'
' Everything is worked out from the export itself (Rooms "CMD", TBInfo "Action",
' the Items' references); the data files are not needed.  The selected export is
' only read; the result is always a new file.
' (Nightmare-Redux issue #21 / MMUD-Explorer issue #105)
'=====================================================================================

Dim dictTBText As Dictionary    'textblock number -> its text (TBInfo "Action")
Dim dictObtained As Dictionary  'item number -> "Obtained From", as it is going to be written
Dim dictItemFlag As Dictionary  'item number -> True when the export shows the item as in the game
Dim dictChanged As Dictionary   'the items whose "Obtained From" changed
Dim nEntriesAdded As Long       'Room(...) entries added
Dim nRoomsNamed As Long         'rooms those entries name
Dim nCut As Long                'entries that ran into the 2000 characters of the column
Dim sLastProposal As String     'the output name this window filled in by itself
Dim bBusy As Boolean

Private Sub Form_Load()
On Error Resume Next
lblResult.Caption = ""
txtReport.Text = "Select an MME export (an .mdb written by Export to MMUD Explorer), " _
    & "choose a name for the new file, and press Fix." & vbCrLf & vbCrLf _
    & "The selected export is only read.  An existing file is never replaced."
End Sub

Private Sub cmdClose_Click()
If bBusy Then Exit Sub
Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If bBusy Then Cancel = 1
End Sub

Private Sub cmdBrowseSource_Click()
On Error GoTo canceled:

CommonDialog1.CancelError = True
CommonDialog1.Filter = "MDB Files (*.mdb)|*.mdb|All Files (*.*)|*.*"
CommonDialog1.DialogTitle = "Select the exported MDB ..."
CommonDialog1.FileName = ""
CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then Exit Sub

txtSource.Text = CommonDialog1.FileName

'propose "<name> (room commands).mdb" next to it, unless the user typed a name of his own
If Len(Trim(txtDest.Text)) = 0 Or txtDest.Text = sLastProposal Then
    sLastProposal = ProposedOutput(txtSource.Text)
    txtDest.Text = sLastProposal
End If

canceled:
End Sub

Private Sub cmdBrowseDest_Click()
Dim sFile As String
On Error GoTo canceled:

CommonDialog1.CancelError = True
CommonDialog1.Filter = "MDB Files (*.mdb)|*.mdb"
CommonDialog1.DialogTitle = "Save the fixed MDB as ..."
CommonDialog1.FileName = Trim(txtDest.Text)
CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist
CommonDialog1.ShowSave
If CommonDialog1.FileName = "" Then Exit Sub

sFile = CommonDialog1.FileName
If LCase(Right(sFile, 4)) <> ".mdb" Then sFile = sFile & ".mdb"
txtDest.Text = sFile

canceled:
End Sub

'"C:\Exports\data-v1.11p.mdb" -> "C:\Exports\data-v1.11p (room commands).mdb"
Private Function ProposedOutput(ByVal sSource As String) As String
Dim nDot As Long, nSlash As Long

nSlash = InStrRev(sSource, "\")
nDot = InStrRev(sSource, ".")
If nDot > nSlash + 1 Then
    ProposedOutput = Left(sSource, nDot - 1) & " (room commands)" & Mid(sSource, nDot)
Else
    ProposedOutput = sSource & " (room commands).mdb"
End If

End Function

Private Sub cmdFix_Click()
Dim bOK As Boolean, sReport As String

If bBusy Then Exit Sub
bBusy = True

cmdFix.Enabled = False
cmdClose.Enabled = False
cmdBrowseSource.Enabled = False
cmdBrowseDest.Enabled = False
txtSource.Enabled = False
txtDest.Enabled = False
Screen.MousePointer = vbHourglass

lblResult.ForeColor = vbButtonText
lblResult.Caption = "Working ..."
txtReport.Text = ""
DoEvents

bOK = RunUpdate(Trim(txtSource.Text), Trim(txtDest.Text), sReport)

txtReport.Text = sReport
If bOK Then
    lblResult.ForeColor = &H8000&       'green
    lblResult.Caption = "Successful"
Else
    lblResult.ForeColor = &HC0&         'red
    lblResult.Caption = "Not successful - see below"
End If

Screen.MousePointer = vbDefault
txtSource.Enabled = True
txtDest.Enabled = True
cmdBrowseSource.Enabled = True
cmdBrowseDest.Enabled = True
cmdClose.Enabled = True
cmdFix.Enabled = True
bBusy = False

End Sub

'The whole job.  Never raises: True when the new file was written, and sReport says
'what was done or why not.
Private Function RunUpdate(ByVal sSource As String, ByVal sDest As String, ByRef sReport As String) As Boolean
Dim fso As FileSystemObject, DB As DAO.Database
Dim nStart As Single, sStep As String, sError As String
Dim bCopied As Boolean, nTagged As Long

On Error GoTo error:

Set fso = New FileSystemObject
nStart = Timer
sReport = ""

'--- what is refused before anything is read
If Len(sSource) = 0 Then
    sReport = "No exported MDB selected."
    GoTo out:
End If
If Not fso.FileExists(sSource) Then
    sReport = "File not found:" & vbCrLf & "  " & sSource
    GoTo out:
End If
If Len(sDest) = 0 Then
    sReport = "No name given for the fixed MDB."
    GoTo out:
End If
If InStr(1, sDest, "\") = 0 Then sDest = fso.BuildPath(fso.GetParentFolderName(sSource), sDest)
If LCase(Right(sDest, 4)) <> ".mdb" Then sDest = sDest & ".mdb"
If LCase(fso.GetAbsolutePathName(sSource)) = LCase(fso.GetAbsolutePathName(sDest)) Then
    sReport = "The fixed MDB must be a new file, not the export itself." & vbCrLf & "Nothing was written."
    GoTo out:
End If
If fso.FileExists(sDest) Then
    sReport = "This file already exists:" & vbCrLf & "  " & sDest & vbCrLf & vbCrLf _
        & "Choose a new file name.  Nothing was written."
    GoTo out:
End If
If Not fso.FolderExists(fso.GetParentFolderName(sDest)) Then
    sReport = "Folder not found:" & vbCrLf & "  " & fso.GetParentFolderName(sDest) & vbCrLf & "Nothing was written."
    GoTo out:
End If

'--- read the export and work out the entries (the export is opened read only)
sReport = "Reading   " & sSource & vbCrLf
sStep = "reading the export"
Set DB = DBEngine.OpenDatabase(sSource, False, True)
Call LoadExport(DB, sSource)
sStep = "scanning the room commands"
Call ScanExportRoomCommands(DB)
DB.Close
Set DB = Nothing

'--- the new file: an exact copy, in which only the changed items are rewritten
sStep = "copying the export"
FileCopy sSource, sDest
bCopied = True

sStep = "writing the new file"
Set DB = DBEngine.OpenDatabase(sDest)
If dictChanged.Count > 0 Then Call WriteItems(DB)
nTagged = CountTaggedItems(DB)
DB.Close
Set DB = Nothing

sReport = sReport & "Written   " & sDest & vbCrLf & vbCrLf
If dictChanged.Count > 0 Then
    sReport = sReport & "SUCCESS: Added " & nEntriesAdded & " room command source" & IIf(nEntriesAdded = 1, "", "s") _
        & " to " & dictChanged.Count & " item" & IIf(dictChanged.Count = 1, "", "s") & "." & vbCrLf
    sReport = sReport & "  room / command / item combinations covered: " & nRoomsNamed & vbCrLf
    If nCut > 0 Then
        sReport = sReport & "  WARNING: " & nCut & " entr" & IIf(nCut = 1, "y", "ies") _
            & " did not fit into the column's 2000 characters." & vbCrLf
    End If
Else
    sReport = sReport & "SUCCESS: Nothing to add: no item of this export gains a room command source." & vbCrLf
    sReport = sReport & "  (the export already has the entries, or no room command gives an item)" & vbCrLf
    sReport = sReport & "  The new file is a plain copy." & vbCrLf
End If
sReport = sReport & "  items with Room( entries in the new file: " & nTagged & vbCrLf
sReport = sReport & "  the selected export was not changed; all its other tables and cells were copied as they are" & vbCrLf
sReport = sReport & "  time: " & Format(Timer - nStart, "0.0") & " s" & vbCrLf

RunUpdate = True

out:
On Error Resume Next
Set DB = Nothing
Set fso = Nothing
Set dictTBText = Nothing
Set dictObtained = Nothing
Set dictItemFlag = Nothing
Set dictChanged = Nothing
Exit Function

error:
sError = "FAILED while " & sStep & ":" & vbCrLf & "  " & Err.Description & " (error " & Err.Number & ")"
Resume failed:

failed:
'(a new error here must not start the handler again)
On Error Resume Next
DB.Close
Set DB = Nothing
sReport = sReport & vbCrLf & sError & vbCrLf & vbCrLf
If bCopied Then
    Kill sDest
    If fso.FileExists(sDest) Then
        sReport = sReport & "The unfinished file could not be removed - delete it yourself:" & vbCrLf & "  " & sDest & vbCrLf
    Else
        sReport = sReport & "No updated file was written (the unfinished copy was removed)." & vbCrLf
    End If
Else
    sReport = sReport & "No updated file was written." & vbCrLf
End If
sReport = sReport & "The selected export was not changed." & vbCrLf
RunUpdate = False
GoTo out:

End Function

'Is there such a column?  (False also when there is no such table.)
Private Function HasColumn(DB As DAO.Database, ByVal sTable As String, ByVal sColumn As String) As Boolean
Dim fld As DAO.Field
On Error Resume Next
Set fld = DB.TableDefs(sTable).Fields(sColumn)
HasColumn = (Err.Number = 0)
Err.Clear
Set fld = Nothing
End Function

Private Function HasTable(DB As DAO.Database, ByVal sTable As String) As Boolean
Dim tdf As DAO.TableDef
On Error Resume Next
Set tdf = DB.TableDefs(sTable)
HasTable = (Err.Number = 0)
Err.Clear
Set tdf = Nothing
End Function

'Reads what the scan needs into memory.  Raises an error with a plain text when the file
'is not an MME export.
Private Sub LoadExport(DB As DAO.Database, ByVal sSource As String)
Dim rs As DAO.Recordset, vNeed As Variant, vCols As Variant, x As Long, y As Long
Dim sKey As String, sText As String, sObtained As String, bFlag As Boolean

vNeed = Array("Rooms", "Map Number|Room Number|CMD", _
    "TBInfo", "Number|LinkTo|Action", _
    "Items", "Number|Gettable|In Game|Obtained From|References")
For x = 0 To UBound(vNeed) Step 2
    If Not HasTable(DB, vNeed(x)) Then
        Err.Raise vbObjectError + 1001, "LoadExport", "This is not an MME export: it has no " & vNeed(x) & " table"
    End If
    vCols = Split(vNeed(x + 1), "|")
    For y = 0 To UBound(vCols)
        If Not HasColumn(DB, vNeed(x), vCols(y)) Then
            Err.Raise vbObjectError + 1002, "LoadExport", "The " & vNeed(x) & " table has no [" & vCols(y) _
                & "] column (an export of an older version?)"
        End If
    Next y
Next x

Set dictTBText = New Dictionary
Set dictObtained = New Dictionary
Set dictItemFlag = New Dictionary
Set dictChanged = New Dictionary
nEntriesAdded = 0
nRoomsNamed = 0
nCut = 0

'the text of the textblocks the export marked as executed
Set rs = DB.OpenRecordset("SELECT [Number], [Action] FROM TBInfo", dbOpenSnapshot)
Do While Not rs.EOF
    sText = "" & rs.Fields("Action")
    If sText = Chr(0) Then sText = ""
    dictTBText(CStr(CLng(rs.Fields("Number")))) = sText
    rs.MoveNext
Loop
rs.Close

'the items: their text, and whether the export shows them as "in the game" - in game,
'or with any reference (the Room( entries of an earlier run of this tool do not count)
Set rs = DB.OpenRecordset("SELECT [Number], [In Game], [Obtained From], [References] FROM Items", dbOpenSnapshot)
Do While Not rs.EOF
    sKey = CStr(CLng(rs.Fields("Number")))
    sObtained = "" & rs.Fields("Obtained From")
    dictObtained(sKey) = sObtained

    bFlag = False
    If Not IsNull(rs.Fields("In Game")) Then
        If CLng(rs.Fields("In Game")) <> 0 Then bFlag = True
    End If
    If Not IsNoText("" & rs.Fields("References")) Then bFlag = True
    If Not IsNoText(WithoutRoomCommandEntries(sObtained)) Then bFlag = True
    dictItemFlag(sKey) = bFlag

    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

End Sub

Private Function IsNoText(ByVal sText As String) As Boolean
IsNoText = (Len(sText) = 0 Or sText = Chr(0))
End Function

'"Shop #12, Room(dig grave) 1/816, NPC #5" -> "Shop #12, NPC #5"
Private Function WithoutRoomCommandEntries(ByVal sObtained As String) As String
Dim sArr() As String, x As Long, sOut As String, bAny As Boolean

If InStr(1, sObtained, "Room(") = 0 Then
    WithoutRoomCommandEntries = sObtained
    Exit Function
End If

sArr = Split(sObtained, ", ")
For x = 0 To UBound(sArr)
    If Left(sArr(x), 5) <> "Room(" Then
        If bAny Then sOut = sOut & ", "
        sOut = sOut & sArr(x)
        bAny = True
    End If
Next x
WithoutRoomCommandEntries = sOut

End Function

'The scan's "is this item in the game?".  An item the export does not list counts as in
'the game: the original scan had marked it when it met the number in a textblock.
Private Function ItemFlagged(ByVal nItem As Long) As Boolean
Dim sKey As String

sKey = CStr(nItem)
If dictItemFlag.Exists(sKey) Then
    ItemFlagged = dictItemFlag(sKey)
Else
    ItemFlagged = True
End If

End Function

Private Function TBText(ByVal nTBNumber As Long) As String
Dim sKey As String

sKey = CStr(nTBNumber)
If dictTBText.Exists(sKey) Then TBText = dictTBText(sKey)

End Function

'The pass itself: the rooms that have a command textblock, in the order a fresh export
'scans them (room number, then map number).
Private Sub ScanExportRoomCommands(DB As DAO.Database)
Dim rs As DAO.Recordset
Dim dictFound As Dictionary     'command textblock number -> Dictionary of what its commands give
Dim dictGroups As Dictionary    '"item<Tab>chance<Tab>commands" -> "map/room|map/room|..."
Dim dictGives As Dictionary
Dim vKey As Variant, sKey As String, sGroup As String, sRoom As String
Dim sParts() As String, sEntry As String, sRooms As String

Set dictFound = New Dictionary
Set dictGroups = New Dictionary

Set rs = DB.OpenRecordset("SELECT [Map Number], [Room Number], [CMD] FROM Rooms " _
    & "WHERE [CMD] > 0 ORDER BY [Room Number], [Map Number]", dbOpenSnapshot)
Do While Not rs.EOF
    'each command textblock is read once, however many rooms use it
    sKey = CStr(CLng(rs.Fields("CMD")))
    If Not dictFound.Exists(sKey) Then
        dictFound.Add sKey, GetRoomCommandGives(CLng(rs.Fields("CMD")))
        DoEvents
    End If
    Set dictGives = dictFound(sKey)

    If dictGives.Count > 0 Then
        sRoom = CLng(rs.Fields("Map Number")) & "/" & CLng(rs.Fields("Room Number"))
        For Each vKey In dictGives.Keys
            'rooms that give the same item with the same chance by the same commands share one entry
            sGroup = vKey & vbTab & dictGives(vKey)
            If dictGroups.Exists(sGroup) Then
                dictGroups(sGroup) = dictGroups(sGroup) & "|" & sRoom
            Else
                dictGroups.Add sGroup, sRoom
            End If
        Next vKey
    End If

    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

For Each vKey In dictGroups.Keys
    sParts = Split(CStr(vKey), vbTab)       '0 = item number, 1 = chance, 2 = commands
    sRooms = dictGroups(vKey)
    sEntry = "Room(" & sParts(2) & ") " & RoomListText(sRooms)
    If Len(sParts(1)) > 0 Then sEntry = sEntry & "(" & sParts(1) & "%)"
    Call AddObtained(Val(sParts(0)), sEntry, UBound(Split(sRooms, "|")) + 1)
Next vKey

Set dictGives = Nothing
Set dictGroups = Nothing
Set dictFound = Nothing

End Sub

'MarkItemInGame's way of adding a reference, on the text held in memory: an entry that is
'already there is not added again, and the text is cut at 2000 characters with "+".
Private Sub AddObtained(ByVal nItem As Long, ByVal sFrom As String, ByVal nRooms As Long)
Dim sKey As String, sTemp As String, sBefore As String

sKey = CStr(nItem)
If Not dictObtained.Exists(sKey) Then Exit Sub      'not an item of this export

sBefore = dictObtained(sKey)
sTemp = sBefore
If Not InStr(1, sTemp, sFrom) = 0 Then Exit Sub

If IsNoText(sTemp) Then
    sTemp = ""
Else
    sTemp = sTemp & ", "
End If

sTemp = sTemp & sFrom
If Len(sTemp) > 2000 Then sTemp = Left(sTemp, 1995) & "+" & Chr(0)

If sTemp = sBefore Then Exit Sub                    'the column was full already

dictObtained(sKey) = sTemp
dictChanged(sKey) = True
If Right(sTemp, 2) = "+" & Chr(0) Then
    nCut = nCut + 1
Else
    nEntriesAdded = nEntriesAdded + 1
    nRoomsNamed = nRoomsNamed + nRooms
End If

End Sub

Private Sub WriteItems(DB As DAO.Database)
Dim rs As DAO.Recordset, sKey As String

Set rs = DB.OpenRecordset("SELECT [Number], [Obtained From] FROM Items", dbOpenDynaset)
Do While Not rs.EOF
    sKey = CStr(CLng(rs.Fields("Number")))
    If dictChanged.Exists(sKey) Then
        rs.Edit
        rs.Fields("Obtained From") = dictObtained(sKey)
        rs.Update
    End If
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

End Sub

Private Function CountTaggedItems(DB As DAO.Database) As Long
Dim rs As DAO.Recordset

Set rs = DB.OpenRecordset("SELECT Count(*) AS n FROM Items WHERE [Obtained From] LIKE '*Room(*'", dbOpenSnapshot)
If Not rs.EOF Then CountTaggedItems = rs.Fields("n")
rs.Close
Set rs = Nothing

End Function

'What the commands of a room command textblock hand out.
'Returns a Dictionary:  "<item number><Tab><chance>" -> "<command>|<command>..."
'  chance ""     = the command gives the item itself (giveitem)
'  chance "18.6" = the item comes out of a "random" textblock, percent as for the chests
'  chance "?"    = it comes out of one, but the arithmetic gives nothing above 0
Private Function GetRoomCommandGives(ByVal nTBNumber As Long) As Dictionary
Dim dictOut As Dictionary
Dim sData As String, sLower As String, sLine As String, sCommand As String
Dim nLineStart As Long, nLineEnd As Long, nColon As Long, nSep As Long
Dim nPos As Long, nKeyPos As Long, nKeyLen As Long, nDigits As Long, nValue As Long
Dim nGiven() As Long, nGivenCount As Long, bGiven As Boolean
Dim nChest() As Currency, nNest As Long
Dim nChance As Currency, sChance As String, x As Long, z As Long

Set dictOut = New Dictionary
Set GetRoomCommandGives = dictOut

sData = TBText(nTBNumber)
If Len(sData) = 0 Then Exit Function
sLower = LCase(sData)

'a command textblock is made of lines "<command>:<action>:<action>..."
nLineStart = 1
Do While nLineStart <= Len(sLower)
    nLineEnd = InStr(nLineStart, sLower, Chr(10))
    If nLineEnd = 0 Then nLineEnd = Len(sLower) + 1
    sLine = Mid(sLower, nLineStart, nLineEnd - nLineStart)

    nColon = InStr(1, sLine, ":")
    If nColon = 0 Then GoTo next_line:

    'the command is taken from the original text (not the lower case copy)
    sCommand = CleanRoomCommand(Mid(sData, nLineStart, nColon - 1))
    If Len(sCommand) = 0 Then GoTo next_line:

    nGivenCount = 0
    ReDim nGiven(0)
    Erase nChest()
    ReDim nChest(1 To 3, 0) '1=number, 2=percent, 3=percent failure (as in ScanContainers)

    'look at what follows each ":" of the line
    nSep = nColon
    Do While nSep > 0
        nPos = nSep + 1
        Do While Mid(sLine, nPos, 1) = " "
            nPos = nPos + 1
        Loop

        nKeyPos = nPos
        nKeyLen = 0
        If Mid(sLine, nPos, 8) = "giveitem" Then
            nKeyLen = 8
        ElseIf Mid(sLine, nPos, 6) = "random" Then
            nKeyLen = 6
        End If

        'the word, at least one blank, then the number
        If nKeyLen > 0 And Mid(sLine, nPos + nKeyLen, 1) = " " Then
            nPos = nPos + nKeyLen
            Do While Mid(sLine, nPos, 1) = " "
                nPos = nPos + 1
            Loop

            nDigits = 0
            Do While IsRoomCmdDigit(Mid(sLine, nPos + nDigits, 1))
                nDigits = nDigits + 1
            Loop

            nValue = 0
            If nDigits > 0 And nDigits < 10 Then nValue = Val(Mid(sLine, nPos, nDigits))

            If nValue > 0 Then
                'an action behind a checkitem / roomitem / takeitem of an item that is
                'not in the game does not count (the same rule as in the export)
                If Not UpdTestItemFail(sLower, nLineStart + nKeyPos - 1) Then
                    If nKeyLen = 8 Then
                        'giveitem
                        bGiven = False
                        For x = 1 To nGivenCount
                            If nGiven(x) = nValue Then bGiven = True
                        Next x
                        If Not bGiven Then
                            nGivenCount = nGivenCount + 1
                            ReDim Preserve nGiven(nGivenCount)
                            nGiven(nGivenCount) = nValue
                        End If
                    Else
                        'random: the items of that textblock and of the ones it draws from
                        nNest = 0
                        Call UpdGetRandomTBItems(nChest(), nValue, nNest, 1)
                    End If
                End If
            End If
        End If

        nSep = InStr(nSep + 1, sLine, ":")
    Loop

    'what the command gives for sure ...
    For x = 1 To nGivenCount
        Call AddRoomCommandGive(dictOut, nGiven(x), "", sCommand)
    Next x

    '... and what it may give, with the chance worked out as for the chests
    For z = 0 To UBound(nChest(), 2)
        If nChest(1, z) > 0 Then
            bGiven = False
            For x = 1 To nGivenCount
                If nGiven(x) = nChest(1, z) Then bGiven = True
            Next x
            If Not bGiven Then
                nChance = Round(nChest(2, z) * 100, 1)
                If nChance <= 0 Then nChance = Round(nChest(2, z) * 100, 2)
                If nChance > 0 Then
                    sChance = "" & nChance
                Else
                    sChance = "?"
                End If
                Call AddRoomCommandGive(dictOut, CLng(nChest(1, z)), sChance, sCommand)
            End If
        End If
    Next z

next_line:
    nLineStart = nLineEnd + 1
Loop

Erase nGiven()
Erase nChest()

End Function

Private Sub AddRoomCommandGive(ByRef dictOut As Dictionary, ByVal nItem As Long, _
    ByVal sChance As String, ByVal sCommand As String)
Dim sKey As String

sKey = nItem & vbTab & sChance
If dictOut.Exists(sKey) Then
    'the same item with the same chance from another command of this textblock
    If InStr(1, "|" & dictOut(sKey) & "|", "|" & sCommand & "|") = 0 Then
        dictOut(sKey) = dictOut(sKey) & "|" & sCommand
    End If
Else
    dictOut.Add sKey, sCommand
End If

End Sub

Private Function IsRoomCmdDigit(ByVal sChar As String) As Boolean
Select Case sChar
    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9": IsRoomCmdDigit = True
End Select
End Function

'A room command as it is written into the reference: blanks squeezed, and the characters
'that separate references (", ") or delimit the command ("(", ")", "|") replaced.
Private Function CleanRoomCommand(ByVal sText As String) As String

sText = Replace(sText, vbTab, " ")
sText = Replace(sText, vbCr, " ")
sText = Replace(sText, vbLf, " ")
Do While InStr(1, sText, "  ") > 0
    sText = Replace(sText, "  ", " ")
Loop
sText = Trim(sText)

sText = Replace(sText, ",", ";")
sText = Replace(sText, "(", "[")
sText = Replace(sText, ")", "]")
sText = Replace(sText, "|", "/")

CleanRoomCommand = sText

End Function

'"17/278|17/287|17/2073|17/2074|17/2075" -> "17/278|17/287|17/2073-2075"
'(sorted by map, then room; three or more consecutive room numbers of a map are written first-last)
Private Function RoomListText(ByVal sRooms As String) As String
Dim sArr() As String, nMap() As Long, nRoom() As Long
Dim i As Long, j As Long, n As Long, sOut As String
Dim nTmpMap As Long, nTmpRoom As Long

sArr = Split(sRooms, "|")
n = UBound(sArr)
ReDim nMap(n)
ReDim nRoom(n)
For i = 0 To n
    j = InStr(1, sArr(i), "/")
    nMap(i) = Val(Left(sArr(i), j - 1))
    nRoom(i) = Val(Mid(sArr(i), j + 1))
Next i

'the rooms come room number first (1/1, 2/1, 3/1 ... 1/2, 2/2 ...): sort by map, then room
For i = 1 To n
    nTmpMap = nMap(i)
    nTmpRoom = nRoom(i)
    j = i - 1
    Do While j >= 0
        If nMap(j) > nTmpMap Or (nMap(j) = nTmpMap And nRoom(j) > nTmpRoom) Then
            nMap(j + 1) = nMap(j)
            nRoom(j + 1) = nRoom(j)
            j = j - 1
        Else
            Exit Do
        End If
    Loop
    nMap(j + 1) = nTmpMap
    nRoom(j + 1) = nTmpRoom
Next i

i = 0
Do While i <= n
    j = i
    Do While j < n
        If nMap(j + 1) = nMap(j) And nRoom(j + 1) = nRoom(j) + 1 Then
            j = j + 1
        Else
            Exit Do
        End If
    Loop

    If Len(sOut) > 0 Then sOut = sOut & "|"
    If j - i >= 2 Then
        sOut = sOut & nMap(i) & "/" & nRoom(i) & "-" & nRoom(j)
        i = j + 1
    Else
        sOut = sOut & nMap(i) & "/" & nRoom(i)
        i = i + 1
    End If
Loop

RoomListText = sOut

End Function

'TestItemFail of frmMME_Export, unchanged except that it asks ItemFlagged() instead of ItemInGame()
Private Function UpdTestItemFail(ByVal sWholeData As String, ByVal nCurrentPosition As Long) As Boolean
Dim sLook As String, sChar As String, sTest As String
Dim x As Long, y1 As Long, y2 As Long, z As Integer, nItem As Long

'make sTest just the the part of the data up until the current position
sTest = Mid(sWholeData, 1, nCurrentPosition - 1)

x = 1
NextLine:
'move x to the position of the first linebreak
If Not InStr(x, sTest, Chr(10)) = 0 Then
    x = InStr(x, sTest, Chr(10)) + 1
    GoTo NextLine:
End If

sTest = Mid(sWholeData, x, nCurrentPosition - x)

For z = 1 To 3
    
    x = 1
    'check for each type of item fail textblock
    Select Case z
        Case 1: sLook = "checkitem "
        Case 2: sLook = "roomitem "    'took out failitem because it means fail if you DO have it
        Case 3: sLook = "takeitem "
    End Select

checknext:
    'if there is one of those strings in there before our current position,
    'see if there is a linebreak between the string and the current position
    '(which would mean it was in a different command string)
    If Not InStr(x, sTest, sLook) = 0 Then

        x = InStr(x, sTest, sLook) 'sets x to the position of the failed string
        
        If z = 2 And x > 4 Then 'check for "failroomitem X" which we dont want to block
            If Mid(sTest, x - 4, 4) = "fail" Then
                x = x + 1
                GoTo checknext:
            End If
        End If
        
        y1 = x + Len(sLook) 'len of string searching (to position y1 at first number)
        y2 = 0
nextnumber:
        sChar = Mid(sTest, y1 + y2, 1)
        Select Case sChar
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
                If Not y1 + y2 - 1 = Len(sTest) Then
                    y2 = y2 + 1
                    GoTo nextnumber:
                End If
            Case Else:
        End Select
        
        If y2 = 0 Then
            'if there were no numbers after the string
            x = y1
            GoTo checknext:
        End If
        
        nItem = Val(Mid(sTest, y1, y2))
    
        'if the item is in the game then we're clear, check for any more matches
        If ItemFlagged(nItem) Then
            x = y1
            GoTo checknext:
        Else
            'item not in game, line fails
            UpdTestItemFail = True
            Exit Function
        End If
    End If
Next z

End Function

'GetRandomTBItems of frmMME_Export, unchanged except that the text comes from TBText() and the
'spell part (bCheckSpells) is left out
Private Sub UpdGetRandomTBItems(ByRef nReturnArray() As Currency, ByVal nTBNumber As Long, _
    ByRef nNest As Long, ByVal nPercentMod As Currency)
Dim sData As String, nDataPos As Long, x As Long, y As Long, nLineStart As Long
Dim nPer1 As Long, nPer2 As Long, sLine As String, nValue As Long, nPercent As Currency
Dim nItemArray() As Currency
On Error GoTo error:

If nNest > 5 Then Exit Sub

'the text comes from the export (TBInfo "Action"), not from the textblock file
sData = TBText(nTBNumber)
If Len(sData) < 5 Then Exit Sub

nDataPos = 1

nNest = nNest + 1

If nPercentMod <= 0 Then nPercentMod = 1
ReDim nItemArray(1 To 2, 0) '1=number, 2=percent

'first we collect all the items and total their %'s
Do While nDataPos < Len(sData)
    x = InStr(nDataPos, sData, ":")
    If x > nDataPos Then
        nPer1 = Val(Mid(sData, nDataPos, x - nDataPos))
        nPercent = (nPer1 - nPer2) / 100
        nPer2 = nPer1
        
        nDataPos = x + 1
        
        x = InStr(nDataPos, sData, Chr(10))
        If x <= 0 Then x = Len(sData)
        sLine = LCase(Mid(sData, nDataPos, x - nDataPos))
        nLineStart = nDataPos
        nDataPos = x
        
        y = 1
check_give_again:
        y = InStr(y, sLine, "giveitem ")
        If y > 0 Then
            nValue = ExtractValueFromString(Mid(sLine, y), "giveitem ")
            
            For x = 0 To UBound(nItemArray(), 2)
                If nItemArray(1, x) = nValue Then
                    nItemArray(2, x) = nItemArray(2, x) + nPercent
                    x = -1
                    Exit For
                End If
            Next x
            If x >= 0 Then
                x = UBound(nItemArray(), 2) + 1
                ReDim Preserve nItemArray(1 To 2, x)
                nItemArray(1, x) = nValue
                nItemArray(2, x) = nPercent
            End If
            
            y = y + 1
            GoTo check_give_again:
        End If
        
        y = 1
check_random_again:
        y = InStr(y, sLine, "random ")
        If y > 0 Then
            
            nValue = ExtractValueFromString(Mid(sLine, y), "random ")
            If nValue > 0 Then
                Call UpdGetRandomTBItems(nReturnArray(), nValue, nNest, (nPercent * nPercentMod))
            End If
            
            y = y + 1
            GoTo check_random_again:
        End If
    ElseIf x = 0 Then
        Exit Do
    Else
        nDataPos = nDataPos + 1
    End If
Loop

'then we put the collected items into the chest array
'...this is actually sort of unecessary i found out afterwards
For y = 0 To UBound(nItemArray(), 2)
    If nItemArray(1, y) > 0 Then
        nPercent = nItemArray(2, y)
        
        For x = 0 To UBound(nReturnArray(), 2)
            If nReturnArray(1, x) = nItemArray(1, y) Then
                'If nReturnArray(3, x) = 0 Then nReturnArray(3, x) = 1
                nReturnArray(2, x) = nReturnArray(2, x) + _
                    (nReturnArray(3, x) * nPercent * nPercentMod)
                nReturnArray(3, x) = nReturnArray(3, x) * (1 - nPercent)
                x = -1
                Exit For
            End If
        Next x
        If x >= 0 Then
            x = UBound(nReturnArray(), 2) + 1
            ReDim Preserve nReturnArray(1 To 3, x)
            nReturnArray(1, x) = nItemArray(1, y)
            nReturnArray(2, x) = nPercent * nPercentMod
            nReturnArray(3, x) = 1 - nReturnArray(2, x)
        End If
        
    End If
Next y

nNest = nNest - 1

Erase nItemArray()

Exit Sub
error:
Call HandleError("UpdGetRandomTBItems-#" & nTBNumber)
Erase nItemArray()
End Sub
