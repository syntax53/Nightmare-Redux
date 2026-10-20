Attribute VB_Name = "modMMEMessages"
Option Base 0
Option Explicit

'MME export: builds the Messages table (client-relevant game message lines).
'Spec: PLAN-mme-messages-export.md, "Revision 2" (R2.1-R2.17; later sections override earlier ones).
'Order: items (learnable / item-cast spell lists, item shields), rooms (traps, room and exit spells), spells,
'item shield rows, monsters, static rows (chase lines, monster entrances).
'Kinds: Cast, Self, Other. Tokens: {source} {target} {dmg} {attacker}.

Private mFso As Scripting.FileSystemObject
Private mLog As Scripting.TextStream
Private mdictLogged As Scripting.Dictionary
Private msLogPath As String
Private mnWarnings As Long
Private mbLogBroken As Boolean
Private mbFailed As Boolean
Private mbGmud As Boolean
Private mbHaveLearnable As Boolean    'dictLearn holds the export's Learnable spells (cross-referenced)

'hard-coded (never exported) formats, applied to the substituted text
Private mReP1 As Object    'own standard cast line
Private mReP2 As Object    'standard incoming cast line
Private mReP3 As Object    'standard incoming damage lines
Private mReP4 As Object    'third-party damage observation / shield owner line
Private mReP5 As Object    'standard monster hit on me
Private mReP6 As Object    'standard monster hit on someone else
Private mReP7 As Object    'standard own cast line without an amount ("You cast/invoke/sing ...")

'spells cast with the player as caster without the player choosing them (textblocks, cast exits, spell-trap
'exits): the player sees line 1. Value "Trap" for spell-trap exits, "" otherwise.
Private mdictSelfCast As Scripting.Dictionary
'item damage shields: ShockMsg number -> "|" list of the names of the items that carry it
Private mdictItemShield As Scripting.Dictionary

'in-memory row store: rows are collected by the passes, then merged and written once by FinalizeRows
Private mRowKind() As String
Private mRowName() As String
Private mRowMsg() As String
Private mRowEnd() As String
Private mRowEff() As String
Private mRowDrop() As Boolean
Private mRowCount As Long
Private mdictRows As Scripting.Dictionary    'row key -> row index

'dictLearnable: numbers (as strings) of the spells the export's cross-referencing marked Learnable; Nothing or empty
'when no cross-reference data exists, in which case a spell with a magery type counts as player-castable.
'dictTextblockCast: numbers (as strings) of the spells textblocks cast (cross-referencing, plus Spells "Casted By");
'Nothing when unknown.
Public Function MME_ExportMessages(tabMessages As DAO.Recordset, ByVal sLogPath As String, ByRef bStop As Boolean, _
    Optional ByVal dictLearnable As Scripting.Dictionary, Optional ByVal dictTextblockCast As Scripting.Dictionary) As Long
Dim dictRoomSpells As Scripting.Dictionary, vKey As Variant
Dim dictLearn As Scripting.Dictionary, dictItemCast As Scripting.Dictionary
On Error GoTo error:

MME_ExportMessages = 0
If Not InitState(sLogPath) Then
    Call CleanupState
    MME_ExportMessages = -1
    Exit Function
End If

mbGmud = (frmMME_Export.chkLegit(1).Value = 1)

Set dictRoomSpells = New Scripting.Dictionary
dictRoomSpells.CompareMode = vbBinaryCompare
Set dictLearn = New Scripting.Dictionary
dictLearn.CompareMode = vbBinaryCompare
Set dictItemCast = New Scripting.Dictionary
dictItemCast.CompareMode = vbBinaryCompare

mbHaveLearnable = False
If Not dictLearnable Is Nothing Then
    If dictLearnable.Count > 0 Then
        mbHaveLearnable = True
        For Each vKey In dictLearnable.Keys
            If Not dictLearn.Exists(CStr(vKey)) Then dictLearn.Add CStr(vKey), True
        Next vKey
    End If
End If

If Not dictTextblockCast Is Nothing Then
    For Each vKey In dictTextblockCast.Keys
        If Not mdictSelfCast.Exists(CStr(vKey)) Then mdictSelfCast.Add CStr(vKey), ""
    Next vKey
End If

tryagain:
If tabMessages.RecordCount <> 0 Then
    tabMessages.MoveFirst
    tabMessages.Delete
    GoTo tryagain:
End If

Call ExportItemScan(dictLearn, dictItemCast, bStop)
If bStop = False And mbFailed = False Then Call ExportRoomMessages(dictRoomSpells, bStop)
If bStop = False And mbFailed = False Then Call ExportSpellMessages(dictRoomSpells, dictLearn, dictItemCast, bStop)
If bStop = False And mbFailed = False Then Call ExportItemShieldRows
If bStop = False And mbFailed = False Then Call ExportMonsterMessages(bStop)
If bStop = False And mbFailed = False Then Call ExportStaticRows
If bStop = False And mbFailed = False Then Call FinalizeRows(tabMessages)

If mbFailed Then
    MME_ExportMessages = -1
Else
    MME_ExportMessages = mnWarnings
End If

Call CleanupState
Set dictRoomSpells = Nothing
Set dictLearn = Nothing
Set dictItemCast = Nothing

Exit Function
error:
Call HandleError("MME_ExportMessages")
Call CleanupState
Set dictRoomSpells = Nothing
Set dictLearn = Nothing
Set dictItemCast = Nothing
MME_ExportMessages = -1
End Function

'---------------------------------------------------------------------------
' State, logging, regex
'---------------------------------------------------------------------------

Private Function InitState(ByVal sLogPath As String) As Boolean
On Error GoTo error:

mbFailed = False
mbLogBroken = False
mbGmud = False
mnWarnings = 0
msLogPath = sLogPath
Set mLog = Nothing

Set mFso = CreateObject("Scripting.FileSystemObject")
If Not sLogPath = "" Then
    If mFso.FileExists(sLogPath) Then mFso.DeleteFile sLogPath, True
End If

Set mdictLogged = New Scripting.Dictionary
mdictLogged.CompareMode = vbBinaryCompare

mRowCount = 0
ReDim mRowKind(255)
ReDim mRowName(255)
ReDim mRowMsg(255)
ReDim mRowEnd(255)
ReDim mRowEff(255)
ReDim mRowDrop(255)
Set mdictRows = New Scripting.Dictionary
mdictRows.CompareMode = vbBinaryCompare
Set mdictSelfCast = New Scripting.Dictionary
mdictSelfCast.CompareMode = vbBinaryCompare
Set mdictItemShield = New Scripting.Dictionary
mdictItemShield.CompareMode = vbBinaryCompare

Set mReP1 = NewRegExp("^You cast .+ (on|at|upon) .*\{dmg\} (damage|healing|hit points)!?$")
Set mReP2 = NewRegExp("^\{source\} casts .+ (on|at|upon) you.*\{dmg\} damage!?$")
Set mReP3 = NewRegExp("^You are .* for {dmg} damage!?$|^You take {dmg} damage|you for {dmg} damage")
Set mReP4 = NewRegExp("^{target} .*{dmg}.* damage|^{attacker} .*{dmg}.* damage")
Set mReP5 = NewRegExp("^(The |An? )?\{source\} [a-z-]+ you for \{dmg\} damage!$")
Set mReP6 = NewRegExp("{target} for {dmg} damage")
Set mReP7 = NewRegExp("^You (cast|invoke|sing) ")

InitState = Not mbFailed

Exit Function
error:
Call HandleError("InitState")
mbFailed = True
InitState = False
End Function

Private Sub CleanupState()
'resume next on purpose: every object must be released even if Close fails
On Error Resume Next

If Not mLog Is Nothing Then mLog.Close
Set mLog = Nothing
Set mFso = Nothing
Set mdictLogged = Nothing
Set mReP1 = Nothing
Set mReP2 = Nothing
Set mReP3 = Nothing
Set mReP4 = Nothing
Set mReP5 = Nothing
Set mReP6 = Nothing
Set mReP7 = Nothing

Set mdictRows = Nothing
Set mdictSelfCast = Nothing
Set mdictItemShield = Nothing
Erase mRowKind
Erase mRowName
Erase mRowMsg
Erase mRowEnd
Erase mRowEff
Erase mRowDrop
mRowCount = 0

Err.clear
End Sub

Private Function NewRegExp(ByVal sPattern As String) As Object
Dim oRe As Object
On Error GoTo error:

Set oRe = CreateObject("VBScript.RegExp")
oRe.Pattern = sPattern
oRe.IgnoreCase = False
oRe.Global = False
Set NewRegExp = oRe
Set oRe = Nothing

Exit Function
error:
Call HandleError("NewRegExp")
Set oRe = Nothing
mbFailed = True
End Function

Private Function IsStandardForm(oRe As Object, ByVal sText As String) As Boolean
On Error GoTo error:

IsStandardForm = oRe.Test(sText)

Exit Function
error:
Call HandleError("IsStandardForm")
mbFailed = True
End Function

Private Sub LogWarning(ByVal sText As String)
On Error GoTo error:

mnWarnings = mnWarnings + 1
If mbLogBroken Then Exit Sub
If mLog Is Nothing Then Set mLog = mFso.CreateTextFile(msLogPath, True)
mLog.WriteLine sText

Exit Sub
error:
Call HandleError("LogWarning")
mbLogBroken = True
End Sub

Private Sub LogUnmatched(ByVal sWhere As String, ByVal sRaw As String)
On Error GoTo error:

Call LogWarning(sWhere & ": unmatched %s -> {?} : " & sRaw)

Exit Sub
error:
Call HandleError("LogUnmatched")
mbFailed = True
End Sub

'---------------------------------------------------------------------------
' Message reading and normalization
'---------------------------------------------------------------------------

Private Function ReadMsgLine(ByVal nMsg As Long, ByVal nLine As Integer, ByVal sSrc As String, ByVal sRole As String) As String
Dim sTemp As String, sKey As String
On Error GoTo error:

ReadMsgLine = ""
If nMsg <= 1 Then Exit Function

sTemp = GetMessages(nMsg, nLine)
If sTemp = "MSG NOT IN DATABASE!" Then
    sKey = sSrc & "|" & nMsg
    If Not mdictLogged.Exists(sKey) Then
        mdictLogged.Add sKey, True
        Call LogWarning(sSrc & " " & sRole & ": MSG NOT IN DATABASE! (#" & nMsg & ")")
    End If
    Exit Function
End If
If sTemp = "none" Then Exit Function

ReadMsgLine = RTrim(StripControl(sTemp))

Exit Function
error:
Call HandleError("ReadMsgLine")
mbFailed = True
End Function

'Removes ANSI escape sequences (ESC [ ... letter) and any other control character. Also used for the Titles table.
Public Function StripControl(ByVal sLine As String) As String
Dim nPos As Long, nLen As Long, sCh As String, sOut As String, bInEsc As Boolean
On Error GoTo error:

nLen = Len(sLine)
nPos = 1
Do While nPos <= nLen
    sCh = Mid$(sLine, nPos, 1)
    If bInEsc Then
        If (sCh >= "A" And sCh <= "Z") Or (sCh >= "a" And sCh <= "z") Then bInEsc = False
    ElseIf sCh = Chr$(27) Then
        If nPos < nLen Then
            If Mid$(sLine, nPos + 1, 1) = "[" Then
                bInEsc = True
                nPos = nPos + 1
            End If
        End If
    ElseIf Asc(sCh) >= 32 And Asc(sCh) <> 127 Then
        sOut = sOut & sCh
    End If
    nPos = nPos + 1
Loop

StripControl = sOut

Exit Function
error:
Call HandleError("StripControl")
mbFailed = True
End Function

Private Function NormalizeLine(ByVal sLine As String, sTokens() As String, ByVal nTokens As Integer, ByRef bFlagged As Boolean) As String
Dim nPos As Long, nLen As Long, nUsed As Integer
Dim sOut As String, sCh As String, sNext As String
On Error GoTo error:

nLen = Len(sLine)
nPos = 1
Do While nPos <= nLen
    sCh = Mid$(sLine, nPos, 1)
    If sCh = "%" And nPos < nLen Then
        sNext = Mid$(sLine, nPos + 1, 1)
        If sNext = "d" Then
            sOut = sOut & "{dmg}"
            If nUsed < nTokens Then
                If sTokens(nUsed) = "{dmg}" Then nUsed = nUsed + 1
            End If
        ElseIf sNext = "s" Then
            If nUsed < nTokens Then
                sOut = sOut & sTokens(nUsed)
                nUsed = nUsed + 1
            Else
                sOut = sOut & "{?}"
                bFlagged = True
            End If
        Else
            sOut = sOut & sCh & sNext
        End If
        nPos = nPos + 2
    Else
        sOut = sOut & sCh
        nPos = nPos + 1
    End If
Loop

NormalizeLine = sOut

Exit Function
error:
Call HandleError("NormalizeLine")
mbFailed = True
End Function

'sList = "|"-separated %s substitutions in order ("" for none); an empty element substitutes nothing.
'The result is RTrimmed (an empty trailing substitution must not leave a space).
Private Function NormalizeWith(ByVal sLine As String, ByVal sList As String, ByRef bFlagged As Boolean) As String
Dim sTokens() As String, nTokens As Integer
On Error GoTo error:

If sList = "" Then
    ReDim sTokens(0)
    nTokens = 0
Else
    sTokens = Split(sList, "|")
    nTokens = UBound(sTokens) + 1
End If

NormalizeWith = RTrim(NormalizeLine(sLine, sTokens, nTokens, bFlagged))

Exit Function
error:
Call HandleError("NormalizeWith")
mbFailed = True
End Function

'---------------------------------------------------------------------------
' Row store (dedupe + truncation) and finalize
'---------------------------------------------------------------------------

'returns 1 when the line is exportable (stored now or already present), 0 when blank
Private Function AddRow(ByVal sName As String, ByVal sKind As String, ByVal sText As String, ByVal sEndText As String, _
    ByVal sEffects As String) As Long
Dim sKey As String, bTrunc As Boolean, i As Long, nCap As Long
On Error GoTo error:

AddRow = 0
If sText = "" And sEndText = "" Then Exit Function
AddRow = 1

If Len(sText) > 100 Then sText = Left$(sText, 100): bTrunc = True
If Len(sEndText) > 100 Then sEndText = Left$(sEndText, 100): bTrunc = True
If Len(sEffects) > 255 Then sEffects = Left$(sEffects, 255): bTrunc = True

'Cast rows and named Self rows belong to one player spell or casting item: every name gets its own row, and
'rows collapse only when name, Msg and EndMsg are all the same (effects merged). Nameless lines (monster-only
'sources, hostile-spell status text, hits, traps) are realm-wide and collapse by text alone.
If sKind = "Cast" Or (sKind = "Self" And Not sName = "") Then
    sKey = sKind & "@" & sName & "|" & sText & "|" & sEndText
Else
    sKey = sKind & "|" & sText & "|" & sEndText
End If

If mdictRows.Exists(sKey) Then
    i = mdictRows.Item(sKey)
    mRowEff(i) = MergeEffects(mRowEff(i), sEffects)
    If Not mRowName(i) = "" And Not sKind = "Cast" And Not sKind = "Self" Then
        'same death text on differently named monsters: the line is realm-wide, so the row loses its name
        If StrComp(mRowName(i), sName, vbTextCompare) <> 0 Then mRowName(i) = ""
    End If
    Exit Function
End If

If bTrunc Then Call LogWarning("Row truncated: Kind=" & sKind & " Name=" & sName & " Text=" & sText)

If mRowCount > UBound(mRowKind) Then
    nCap = UBound(mRowKind) + 256
    ReDim Preserve mRowKind(nCap)
    ReDim Preserve mRowName(nCap)
    ReDim Preserve mRowMsg(nCap)
    ReDim Preserve mRowEnd(nCap)
    ReDim Preserve mRowEff(nCap)
    ReDim Preserve mRowDrop(nCap)
End If

i = mRowCount
mRowKind(i) = sKind
mRowName(i) = sName
mRowMsg(i) = sText
mRowEnd(i) = sEndText
mRowEff(i) = sEffects
mRowDrop(i) = False
mdictRows.Add sKey, i
mRowCount = mRowCount + 1

Exit Function
error:
Call HandleError("AddRow")
mbFailed = True
End Function

'union of two comma lists: order of first appearance, no blanks, no duplicates
Private Function MergeEffects(ByVal sA As String, ByVal sB As String) As String
Dim sParts() As String, sOut As String, sID As String, x As Long
On Error GoTo error:

sOut = ""
sParts = Split(sA & "," & sB, ",")
For x = 0 To UBound(sParts)
    sID = Trim$(sParts(x))
    If Not sID = "" Then
        If InStr(1, "," & sOut & ",", "," & sID & ",", vbBinaryCompare) = 0 Then
            If Not sOut = "" Then sOut = sOut & ","
            sOut = sOut & sID
        End If
    End If
Next x

MergeEffects = sOut

Exit Function
error:
Call HandleError("MergeEffects")
mbFailed = True
End Function

'stores the same row once per name; returns 1 when the line is exportable
Private Function AddRowEach(sNames() As String, ByVal nNames As Long, ByVal sKind As String, ByVal sText As String, _
    ByVal sEndText As String, ByVal sEffects As String) As Long
Dim x As Long
On Error GoTo error:

AddRowEach = 0
For x = 0 To nNames - 1
    If AddRow(sNames(x), sKind, sText, sEndText, sEffects) = 1 Then AddRowEach = 1
    If mbFailed Then Exit For
Next x

Exit Function
error:
Call HandleError("AddRowEach")
mbFailed = True
End Function

'Drops each nameless Self row that is a "you"-twin of a {target} row (a client reads {target} as "you"/"your" for the
'player), merging its effects into that row, then writes every surviving row once with IDs 1, 2, 3...
'A row without a name is named by its text.
'Name for a row that has no source name: the line itself, minus the parts that carry no identity
'(leading "The {source}['s] " / "{source}['s] ", trailing " for {dmg} damage!" / " {dmg} damage!" / " {dmg} hit points!",
'and any "{target}'s " / "{target} "). Falls back to the untouched line if nothing would be left.
Private Function CleanName(ByVal sLine As String) As String
Dim s As String
On Error GoTo error:

s = sLine
If Left$(s, 15) = "The {source}'s " Then
    s = Mid$(s, 16)
ElseIf Left$(s, 13) = "The {source} " Then
    s = Mid$(s, 14)
ElseIf Left$(s, 11) = "{source}'s " Then
    s = Mid$(s, 12)
ElseIf Left$(s, 9) = "{source} " Then
    s = Mid$(s, 10)
End If
If Right$(s, 18) = " for {dmg} damage!" Then
    s = Left$(s, Len(s) - 18)
ElseIf Right$(s, 14) = " {dmg} damage!" Then
    s = Left$(s, Len(s) - 14)
ElseIf Right$(s, 18) = " {dmg} hit points!" Then
    s = Left$(s, Len(s) - 18)
End If
s = Replace(s, "{target}'s ", "")
s = Replace(s, "{target} ", "")
s = Trim$(s)
If s = "" Then s = sLine
CleanName = s

Exit Function
error:
Call HandleError("CleanName")
mbFailed = True
CleanName = sLine
End Function

Private Sub FinalizeRows(rs As DAO.Recordset)
Dim dictYou As Scripting.Dictionary, sVariant As String, sName As String
Dim i As Long, j As Long, nID As Long
On Error GoTo error:

frmMME_Export.lblPanel(1).Caption = "Messages: writing ..."

Set dictYou = New Scripting.Dictionary
dictYou.CompareMode = vbBinaryCompare

For i = 0 To mRowCount - 1
    If Not mRowKind(i) = "Self" And mRowEnd(i) = "" Then
        If InStr(1, mRowMsg(i), "{target}", vbBinaryCompare) > 0 Then
            sVariant = Replace(Replace(mRowMsg(i), "{target}'s", "your"), "{target}", "you")
            If Not dictYou.Exists(sVariant) Then dictYou.Add sVariant, i
        End If
    End If
Next i

For i = 0 To mRowCount - 1
    If mRowKind(i) = "Self" And mRowEnd(i) = "" And mRowName(i) = "" Then
        If dictYou.Exists(mRowMsg(i)) Then
            j = dictYou.Item(mRowMsg(i))
            mRowEff(j) = MergeEffects(mRowEff(j), mRowEff(i))
            mRowDrop(i) = True
        End If
    End If
Next i

nID = 1
For i = 0 To mRowCount - 1
    If Not mRowDrop(i) Then
        sName = mRowName(i)
        If sName = "" Then sName = CleanName(mRowMsg(i))
        If sName = "" Then sName = CleanName(mRowEnd(i))

        rs.AddNew
        rs.Fields("ID") = nID
        rs.Fields("Name") = IIf(sName = "", Chr(0), Left$(sName, 100))
        rs.Fields("Kind") = Left$(mRowKind(i), 100)
        rs.Fields("Msg") = IIf(mRowMsg(i) = "", Chr(0), Left$(mRowMsg(i), 100))
        rs.Fields("EndMsg") = IIf(mRowEnd(i) = "", Chr(0), Left$(mRowEnd(i), 100))
        rs.Fields("Effects") = IIf(mRowEff(i) = "", Chr(0), Left$(mRowEff(i), 255))
        rs.Update

        If nID Mod 100 = 0 Then frmMME_Export.lblPanel(1).Caption = "Messages: writing (" & nID & ")"
        nID = nID + 1
        If Not bUseCPU Then DoEvents
    End If
Next i

Set dictYou = Nothing

Exit Sub
error:
Call HandleError("FinalizeRows")
Set dictYou = Nothing
mbFailed = True
End Sub

'---------------------------------------------------------------------------
' Pass 0: items (spells learnable from items, spells cast by items)
'---------------------------------------------------------------------------

Private Sub ExportItemScan(dictLearn As Scripting.Dictionary, dictItemCast As Scripting.Dictionary, ByRef bStop As Boolean)
Dim nStatus As Integer, nCount As Long, x As Integer
Dim sItem As String, sKey As String, sCur As String
On Error GoTo error:

frmMME_Export.lblPanel(0).Caption = "w" & strDatCallLetters & "item2.dat"
frmMME_Export.lblPanel(1).Caption = "Messages: items ..."

nStatus = BTRCALL(BGETFIRST, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
If nStatus <> 0 And nStatus <> 9 Then Call LogWarning("Items: could not get first record: " & BtrieveErrorCode(nStatus, True))

Do While nStatus = 0 And bStop = False And mbFailed = False
    ItemRowToStruct Itemdatabuf.buf

    nCount = nCount + 1
    If nCount Mod 25 = 0 Then frmMME_Export.lblPanel(1).Caption = "Messages: items (" & nCount & ")"

    sItem = RTrim(ClipNull(Itemrec.Name))
    For x = 0 To 19
        Select Case Itemrec.AbilityA(x)
            Case 42: 'learn spell
                If Itemrec.AbilityB(x) > 0 Then
                    sKey = CStr(Itemrec.AbilityB(x))
                    If Not dictLearn.Exists(sKey) Then dictLearn.Add sKey, True
                End If
            Case 43: 'casts spell
                If Itemrec.AbilityB(x) > 0 And Not sItem = "" Then
                    sKey = CStr(Itemrec.AbilityB(x))
                    If dictItemCast.Exists(sKey) Then
                        sCur = dictItemCast.Item(sKey)
                        If InStr(1, "|" & sCur & "|", "|" & sItem & "|", vbBinaryCompare) = 0 Then
                            dictItemCast.Item(sKey) = sCur & "|" & sItem
                        End If
                    Else
                        dictItemCast.Add sKey, sItem
                    End If
                End If
            Case 137: 'damage shield message (ShockMsg)
                If Itemrec.AbilityB(x) > 1 And Not sItem = "" Then
                    sKey = CStr(Itemrec.AbilityB(x))
                    If mdictItemShield.Exists(sKey) Then
                        sCur = mdictItemShield.Item(sKey)
                        If InStr(1, "|" & sCur & "|", "|" & sItem & "|", vbBinaryCompare) = 0 Then
                            mdictItemShield.Item(sKey) = sCur & "|" & sItem
                        End If
                    Else
                        mdictItemShield.Add sKey, sItem
                    End If
                End If
        End Select
    Next x

    nStatus = BTRCALL(BGETNEXT, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

Exit Sub
error:
Call HandleError("ExportItemScan")
mbFailed = True
End Sub

'---------------------------------------------------------------------------
' Pass 1: rooms (trap rows + room spell list)
'---------------------------------------------------------------------------

Private Sub ExportRoomMessages(dictRoomSpells As Scripting.Dictionary, ByRef bStop As Boolean)
Dim nStatus As Integer, nCount As Long
On Error GoTo error:

frmMME_Export.lblPanel(0).Caption = "w" & strDatCallLetters & "mp002.dat"
frmMME_Export.lblPanel(1).Caption = "Messages: rooms ..."

nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
If nStatus <> 0 And nStatus <> 9 Then Call LogWarning("Rooms: could not get first record: " & BtrieveErrorCode(nStatus, True))

Do While nStatus = 0 And bStop = False And mbFailed = False
    RoomRowToStruct Roomdatabuf.buf

    nCount = nCount + 1
    If nCount Mod 25 = 0 Then frmMME_Export.lblPanel(1).Caption = "Messages: rooms (" & nCount & ")"

    If Not UCase(Left(Roomrec.Name, 11)) = "BUFFER ROOM" Then
        Call ProcessRoom(dictRoomSpells)
    End If

    nStatus = BTRCALL(BGETNEXT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

Exit Sub
error:
Call HandleError("ExportRoomMessages")
mbFailed = True
End Sub

Private Sub ProcessRoom(dictRoomSpells As Scripting.Dictionary)
Dim x As Integer, sSrc As String, sRole As String
On Error GoTo error:

If Roomrec.Spell > 0 Then
    If Not dictRoomSpells.Exists(CStr(Roomrec.Spell)) Then dictRoomSpells.Add CStr(Roomrec.Spell), True
End If

sSrc = "Room " & Roomrec.MapNumber & "/" & Roomrec.RoomNumber
For x = 0 To 9
    sRole = "exit " & x
    If Not Roomrec.RoomExit(x) = 0 Then
        Select Case Roomrec.RoomType(x)
            Case 9:
                Call TrapRow(Roomrec.Para3(x), sSrc, sRole)
                Call TrapRow(Roomrec.Para4(x), sSrc, sRole)
            Case 24:
                Call TrapRow(Roomrec.Para4(x), sSrc, sRole)
                Call TrapRow(Roomrec.Para3(x), sSrc, sRole)
                Call AddSelfCast(Roomrec.Para1(x), "Trap")    'spell trap: cast on the mover
            Case 22:
                Call AddSelfCast(Roomrec.Para1(x), "")        'cast exit: spells before / after the move
                Call AddSelfCast(Roomrec.Para2(x), "")
        End Select
    End If
Next x

Exit Sub
error:
Call HandleError("ProcessRoom")
mbFailed = True
End Sub

'records a spell the player ends up casting on themselves; "Trap" wins over ""
Private Sub AddSelfCast(ByVal nSpell As Long, ByVal sTag As String)
Dim sKey As String
On Error GoTo error:

If nSpell <= 0 Then Exit Sub
sKey = CStr(nSpell)
If mdictSelfCast.Exists(sKey) Then
    If sTag = "Trap" Then mdictSelfCast.Item(sKey) = "Trap"
Else
    mdictSelfCast.Add sKey, sTag
End If

Exit Sub
error:
Call HandleError("AddSelfCast")
mbFailed = True
End Sub

'trap trigger / disarm-fail line 1: Self, nameless, Damage,Trap
Private Sub TrapRow(ByVal nMsg As Long, ByVal sSrc As String, ByVal sRole As String)
Dim sRaw As String, sText As String, bFlag As Boolean
On Error GoTo error:

If nMsg <= 1 Then Exit Sub
sRaw = ReadMsgLine(nMsg, 1, sSrc, sRole)
If sRaw = "" Then Exit Sub

bFlag = False
sText = NormalizeWith(sRaw, "", bFlag)
If bFlag Then Call LogUnmatched(sSrc & " " & sRole & " line 1", sRaw)
Call AddRow("", "Self", sText, "", "Damage,Trap")

Exit Sub
error:
Call HandleError("TrapRow")
mbFailed = True
End Sub

'---------------------------------------------------------------------------
' Pass 2: spells
'---------------------------------------------------------------------------

Private Sub ExportSpellMessages(dictRoomSpells As Scripting.Dictionary, dictLearn As Scripting.Dictionary, dictItemCast As Scripting.Dictionary, ByRef bStop As Boolean)
Dim nStatus As Integer, nCount As Long
On Error GoTo error:

frmMME_Export.lblPanel(0).Caption = "w" & strDatCallLetters & "spel2.dat"
frmMME_Export.lblPanel(1).Caption = "Messages: spells ..."

nStatus = BTRCALL(BGETFIRST, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
If nStatus <> 0 And nStatus <> 9 Then Call LogWarning("Spells: could not get first record: " & BtrieveErrorCode(nStatus, True))

Do While nStatus = 0 And bStop = False And mbFailed = False
    SpellRowToStruct Spelldatabuf.buf

    nCount = nCount + 1
    If nCount Mod 25 = 0 Then frmMME_Export.lblPanel(1).Caption = "Messages: spells (" & nCount & ")"

    Call ProcessSpell(dictRoomSpells, dictLearn, dictItemCast)

    nStatus = BTRCALL(BGETNEXT, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

Exit Sub
error:
Call HandleError("ExportSpellMessages")
mbFailed = True
End Sub

Private Sub ProcessSpell(dictRoomSpells As Scripting.Dictionary, dictLearn As Scripting.Dictionary, dictItemCast As Scripting.Dictionary)
Dim sName As String, sSrc As String, sEffects As String, sKey As String
Dim sItemNames As String, sTargetTok As String, sNames() As String, nNames As Long
Dim sSelfNames() As String, nSelfNames As Long, bFriendly As Boolean
Dim nDur As Integer, bShort As Boolean, bPlayer As Boolean, bArea As Boolean, nRows As Long, nCast As Long, bStdCast As Boolean
Dim nDesc As Long, nConf As Long, nShock As Long
Dim sDesc3 As String, sConf1 As String, bMergeFail As Boolean
On Error GoTo error:

sName = RTrim(ClipNull(Spellrec.Name))
sKey = CStr(Spellrec.Number)
nDur = Spellrec.duration
If mbGmud Then
    bShort = (Spellrec.MsgStyle = 1)
Else
    bShort = ((Spellrec.MsgStyle And 1) = 1)
End If
sSrc = "Spell " & Spellrec.Number & " (" & sName & ")"

'player-castable: learnable somewhere in the realm (books, trainers, textblocks, items). Without cross-reference
'data a magery type is the best guess, but realms keep monster-only spells with a magery type too.
If mbHaveLearnable Then
    bPlayer = dictLearn.Exists(sKey)
Else
    bPlayer = (Spellrec.MageryA > 0) Or dictLearn.Exists(sKey)
End If
sItemNames = ""
If dictItemCast.Exists(sKey) Then sItemNames = dictItemCast.Item(sKey)

Select Case Spellrec.Target
    Case 3, 9, 10, 11, 12, 13:
        bArea = True
    Case Else:
        bArea = False
End Select
If bArea Then
    sTargetTok = ""
Else
    sTargetTok = "{target}"
End If

nNames = BuildNameList(sName, bPlayer, sItemNames, sNames)

'Lines addressed to the affected player are named only when the player can be the one affected by their own
'spell or item (self/friendly target types); hostile-spell status text is seen only by the enemy, so it is
'a realm-wide line and gets no name (one row per distinct text).
Select Case Spellrec.Target
    Case 0, 1, 2, 6, 13:
        bFriendly = True
    Case Else:
        bFriendly = False
End Select
If bFriendly And (bPlayer Or Not sItemNames = "") Then
    sSelfNames = sNames
    nSelfNames = nNames
Else
    ReDim sSelfNames(0)
    sSelfNames(0) = ""
    nSelfNames = 1
End If

nDesc = SpellAbilityValue(115)
nConf = SpellAbilityValue(101)
nShock = SpellAbilityValue(137)
sEffects = SpellEffects(nDur, dictRoomSpells.Exists(sKey))

'a ConfuseMsg identical to the status text merges into the status row as Fail
sDesc3 = ""
If nDesc > 1 Then sDesc3 = ReadMsgLine(nDesc, 3, sSrc, "desc")
sConf1 = ""
If nConf > 1 Then sConf1 = ReadMsgLine(nConf, 1, sSrc, "confuse")
bMergeFail = False
If Not Trim(sConf1) = "" Then
    If Trim(sConf1) = Trim(sDesc3) Then bMergeFail = True
End If

nRows = 0
nRows = nRows + SpellStatusRow(sSelfNames, nSelfNames, sName, sSrc, nDesc, sDesc3, sEffects, bMergeFail)
If Not bMergeFail Then nRows = nRows + SpellFailRow(sSrc, sConf1)
If bPlayer Or Not sItemNames = "" Then
    nCast = SpellCastRow(sNames, nNames, sName, sSrc, bShort, sTargetTok, sEffects, bStdCast)
    'cast line empty or in a standard form: the start line confirms the cast too
    If bStdCast Then nCast = nCast + SpellStartRow(sNames, nNames, sSrc, sTargetTok, sEffects)
    nRows = nRows + nCast
ElseIf mdictSelfCast.Exists(sKey) Then
    nRows = nRows + SpellSelfCastRow(sName, sSrc, bShort, sTargetTok, sEffects, mdictSelfCast.Item(sKey))
End If
nRows = nRows + SpellIncomingRow(sSelfNames, nSelfNames, sName, sSrc, nDesc, bShort, sEffects)
nRows = nRows + SpellShieldRows(sSelfNames, nSelfNames, sSrc, nShock)

If sName = "" And nRows > 0 Then Call LogWarning("Spell " & Spellrec.Number & ": blank spell name")

Exit Sub
error:
Call HandleError("ProcessSpell")
mbFailed = True
End Sub

'Name list for a spell's rows: the spell name when players can cast it, then each casting item's name
'not already present (case-insensitive); just the spell name when nothing applies. Returns the count.
Private Function BuildNameList(ByVal sName As String, ByVal bPlayer As Boolean, ByVal sItemNames As String, sNames() As String) As Long
Dim sItems() As String, x As Long, j As Long, nNames As Long, bDup As Boolean
On Error GoTo error:

ReDim sNames(0)
sNames(0) = ""
nNames = 0

If bPlayer And Not sName = "" Then
    sNames(0) = sName
    nNames = 1
End If

If Not sItemNames = "" Then
    sItems = Split(sItemNames, "|")
    For x = 0 To UBound(sItems)
        If Not sItems(x) = "" Then
            bDup = False
            For j = 0 To nNames - 1
                If StrComp(sNames(j), sItems(x), vbTextCompare) = 0 Then
                    bDup = True
                    Exit For
                End If
            Next j
            If Not bDup Then
                ReDim Preserve sNames(nNames)
                sNames(nNames) = sItems(x)
                nNames = nNames + 1
            End If
        End If
    Next x
End If

If nNames = 0 Then
    sNames(0) = sName
    nNames = 1
End If

BuildNameList = nNames

Exit Function
error:
Call HandleError("BuildNameList")
mbFailed = True
BuildNameList = 0
End Function

'DescMsg (115): Msg = line 3, EndMsg = line 1; Kind Self
Private Function SpellStatusRow(sNames() As String, _
    ByVal nNames As Long, ByVal sName As String, ByVal sSrc As String, ByVal nDesc As Long, ByVal sRaw3 As String, _
    ByVal sEffects As String, ByVal bMergeFail As Boolean) As Long
Dim sRaw1 As String, sText As String, sEnd As String, bFlag As Boolean
On Error GoTo error:

SpellStatusRow = 0
If nDesc <= 1 Then Exit Function

sRaw1 = ReadMsgLine(nDesc, 1, sSrc, "desc")
If sRaw3 = "" And sRaw1 = "" Then Exit Function

bFlag = False
sText = NormalizeWith(sRaw3, "", bFlag)
If bFlag Then Call LogUnmatched(sSrc & " desc line 3", sRaw3)

bFlag = False
sEnd = NormalizeWith(sRaw1, sName, bFlag)
If bFlag Then Call LogUnmatched(sSrc & " desc line 1", sRaw1)

If bMergeFail Then Call AddEffect(sEffects, "Fail")
SpellStatusRow = AddRowEach(sNames, nNames, "Self", sText, sEnd, sEffects)

Exit Function
error:
Call HandleError("SpellStatusRow")
mbFailed = True
End Function

'ConfuseMsg (101) line 1, not merged into the status row: Self, nameless, Fail only
Private Function SpellFailRow(ByVal sSrc As String, ByVal sRaw As String) As Long
Dim sText As String, bFlag As Boolean
On Error GoTo error:

SpellFailRow = 0
If Trim(sRaw) = "" Then Exit Function

bFlag = False
sText = NormalizeWith(sRaw, "", bFlag)
If bFlag Then Call LogUnmatched(sSrc & " confuse line 1", sRaw)
SpellFailRow = AddRow("", "Self", sText, "", "Fail")

Exit Function
error:
Call HandleError("SpellFailRow")
mbFailed = True
End Function

'CastMsgB line 1 (what the caster sees); only for player-castable or item-cast spells. Every such spell and item
'gets its cast line under its own name, standard form or not, so a client can tie the line to its spell/item.
'bStandard returns True when the line is empty or one of the engine's standard forms (the start line is then
'exported as well, see SpellStartRow).
Private Function SpellCastRow(sNames() As String, _
    ByVal nNames As Long, ByVal sName As String, ByVal sSrc As String, ByVal bShort As Boolean, ByVal sTargetTok As String, _
    ByVal sEffects As String, ByRef bStandard As Boolean) As Long
Dim sRaw As String, sText As String, sList As String, bFlag As Boolean
On Error GoTo error:

SpellCastRow = 0
bStandard = True
If Spellrec.CastMsgB <= 1 Then Exit Function

sRaw = ReadMsgLine(Spellrec.CastMsgB, 1, sSrc, "cast")
If sRaw = "" Then Exit Function

If bShort Then
    sList = sTargetTok & "|{dmg}"
Else
    sList = sName & "|" & sTargetTok & "|{dmg}"
End If
bFlag = False
sText = NormalizeWith(sRaw, sList, bFlag)
If sText = "" Then Exit Function

If InStr(1, sText, "{dmg}", vbBinaryCompare) = 0 Then
    bStandard = IsStandardForm(mReP7, sText)
Else
    bStandard = IsStandardForm(mReP1, sText) Or IsStandardForm(mReP3, sText) Or IsStandardForm(mReP4, sText)
End If

If bFlag Then Call LogUnmatched(sSrc & " cast line 1", sRaw)
SpellCastRow = AddRowEach(sNames, nNames, "Cast", sText, "", sEffects)

Exit Function
error:
Call HandleError("SpellCastRow")
mbFailed = True
End Function

'StartMsg (120) line 1: shown to the caster once the spell has landed (Paradigm Spell.cs), every %s = the target.
'Used only when the cast line itself is empty or standard; Kind Cast.
Private Function SpellStartRow(sNames() As String, ByVal nNames As Long, ByVal sSrc As String, _
    ByVal sTargetTok As String, ByVal sEffects As String) As Long
Dim nStart As Long, sRaw As String, sText As String, bFlag As Boolean
On Error GoTo error:

SpellStartRow = 0
nStart = SpellAbilityValue(120)
If nStart <= 1 Then Exit Function

sRaw = ReadMsgLine(nStart, 1, sSrc, "start")
If sRaw = "" Then Exit Function

bFlag = False
sText = NormalizeWith(sRaw, sTargetTok & "|" & sTargetTok, bFlag)
If sText = "" Then Exit Function
If IsStandardForm(mReP7, sText) Then Exit Function

If bFlag Then Call LogUnmatched(sSrc & " start line 1", sRaw)
SpellStartRow = AddRowEach(sNames, nNames, "Cast", sText, "", sEffects)

Exit Function
error:
Call HandleError("SpellStartRow")
mbFailed = True
End Function

'CastMsgB line 1 of a spell the player casts without choosing it (textblock, cast exit, spell-trap exit): the
'player is the caster, so this is the line they see. Self, nameless, only when the spell does something a
'client tracks (damage, healing, an ailment...). Spell traps add Trap. Monster/item target types never reach
'the player (Paradigm TextBlockPart.cs "cast").
Private Function SpellSelfCastRow(ByVal sName As String, ByVal sSrc As String, ByVal bShort As Boolean, _
    ByVal sTargetTok As String, ByVal sEffects As String, ByVal sTag As String) As Long
Dim sRaw As String, sText As String, sList As String, sEff As String, bFlag As Boolean
On Error GoTo error:

SpellSelfCastRow = 0
If Spellrec.CastMsgB <= 1 Then Exit Function
If Spellrec.Target = 4 Or Spellrec.Target = 7 Then Exit Function

sEff = Replace(Replace(sEffects, ",RoomSpell", ""), "RoomSpell", "")
If sEff = "" Then Exit Function
If sTag = "Trap" Then Call AddEffect(sEff, "Trap")

sRaw = ReadMsgLine(Spellrec.CastMsgB, 1, sSrc, "cast")
If sRaw = "" Then Exit Function

If bShort Then
    sList = sTargetTok & "|{dmg}"
Else
    sList = sName & "|" & sTargetTok & "|{dmg}"
End If
bFlag = False
sText = NormalizeWith(sRaw, sList, bFlag)
If sText = "" Then Exit Function

If IsStandardForm(mReP7, sText) Then Exit Function
If IsStandardForm(mReP1, sText) Then Exit Function
If IsStandardForm(mReP3, sText) Then Exit Function
If IsStandardForm(mReP4, sText) Then Exit Function

If bFlag Then Call LogUnmatched(sSrc & " cast line 1 (self-cast)", sRaw)
SpellSelfCastRow = AddRow("", "Self", sText, "", sEff)

Exit Function
error:
Call HandleError("SpellSelfCastRow")
mbFailed = True
End Function

'CastMsgB line 2 (what the target sees): Self
Private Function SpellIncomingRow(sNames() As String, _
    ByVal nNames As Long, ByVal sName As String, ByVal sSrc As String, ByVal nDesc As Long, ByVal bShort As Boolean, _
    ByVal sEffects As String) As Long
Dim sRaw As String, sText As String, sList As String, bFlag As Boolean, bExport As Boolean
On Error GoTo error:

SpellIncomingRow = 0
If Spellrec.CastMsgB <= 1 Then Exit Function

sRaw = ReadMsgLine(Spellrec.CastMsgB, 2, sSrc, "cast")
If sRaw = "" Then Exit Function

'GreaterMUD (Paradigm) always passes caster, spell, value to line 2
If mbGmud Or Not bShort Then
    sList = "{source}|" & sName & "|{dmg}"
Else
    sList = "{dmg}"
End If
bFlag = False
sText = NormalizeWith(sRaw, sList, bFlag)

bExport = False
If InStr(1, sText, "{dmg}", vbBinaryCompare) > 0 Then
    bExport = True
    If IsStandardForm(mReP2, sText) Then bExport = False
    If IsStandardForm(mReP3, sText) Then bExport = False
    If IsStandardForm(mReP4, sText) Then bExport = False
End If
If bExport = False And nDesc <= 1 Then
    If SpellHasStatusAbility() Then bExport = True
End If
If bExport = False And nDesc > 1 And InStr(1, sText, "{dmg}", vbBinaryCompare) = 0 Then
    'a status text (printed right after this line) plus a wear-off text already show the effect starting and
    'ending; when the status message lacks either one, this line is the client's only sign of the effect
    If ReadMsgLine(nDesc, 3, sSrc, "desc") = "" Or ReadMsgLine(nDesc, 1, sSrc, "desc") = "" Then bExport = True
End If
If bExport = False Then Exit Function

If bFlag Then Call LogUnmatched(sSrc & " cast line 2", sRaw)
SpellIncomingRow = AddRowEach(sNames, nNames, "Self", sText, "", sEffects)

Exit Function
error:
Call HandleError("SpellIncomingRow")
mbFailed = True
End Function

'Item damage shields (item ability 137): line 1 of the ShockMsg (the wearer hits back), one row per item name.
'Line 2 is what the wearer's attacker sees, never the wearer, so it is left out.
Private Sub ExportItemShieldRows()
Dim vKey As Variant, sNames() As String, nNames As Long
On Error GoTo error:

frmMME_Export.lblPanel(1).Caption = "Messages: item shields ..."

For Each vKey In mdictItemShield.Keys
    sNames = Split(mdictItemShield.Item(vKey), "|")
    nNames = UBound(sNames) + 1
    Call SpellShieldRows(sNames, nNames, "Item shield message " & vKey & " (" & mdictItemShield.Item(vKey) & ")", CLng(vKey), True)
    If mbFailed Then Exit For
Next vKey

Exit Sub
error:
Call HandleError("ExportItemShieldRows")
mbFailed = True
End Sub

'ShockMsg (137 on a spell or an item): line 1 (shield owner) and line 2 (attacker hit by the shield);
'Self, Damage,DamageShield
Private Function SpellShieldRows(sNames() As String, _
    ByVal nNames As Long, ByVal sSrc As String, ByVal nShock As Long, Optional ByVal bOwnerOnly As Boolean = False) As Long
Dim sRaw As String, sText As String, bFlag As Boolean, nRows As Long
On Error GoTo error:

SpellShieldRows = 0
If nShock <= 1 Then Exit Function

sRaw = ReadMsgLine(nShock, 1, sSrc, "shock")
If Not sRaw = "" Then
    bFlag = False
    sText = NormalizeWith(sRaw, "{attacker}|{dmg}", bFlag)
    If Not (IsStandardForm(mReP3, sText) Or IsStandardForm(mReP4, sText)) Then
        If bFlag Then Call LogUnmatched(sSrc & " shock line 1", sRaw)
        nRows = nRows + AddRowEach(sNames, nNames, "Self", sText, "", "Damage,DamageShield")
    End If
End If

If bOwnerOnly Then
    SpellShieldRows = nRows
    Exit Function
End If

sRaw = ReadMsgLine(nShock, 2, sSrc, "shock")
If Not sRaw = "" Then
    bFlag = False
    sText = NormalizeWith(sRaw, "{dmg}", bFlag)
    If Not (IsStandardForm(mReP3, sText) Or IsStandardForm(mReP4, sText)) Then
        If bFlag Then Call LogUnmatched(sSrc & " shock line 2", sRaw)
        nRows = nRows + AddRowEach(sNames, nNames, "Self", sText, "", "Damage,DamageShield")
    End If
End If

SpellShieldRows = nRows

Exit Function
error:
Call HandleError("SpellShieldRows")
mbFailed = True
End Function

'---------------------------------------------------------------------------
' Spell ability helpers (read the current Spellrec)
'---------------------------------------------------------------------------

Private Function SpellHasAbil(ByVal nAbil As Integer) As Boolean
Dim x As Integer
On Error GoTo error:

SpellHasAbil = False
For x = 0 To 9
    If Spellrec.AbilityA(x) = nAbil Then
        SpellHasAbil = True
        Exit Function
    End If
Next x

Exit Function
error:
Call HandleError("SpellHasAbil")
mbFailed = True
End Function

'value of the first slot holding nAbil, 0 when absent
Private Function SpellAbilityValue(ByVal nAbil As Integer) As Long
Dim x As Integer
On Error GoTo error:

SpellAbilityValue = 0
For x = 0 To 9
    If Spellrec.AbilityA(x) = nAbil Then
        SpellAbilityValue = Spellrec.AbilityB(x)
        Exit Function
    End If
Next x

Exit Function
error:
Call HandleError("SpellAbilityValue")
mbFailed = True
End Function

Private Function SpellHasStatusAbility() As Boolean
On Error GoTo error:

SpellHasStatusAbility = False
If SpellHasAbil(19) Or SpellHasAbil(71) Or SpellHasAbil(107) Or SpellHasAbil(74) Then SpellHasStatusAbility = True
If SpellHasAbil(75) Or SpellHasAbil(60) Or SpellHasAbil(76) Or SpellHasAbil(170) Then SpellHasStatusAbility = True

Exit Function
error:
Call HandleError("SpellHasStatusAbility")
mbFailed = True
End Function

Private Sub AddEffect(ByRef sList As String, ByVal sID As String)
On Error GoTo error:

If Not sList = "" Then sList = sList & ","
sList = sList & sID

Exit Sub
error:
Call HandleError("AddEffect")
mbFailed = True
End Sub

'nDur only selects the *OverTime variants
Private Function SpellEffects(ByVal nDur As Integer, bRoomSpell As Boolean) As String
Dim s As String
On Error GoTo error:

s = ""
If SpellHasAbil(19) Then Call AddEffect(s, "Poisoned")
If SpellHasAbil(71) Then Call AddEffect(s, "Confused")
If SpellHasAbil(107) Then Call AddEffect(s, "Blinded")
If SpellHasAbil(74) Or SpellHasAbil(75) Then Call AddEffect(s, "Held")
If SpellHasAbil(60) Then Call AddEffect(s, "Fear")
If SpellHasAbil(76) Then Call AddEffect(s, "Mute")
If SpellHasAbil(170) Then Call AddEffect(s, "Sleep")
If SpellHasAbil(1) Or SpellHasAbil(17) Then
    If nDur > 0 Then
        Call AddEffect(s, "DamageOverTime")
    Else
        Call AddEffect(s, "Damage")
    End If
End If
If SpellHasAbil(8) Or SpellHasAbil(175) Then Call AddEffect(s, "Drain")
If SpellHasAbil(18) Then
    If nDur > 0 Then
        Call AddEffect(s, "HealOverTime")
    Else
        Call AddEffect(s, "Heal")
    End If
End If
If SpellHasAbil(123) Then Call AddEffect(s, "HPRegen")
If SpellHasAbil(145) Then Call AddEffect(s, "ManaRegen")
If SpellHasAbil(150) Then Call AddEffect(s, "ManaHeal")
If SpellHasAbil(174) Or SpellHasAbil(176) Then Call AddEffect(s, "ManaSteal")
If SpellHasAbil(72) Then Call AddEffect(s, "DamageShield")
If SpellHasAbil(20) Or SpellHasAbil(73) Or SpellHasAbil(81) Or SpellHasAbil(84) Then Call AddEffect(s, "Cure")
If bRoomSpell Then Call AddEffect(s, "RoomSpell")

SpellEffects = s

Exit Function
error:
Call HandleError("SpellEffects")
mbFailed = True
End Function

'---------------------------------------------------------------------------
' Static rows: realm-independent lines from the stock MegaMud MESSAGES.md (text as the stock game prints it)
'---------------------------------------------------------------------------

'Chase: what others see when someone leaves through a command exit (a client can follow a party member).
'CheckRoom: a monster appeared or changed form (a client should look at the room again). Kind Other, nameless.
Private Sub ExportStaticRows()
On Error GoTo error:

Call AddRow("", "Other", "The chains snap, and {target} slips inside the warehouse.", "", "Chase")
Call AddRow("", "Other", "{target} carefully begins to climb down.", "", "Chase")
Call AddRow("", "Other", "{target} climbs down the mountain stair.", "", "Chase")
Call AddRow("", "Other", "{target} climbs down through the hole.", "", "Chase")
Call AddRow("", "Other", "{target} climbs down to a lower ledge with the aid of a climbing harness.", "", "Chase")
Call AddRow("", "Other", "{target} climbs into a boat, and rows off across the river.", "", "Chase")
Call AddRow("", "Other", "{target} climbs into a hole.", "", "Chase")
Call AddRow("", "Other", "{target} climbs into one of the skiffs, and rows out into the river.", "", "Chase")
Call AddRow("", "Other", "{target} climbs into the hole in the ceiling!", "", "Chase")
Call AddRow("", "Other", "{target} climbs into the small hole in the hillside.", "", "Chase")
Call AddRow("", "Other", "{target} climbs out of the filthy moat.", "", "Chase")
Call AddRow("", "Other", "{target} climbs out of the small hole in the hillside!", "", "Chase")
Call AddRow("", "Other", "{target} climbs out of the water.", "", "Chase")
Call AddRow("", "Other", "{target} climbs out through the hole to the south.", "", "Chase")
Call AddRow("", "Other", "{target} climbs out through the hole.", "", "Chase")
Call AddRow("", "Other", "{target} climbs the steps encircling the ancient silverwood tree.", "", "Chase")
Call AddRow("", "Other", "{target} climbs through a hole in the wall!", "", "Chase")
Call AddRow("", "Other", "{target} climbs through the hole in the burnt stump.", "", "Chase")
Call AddRow("", "Other", "{target} climbs through the slimy hole.", "", "Chase")
Call AddRow("", "Other", "{target} climbs up the mountain stair.", "", "Chase")
Call AddRow("", "Other", "{target} climbs up through the filthy pipe.", "", "Chase")
Call AddRow("", "Other", "{target} crawls into the largest hole in the avalanche of stone.", "", "Chase")
Call AddRow("", "Other", "{target} descends the steps encircling the ancient silverwood tree.", "", "Chase")
Call AddRow("", "Other", "{target} dives into the black waters of the moat!", "", "Chase")
Call AddRow("", "Other", "{target} drinks a bubbling white potion, and begins to float upwards!", "", "Chase")
Call AddRow("", "Other", "{target} enters a manhole in the ground!", "", "Chase")
Call AddRow("", "Other", "{target} enters the crimson portal and vanishes!", "", "Chase")
Call AddRow("", "Other", "{target} enters the ebony portal and vanishes!", "", "Chase")
Call AddRow("", "Other", "{target} enters the golden portal and vanishes!", "", "Chase")
Call AddRow("", "Other", "{target} follows a hidden path!", "", "Chase")
Call AddRow("", "Other", "{target} gingerly steps inside the cracked willow.", "", "Chase")
Call AddRow("", "Other", "{target} grasps onto a branch of the tree, and climbs upwards.", "", "Chase")
Call AddRow("", "Other", "{target} grasps onto some of the vines, and begins to climb!", "", "Chase")
Call AddRow("", "Other", "{target} hops into one of the boats, and begins to row out across the river.", "", "Chase")
Call AddRow("", "Other", "{target} hops into one of the skiffs, and rows away.", "", "Chase")
Call AddRow("", "Other", "{target} jumps and grabs onto a low branch, and begins to climb the tree.", "", "Chase")
Call AddRow("", "Other", "{target} jumps into the filthy black waters of the moat.", "", "Chase")
Call AddRow("", "Other", "{target} pulls the lever, and falls through a trapdoor!", "", "Chase")
Call AddRow("", "Other", "{target} pulls the lever, and steps through the gate.", "", "Chase")
Call AddRow("", "Other", "{target} pushes through the bushes!", "", "Chase")
Call AddRow("", "Other", "{target} pushes through the bushes, and walks onto a concealed path.", "", "Chase")
Call AddRow("", "Other", "{target} pushes through the tentacles and into the hole.", "", "Chase")
Call AddRow("", "Other", "{target} rows a skiff north, up to the shore.", "", "Chase")
Call AddRow("", "Other", "{target} rows a skiff south, up to the Silvermere docks.", "", "Chase")
Call AddRow("", "Other", "{target} rows a wooden skiff to the east.", "", "Chase")
Call AddRow("", "Other", "{target} rows a wooden skiff to the north.", "", "Chase")
Call AddRow("", "Other", "{target} rows a wooden skiff to the northeast.", "", "Chase")
Call AddRow("", "Other", "{target} rows a wooden skiff to the northwest.", "", "Chase")
Call AddRow("", "Other", "{target} rows a wooden skiff to the south.", "", "Chase")
Call AddRow("", "Other", "{target} rows a wooden skiff to the southeast.", "", "Chase")
Call AddRow("", "Other", "{target} rows a wooden skiff to the southwest.", "", "Chase")
Call AddRow("", "Other", "{target} rows a wooden skiff to the west.", "", "Chase")
Call AddRow("", "Other", "{target} rubs a darkwood ring, and is swept through the arch!", "", "Chase")
Call AddRow("", "Other", "{target} shoves aside the foliage, and disappears among the trees.", "", "Chase")
Call AddRow("", "Other", "{target} slips into the dangerous black alley.", "", "Chase")
Call AddRow("", "Other", "{target} slips into the dark alley.", "", "Chase")
Call AddRow("", "Other", "{target} squeezes into the dark space beneath the manhole.", "", "Chase")
Call AddRow("", "Other", "{target} squeezes into the opening under the bushes.", "", "Chase")
Call AddRow("", "Other", "{target} starts down the slope, but slips and slides down!", "", "Chase")
Call AddRow("", "Other", "{target} steps back, starts running to the east, and makes a mighty leap!", "", "Chase")
Call AddRow("", "Other", "{target} steps back, starts running to the north, and makes a mighty leap!", "", "Chase")
Call AddRow("", "Other", "{target} steps back, starts running to the south, and makes a mighty leap!", "", "Chase")
Call AddRow("", "Other", "{target} steps back, starts running to the west, and makes a mighty leap!", "", "Chase")
Call AddRow("", "Other", "{target} steps into the dark portal, and vanishes!", "", "Chase")
Call AddRow("", "Other", "{target} steps into the portal, and disappears!", "", "Chase")
Call AddRow("", "Other", "{target} steps into the swirling vortex, and vanishes in a flash!", "", "Chase")
Call AddRow("", "Other", "{target} steps into the water at the lake's edge and dives beneath the ice!", "", "Chase")
Call AddRow("", "Other", "{target} steps onto the side-path and quickly vanishes from view.", "", "Chase")
Call AddRow("", "Other", "{target} steps onto the slope and begins to slide down uncontrollably!", "", "Chase")
Call AddRow("", "Other", "{target} steps through the shimmering portal, and vanishes!", "", "Chase")
Call AddRow("", "Other", "{target} steps through the swirling portal.", "", "Chase")
Call AddRow("", "Other", "{target} swings over the edge of the cliff, and climbs down the vines!", "", "Chase")
Call AddRow("", "Other", "{target} touches the black gem, and is consumed in a crackling field of energy!", "", "Chase")
Call AddRow("", "Other", "{target} touches the hideous stone altar.", "", "Chase")
Call AddRow("", "Other", "{target} walks into the wide crack.", "", "Chase")
Call AddRow("", "Other", "{target} walks through the wall to the east!", "", "Chase")
Call AddRow("", "Other", "{target} walks through the wall to the west!", "", "Chase")
Call AddRow("", "Other", "{target} walks through the waterfall!", "", "Chase")
Call AddRow("", "Other", "A barrow wight steps out of the shadows!", "", "CheckRoom")
Call AddRow("", "Other", "A chill runs down your spine. Something is with you!", "", "CheckRoom")
Call AddRow("", "Other", "A coral lobster rises from its rest.", "", "CheckRoom")
Call AddRow("", "Other", "A dark storm approaches!", "", "CheckRoom")
Call AddRow("", "Other", "A deathly shriek announces an evil prescence!", "", "CheckRoom")
Call AddRow("", "Other", "A gibbering horror crawls out of the large gibbering horror's corpse!", "", "CheckRoom")
Call AddRow("", "Other", "A head sprouts from the hydra's shoulders!", "", "CheckRoom")
Call AddRow("", "Other", "A horde of severed heads descend upon you!", "", "CheckRoom")
Call AddRow("", "Other", "A laughter filled with ancient evil sounds behind you!", "", "CheckRoom")
Call AddRow("", "Other", "A mocking female laugh announces an evil prescence!", "", "CheckRoom")
Call AddRow("", "Other", "A moss zombie rises from the ground!", "", "CheckRoom")
Call AddRow("", "Other", "A silvery skull materializes with a metallic shriek!", "", "CheckRoom")
Call AddRow("", "Other", "A skeleton arises from its place of rest.", "", "CheckRoom")
Call AddRow("", "Other", "A small gibbering horror crawls out of the gibbering horror's corpse!", "", "CheckRoom")
Call AddRow("", "Other", "A swirling red mist forms nearby!", "", "CheckRoom")
Call AddRow("", "Other", "A tentacled abomination slithers out of the dark pool.", "", "CheckRoom")
Call AddRow("", "Other", "A whipvine breaks through the ground.", "", "CheckRoom")
Call AddRow("", "Other", "An unholy idol animates!", "", "CheckRoom")
Call AddRow("", "Other", "As the Champion of Blood falls, a tower of fire whirls about his body!", "", "CheckRoom")
Call AddRow("", "Other", "Dark smoke and the smell of burning flesh flood the room!", "", "CheckRoom")
Call AddRow("", "Other", "From the fresh corpse arises a vampire!", "", "CheckRoom")
Call AddRow("", "Other", "Hundreds of insects rise from the ground forming a large swarm.", "", "CheckRoom")
Call AddRow("", "Other", "Out of nothingness a hundred scarab beetles appear to feast on you.", "", "CheckRoom")
Call AddRow("", "Other", "The dark phoenix hatches from its volcanic nest!", "", "CheckRoom")
Call AddRow("", "Other", "The feline form of a grall emerges from the shadows.", "", "CheckRoom")
Call AddRow("", "Other", "The giant shrugs off his glacial prison, and bellows in rage!", "", "CheckRoom")
Call AddRow("", "Other", "The hydra strains against its icy bond, breaks it, and comes to life!", "", "CheckRoom")
Call AddRow("", "Other", "The king crab burrows up through the sand.", "", "CheckRoom")
Call AddRow("", "Other", "The storm giant king stands here in his throne room.", "", "CheckRoom")

Exit Sub
error:
Call HandleError("ExportStaticRows")
mbFailed = True
End Sub

'---------------------------------------------------------------------------
' Pass 3: monsters
'---------------------------------------------------------------------------

Private Sub ExportMonsterMessages(ByRef bStop As Boolean)
Dim nStatus As Integer, nCount As Long
On Error GoTo error:

frmMME_Export.lblPanel(0).Caption = "w" & strDatCallLetters & "knms2.dat"
frmMME_Export.lblPanel(1).Caption = "Messages: monsters ..."

nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
If nStatus <> 0 And nStatus <> 9 Then Call LogWarning("Monsters: could not get first record: " & BtrieveErrorCode(nStatus, True))

Do While nStatus = 0 And bStop = False And mbFailed = False
    MonsterRowToStruct Monsterdatabuf.buf

    nCount = nCount + 1
    If nCount Mod 25 = 0 Then frmMME_Export.lblPanel(1).Caption = "Messages: monsters (" & nCount & ")"

    Call ProcessMonster

    nStatus = BTRCALL(BGETNEXT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

Exit Sub
error:
Call HandleError("ExportMonsterMessages")
mbFailed = True
End Sub

Private Sub ProcessMonster()
Dim sMon As String, sSrc As String, i As Integer
Dim nDeath As Long, sRaw As String, sText As String, bFlag As Boolean, nExp As Double
On Error GoTo error:

sMon = RTrim(ClipNull(Monsterrec.Name))
sSrc = "Monster " & Monsterrec.Number & " (" & sMon & ")"

For i = 0 To 4
    If Monsterrec.AttackType(i) = 1 And Monsterrec.AttackHitMsg(i) > 1 Then
        Call MonsterHitRows(Monsterrec.AttackHitMsg(i), sSrc, "attack " & i)
    End If
Next i

'death text only for monsters that regenerate or are game-limited and give at least 101 experience
'(v1.11j+: the experience given is the base times the multiplier)
nExp = SLong2ULong(Monsterrec.Experience)
If eDatFileVersion >= v111j Then
    If Monsterrec.ExpMulti > 1 Then nExp = nExp * SLong2ULong(Monsterrec.ExpMulti)
End If
If Not ((Monsterrec.RegenTime > 0 Or Monsterrec.GameLimit > 0) And nExp >= 101) Then Exit Sub

nDeath = Monsterrec.DeathMsg
If nDeath <= 1 And Monsterrec.AttackType(0) = 1 Then nDeath = Monsterrec.AttackHitMsg(0)
If nDeath <= 1 Then Exit Sub

sRaw = ReadMsgLine(nDeath, 3, sSrc, "death")
If sRaw = "" Then Exit Sub

bFlag = False
sText = NormalizeWith(sRaw, "{source}", bFlag)
If bFlag Then Call LogUnmatched(sSrc & " death line 3", sRaw)
Call AddRow(sMon, "Other", sText, "", "Death")

Exit Sub
error:
Call HandleError("ProcessMonster")
mbFailed = True
End Sub

'AttackHitMsg line 1 = hit on me (Self); line 2 = hit on someone else (Other), only when line 1 was exported
Private Sub MonsterHitRows(ByVal nMsg As Long, ByVal sSrc As String, ByVal sRole As String)
Dim sRaw As String, sText As String, bFlag As Boolean
On Error GoTo error:

sRaw = ReadMsgLine(nMsg, 1, sSrc, sRole)
If sRaw = "" Then Exit Sub

bFlag = False
sText = NormalizeWith(sRaw, "{source}|{dmg}", bFlag)
If IsStandardForm(mReP5, sText) Then Exit Sub
If IsStandardForm(mReP3, sText) Then Exit Sub
If bFlag Then Call LogUnmatched(sSrc & " " & sRole & " line 1", sRaw)
If AddRow("", "Self", sText, "", "Damage,Monster") = 0 Then Exit Sub

sRaw = ReadMsgLine(nMsg, 2, sSrc, sRole)
If sRaw = "" Then Exit Sub

bFlag = False
sText = NormalizeWith(sRaw, "{source}|{target}|{dmg}", bFlag)
If IsStandardForm(mReP6, sText) Then Exit Sub
If bFlag Then Call LogUnmatched(sSrc & " " & sRole & " line 2", sRaw)
Call AddRow("", "Other", sText, "", "Damage,Monster")

Exit Sub
error:
Call HandleError("MonsterHitRows")
mbFailed = True
End Sub
