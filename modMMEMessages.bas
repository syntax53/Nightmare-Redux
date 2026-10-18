Attribute VB_Name = "modMMEMessages"
Option Base 0
Option Explicit

'MME export: builds the Messages table (client-relevant game message lines).
'Spec: PLAN-mme-messages-export.md, "Revision 2" (R2.1-R2.11; R2.11 overrides).
'Order: items (learnable / item-cast spell lists), rooms (traps + room spell list), spells, monsters.
'Kinds: Cast, Self, Other. Tokens: {source} {target} {dmg} {attacker}.

Private mFso As Scripting.FileSystemObject
Private mLog As Scripting.TextStream
Private mdictLogged As Scripting.Dictionary
Private msLogPath As String
Private mnWarnings As Long
Private mbLogBroken As Boolean
Private mbFailed As Boolean
Private mbGmud As Boolean

'hard-coded (never exported) formats, applied to the substituted text
Private mReP1 As Object    'own standard cast line
Private mReP2 As Object    'standard incoming cast line
Private mReP3 As Object    'standard incoming damage lines
Private mReP4 As Object    'third-party damage observation / shield owner line
Private mReP5 As Object    'standard monster hit on me
Private mReP6 As Object    'standard monster hit on someone else

Public Function MME_ExportMessages(tabMessages As DAO.Recordset, ByVal sLogPath As String, ByRef bStop As Boolean) As Long
Dim dictSeen As Scripting.Dictionary, dictRoomSpells As Scripting.Dictionary
Dim dictLearn As Scripting.Dictionary, dictItemCast As Scripting.Dictionary
Dim nNextID As Long
On Error GoTo error:

MME_ExportMessages = 0
If Not InitState(sLogPath) Then
    Call CleanupState
    MME_ExportMessages = -1
    Exit Function
End If

mbGmud = (frmMME_Export.chkLegit(1).Value = 1)

Set dictSeen = New Scripting.Dictionary
dictSeen.CompareMode = vbBinaryCompare
Set dictRoomSpells = New Scripting.Dictionary
dictRoomSpells.CompareMode = vbBinaryCompare
Set dictLearn = New Scripting.Dictionary
dictLearn.CompareMode = vbBinaryCompare
Set dictItemCast = New Scripting.Dictionary
dictItemCast.CompareMode = vbBinaryCompare

tryagain:
If tabMessages.RecordCount <> 0 Then
    tabMessages.MoveFirst
    tabMessages.Delete
    GoTo tryagain:
End If

nNextID = 1
Call ExportItemScan(dictLearn, dictItemCast, bStop)
If bStop = False And mbFailed = False Then Call ExportRoomMessages(tabMessages, dictSeen, nNextID, dictRoomSpells, bStop)
If bStop = False And mbFailed = False Then Call ExportSpellMessages(tabMessages, dictSeen, nNextID, dictRoomSpells, dictLearn, dictItemCast, bStop)
If bStop = False And mbFailed = False Then Call ExportMonsterMessages(tabMessages, dictSeen, nNextID, bStop)

If mbFailed Then
    MME_ExportMessages = -1
Else
    MME_ExportMessages = mnWarnings
End If

Call CleanupState
Set dictSeen = Nothing
Set dictRoomSpells = Nothing
Set dictLearn = Nothing
Set dictItemCast = Nothing

Exit Function
error:
Call HandleError("MME_ExportMessages")
Call CleanupState
Set dictSeen = Nothing
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

Set mReP1 = NewRegExp("^You cast .+ (on|at|upon) .*\{dmg\} (damage|healing|hit points)!?$")
Set mReP2 = NewRegExp("^\{source\} casts .+ (on|at|upon) you.*\{dmg\} damage!?$")
Set mReP3 = NewRegExp("^You are .* for {dmg} damage!?$|^You take {dmg} damage|you for {dmg} damage")
Set mReP4 = NewRegExp("^\{target\} .*\{dmg\}.* damage!?$|^\{attacker\} .*\{dmg\}.* damage!?$")
Set mReP5 = NewRegExp("^(The |An? )?\{source\} [a-z-]+ you for \{dmg\} damage!$")
Set mReP6 = NewRegExp("^(The |An? )?\{source\} .* \{target\} for \{dmg\} damage!$")

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

'Removes ANSI escape sequences (ESC [ ... letter) and any other control character.
Private Function StripControl(ByVal sLine As String) As String
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
' Row writer (dedupe + truncation)
'---------------------------------------------------------------------------

'returns 1 when the line is exportable (written now or already present), 0 when blank
Private Function AddRow(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, ByVal sName As String, _
    ByVal sKind As String, ByVal sText As String, ByVal sEndText As String, ByVal sEffects As String) As Long
Dim sKey As String, bTrunc As Boolean
On Error GoTo error:

AddRow = 0
If sText = "" And sEndText = "" Then Exit Function
AddRow = 1

If Len(sName) > 50 Then sName = Left$(sName, 50): bTrunc = True
If Len(sText) > 100 Then sText = Left$(sText, 100): bTrunc = True
If Len(sEndText) > 100 Then sEndText = Left$(sEndText, 100): bTrunc = True
If Len(sEffects) > 255 Then sEffects = Left$(sEffects, 255): bTrunc = True

'Cast rows confirm the player's own action, so the name is part of the row; every other row is a line the
'game shows the player, and the same text with the same effects means the same thing whatever produced it.
If sKind = "Cast" Then
    sKey = sKind & "|" & sName & "|" & sText & "|" & sEndText & "|" & sEffects
Else
    sKey = sKind & "|" & sText & "|" & sEndText & "|" & sEffects
End If
If dictSeen.Exists(sKey) Then Exit Function
dictSeen.Add sKey, True

If bTrunc Then Call LogWarning("Row truncated: Kind=" & sKind & " Name=" & sName & " Text=" & sText)

rs.AddNew
rs.Fields("ID") = nNextID
rs.Fields("Name") = IIf(sName = "", Chr(0), sName)
rs.Fields("Kind") = sKind
rs.Fields("Msg") = IIf(sText = "", Chr(0), sText)
rs.Fields("EndMsg") = IIf(sEndText = "", Chr(0), sEndText)
rs.Fields("Effects") = IIf(sEffects = "", Chr(0), sEffects)
rs.Update
nNextID = nNextID + 1

Exit Function
error:
Call HandleError("AddRow")
mbFailed = True
End Function

'writes the same row once per name; returns 1 when the line is exportable
Private Function AddRowEach(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, sNames() As String, _
    ByVal nNames As Long, ByVal sKind As String, ByVal sText As String, ByVal sEndText As String, ByVal sEffects As String) As Long
Dim x As Long
On Error GoTo error:

AddRowEach = 0
For x = 0 To nNames - 1
    If AddRow(rs, dictSeen, nNextID, sNames(x), sKind, sText, sEndText, sEffects) = 1 Then AddRowEach = 1
    If mbFailed Then Exit For
Next x

Exit Function
error:
Call HandleError("AddRowEach")
mbFailed = True
End Function

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

Private Sub ExportRoomMessages(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, _
    dictRoomSpells As Scripting.Dictionary, ByRef bStop As Boolean)
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
        Call ProcessRoom(rs, dictSeen, nNextID, dictRoomSpells)
    End If

    nStatus = BTRCALL(BGETNEXT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

Exit Sub
error:
Call HandleError("ExportRoomMessages")
mbFailed = True
End Sub

Private Sub ProcessRoom(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, dictRoomSpells As Scripting.Dictionary)
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
                Call TrapRow(rs, dictSeen, nNextID, Roomrec.Para3(x), sSrc, sRole)
                Call TrapRow(rs, dictSeen, nNextID, Roomrec.Para4(x), sSrc, sRole)
            Case 24:
                Call TrapRow(rs, dictSeen, nNextID, Roomrec.Para4(x), sSrc, sRole)
                Call TrapRow(rs, dictSeen, nNextID, Roomrec.Para3(x), sSrc, sRole)
        End Select
    End If
Next x

Exit Sub
error:
Call HandleError("ProcessRoom")
mbFailed = True
End Sub

'trap trigger / disarm-fail line 1: Self, nameless, Damage,Trap
Private Sub TrapRow(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, ByVal nMsg As Long, _
    ByVal sSrc As String, ByVal sRole As String)
Dim sRaw As String, sText As String, bFlag As Boolean
On Error GoTo error:

If nMsg <= 1 Then Exit Sub
sRaw = ReadMsgLine(nMsg, 1, sSrc, sRole)
If sRaw = "" Then Exit Sub

bFlag = False
sText = NormalizeWith(sRaw, "", bFlag)
If bFlag Then Call LogUnmatched(sSrc & " " & sRole & " line 1", sRaw)
Call AddRow(rs, dictSeen, nNextID, "", "Self", sText, "", "Damage,Trap")

Exit Sub
error:
Call HandleError("TrapRow")
mbFailed = True
End Sub

'---------------------------------------------------------------------------
' Pass 2: spells
'---------------------------------------------------------------------------

Private Sub ExportSpellMessages(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, _
    dictRoomSpells As Scripting.Dictionary, dictLearn As Scripting.Dictionary, dictItemCast As Scripting.Dictionary, ByRef bStop As Boolean)
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

    Call ProcessSpell(rs, dictSeen, nNextID, dictRoomSpells, dictLearn, dictItemCast)

    nStatus = BTRCALL(BGETNEXT, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

Exit Sub
error:
Call HandleError("ExportSpellMessages")
mbFailed = True
End Sub

Private Sub ProcessSpell(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, _
    dictRoomSpells As Scripting.Dictionary, dictLearn As Scripting.Dictionary, dictItemCast As Scripting.Dictionary)
Dim sName As String, sSrc As String, sEffects As String, sKey As String
Dim sItemNames As String, sTargetTok As String, sNames() As String, nNames As Long
Dim sSelfNames() As String, nSelfNames As Long, bFriendly As Boolean
Dim nDur As Integer, bShort As Boolean, bPlayer As Boolean, bArea As Boolean, nRows As Long
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

bPlayer = (Spellrec.MageryA > 0) Or dictLearn.Exists(sKey)
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
nRows = nRows + SpellStatusRow(rs, dictSeen, nNextID, sSelfNames, nSelfNames, sName, sSrc, nDesc, sDesc3, sEffects, bMergeFail)
If Not bMergeFail Then nRows = nRows + SpellFailRow(rs, dictSeen, nNextID, sSrc, sConf1)
If bPlayer Or Not sItemNames = "" Then
    nRows = nRows + SpellCastRow(rs, dictSeen, nNextID, sNames, nNames, sName, sSrc, bShort, sTargetTok, sEffects)
End If
nRows = nRows + SpellIncomingRow(rs, dictSeen, nNextID, sSelfNames, nSelfNames, sName, sSrc, nDesc, bShort, sEffects)
nRows = nRows + SpellShieldRows(rs, dictSeen, nNextID, sSelfNames, nSelfNames, sSrc, nShock)

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
Private Function SpellStatusRow(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, sNames() As String, _
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
SpellStatusRow = AddRowEach(rs, dictSeen, nNextID, sNames, nNames, "Self", sText, sEnd, sEffects)

Exit Function
error:
Call HandleError("SpellStatusRow")
mbFailed = True
End Function

'ConfuseMsg (101) line 1, not merged into the status row: Self, nameless, Fail only
Private Function SpellFailRow(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, _
    ByVal sSrc As String, ByVal sRaw As String) As Long
Dim sText As String, bFlag As Boolean
On Error GoTo error:

SpellFailRow = 0
If Trim(sRaw) = "" Then Exit Function

bFlag = False
sText = NormalizeWith(sRaw, "", bFlag)
If bFlag Then Call LogUnmatched(sSrc & " confuse line 1", sRaw)
SpellFailRow = AddRow(rs, dictSeen, nNextID, "", "Self", sText, "", "Fail")

Exit Function
error:
Call HandleError("SpellFailRow")
mbFailed = True
End Function

'CastMsgB line 1 (what the caster sees); only for player-castable or item-cast spells
Private Function SpellCastRow(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, sNames() As String, _
    ByVal nNames As Long, ByVal sName As String, ByVal sSrc As String, ByVal bShort As Boolean, ByVal sTargetTok As String, _
    ByVal sEffects As String) As Long
Dim sRaw As String, sText As String, sList As String, bFlag As Boolean
On Error GoTo error:

SpellCastRow = 0
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

If InStr(1, sText, "{dmg}", vbBinaryCompare) = 0 Then Exit Function
If IsStandardForm(mReP1, sText) Then Exit Function
If IsStandardForm(mReP3, sText) Then Exit Function
If IsStandardForm(mReP4, sText) Then Exit Function

If bFlag Then Call LogUnmatched(sSrc & " cast line 1", sRaw)
SpellCastRow = AddRowEach(rs, dictSeen, nNextID, sNames, nNames, "Cast", sText, "", sEffects)

Exit Function
error:
Call HandleError("SpellCastRow")
mbFailed = True
End Function

'CastMsgB line 2 (what the target sees): Self
Private Function SpellIncomingRow(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, sNames() As String, _
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
If bExport = False Then Exit Function

If bFlag Then Call LogUnmatched(sSrc & " cast line 2", sRaw)
SpellIncomingRow = AddRowEach(rs, dictSeen, nNextID, sNames, nNames, "Self", sText, "", sEffects)

Exit Function
error:
Call HandleError("SpellIncomingRow")
mbFailed = True
End Function

'ShockMsg (137): line 1 (shield owner) and line 2 (attacker hit by the shield); Self, Damage,DamageShield
Private Function SpellShieldRows(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, sNames() As String, _
    ByVal nNames As Long, ByVal sSrc As String, ByVal nShock As Long) As Long
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
        nRows = nRows + AddRowEach(rs, dictSeen, nNextID, sNames, nNames, "Self", sText, "", "Damage,DamageShield")
    End If
End If

sRaw = ReadMsgLine(nShock, 2, sSrc, "shock")
If Not sRaw = "" Then
    bFlag = False
    sText = NormalizeWith(sRaw, "{dmg}", bFlag)
    If Not (IsStandardForm(mReP3, sText) Or IsStandardForm(mReP4, sText)) Then
        If bFlag Then Call LogUnmatched(sSrc & " shock line 2", sRaw)
        nRows = nRows + AddRowEach(rs, dictSeen, nNextID, sNames, nNames, "Self", sText, "", "Damage,DamageShield")
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
If SpellHasAbil(20) Or SpellHasAbil(73) Or SpellHasAbil(81) Or SpellHasAbil(84) Or SpellHasAbil(122) Then Call AddEffect(s, "Cure")
If bRoomSpell Then Call AddEffect(s, "RoomSpell")

SpellEffects = s

Exit Function
error:
Call HandleError("SpellEffects")
mbFailed = True
End Function

'---------------------------------------------------------------------------
' Pass 3: monsters
'---------------------------------------------------------------------------

Private Sub ExportMonsterMessages(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, ByRef bStop As Boolean)
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

    Call ProcessMonster(rs, dictSeen, nNextID)

    nStatus = BTRCALL(BGETNEXT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

Exit Sub
error:
Call HandleError("ExportMonsterMessages")
mbFailed = True
End Sub

Private Sub ProcessMonster(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long)
Dim sMon As String, sSrc As String, i As Integer
Dim nDeath As Long, sRaw As String, sText As String, bFlag As Boolean
On Error GoTo error:

sMon = RTrim(ClipNull(Monsterrec.Name))
sSrc = "Monster " & Monsterrec.Number & " (" & sMon & ")"

For i = 0 To 4
    If Monsterrec.AttackType(i) = 1 And Monsterrec.AttackHitMsg(i) > 1 Then
        Call MonsterHitRows(rs, dictSeen, nNextID, Monsterrec.AttackHitMsg(i), sSrc, "attack " & i)
    End If
Next i

'death text only for monsters that regenerate or are game-limited
If Not (Monsterrec.RegenTime > 0 Or Monsterrec.GameLimit > 0) Then Exit Sub

nDeath = Monsterrec.DeathMsg
If nDeath <= 1 And Monsterrec.AttackType(0) = 1 Then nDeath = Monsterrec.AttackHitMsg(0)
If nDeath <= 1 Then Exit Sub

sRaw = ReadMsgLine(nDeath, 3, sSrc, "death")
If sRaw = "" Then Exit Sub

bFlag = False
sText = NormalizeWith(sRaw, "{source}", bFlag)
If bFlag Then Call LogUnmatched(sSrc & " death line 3", sRaw)
Call AddRow(rs, dictSeen, nNextID, sMon, "Other", sText, "", "Death")

Exit Sub
error:
Call HandleError("ProcessMonster")
mbFailed = True
End Sub

'AttackHitMsg line 1 = hit on me (Self); line 2 = hit on someone else (Other), only when line 1 was exported
Private Sub MonsterHitRows(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, ByVal nMsg As Long, _
    ByVal sSrc As String, ByVal sRole As String)
Dim sRaw As String, sText As String, bFlag As Boolean
On Error GoTo error:

sRaw = ReadMsgLine(nMsg, 1, sSrc, sRole)
If sRaw = "" Then Exit Sub

bFlag = False
sText = NormalizeWith(sRaw, "{source}|{dmg}", bFlag)
If IsStandardForm(mReP5, sText) Then Exit Sub
If IsStandardForm(mReP3, sText) Then Exit Sub
If bFlag Then Call LogUnmatched(sSrc & " " & sRole & " line 1", sRaw)
If AddRow(rs, dictSeen, nNextID, "", "Self", sText, "", "Damage,Monster") = 0 Then Exit Sub

sRaw = ReadMsgLine(nMsg, 2, sSrc, sRole)
If sRaw = "" Then Exit Sub

bFlag = False
sText = NormalizeWith(sRaw, "{source}|{target}|{dmg}", bFlag)
If IsStandardForm(mReP6, sText) Then Exit Sub
If bFlag Then Call LogUnmatched(sSrc & " " & sRole & " line 2", sRaw)
Call AddRow(rs, dictSeen, nNextID, "", "Other", sText, "", "Damage,Monster")

Exit Sub
error:
Call HandleError("MonsterHitRows")
mbFailed = True
End Sub
