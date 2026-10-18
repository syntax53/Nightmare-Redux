Attribute VB_Name = "modMMEMessages"
Option Base 0
Option Explicit

'MME export: builds the Messages table (client-relevant game message lines).
'Spec: PLAN-mme-messages-export.md sections 1, 2 and 4.
'Order: rooms (traps + room spell list), then spells, then monsters.

Private mFso As Scripting.FileSystemObject
Private mLog As Scripting.TextStream
Private mdictLogged As Scripting.Dictionary
Private msLogPath As String
Private mnWarnings As Long
Private mbLogBroken As Boolean
Private mbFailed As Boolean

Private mReSelf As Object
Private mReTarget As Object
Private mReHit As Object
Private mReHitOther As Object

Public Function MME_ExportMessages(tabMessages As DAO.Recordset, ByVal sLogPath As String, ByRef bStop As Boolean) As Long
Dim dictSeen As Scripting.Dictionary, dictRoomSpells As Scripting.Dictionary
Dim nNextID As Long
On Error GoTo error:

MME_ExportMessages = 0
If Not InitState(sLogPath) Then
    Call CleanupState
    MME_ExportMessages = -1
    Exit Function
End If

Set dictSeen = New Scripting.Dictionary
dictSeen.CompareMode = vbBinaryCompare
Set dictRoomSpells = New Scripting.Dictionary
dictRoomSpells.CompareMode = vbBinaryCompare

tryagain:
If tabMessages.RecordCount <> 0 Then
    tabMessages.MoveFirst
    tabMessages.Delete
    GoTo tryagain:
End If

nNextID = 1
Call ExportRoomMessages(tabMessages, dictSeen, nNextID, dictRoomSpells, bStop)
If bStop = False And mbFailed = False Then Call ExportSpellMessages(tabMessages, dictSeen, nNextID, dictRoomSpells, bStop)
If bStop = False And mbFailed = False Then Call ExportMonsterMessages(tabMessages, dictSeen, nNextID, bStop)

If mbFailed Then
    MME_ExportMessages = -1
Else
    MME_ExportMessages = mnWarnings
End If

Call CleanupState
Set dictSeen = Nothing
Set dictRoomSpells = Nothing

Exit Function
error:
Call HandleError("MME_ExportMessages")
Call CleanupState
Set dictSeen = Nothing
Set dictRoomSpells = Nothing
MME_ExportMessages = -1
End Function

'---------------------------------------------------------------------------
' State, logging, regex
'---------------------------------------------------------------------------

Private Function InitState(ByVal sLogPath As String) As Boolean
On Error GoTo error:

mbFailed = False
mbLogBroken = False
mnWarnings = 0
msLogPath = sLogPath
Set mLog = Nothing

Set mFso = CreateObject("Scripting.FileSystemObject")
If Not sLogPath = "" Then
    If mFso.FileExists(sLogPath) Then mFso.DeleteFile sLogPath, True
End If

Set mdictLogged = New Scripting.Dictionary
mdictLogged.CompareMode = vbBinaryCompare

Set mReSelf = NewRegExp("^You cast \{spell\} (on|at|upon) \{target\}(,| for| and)? .*\{dmg\} (damage|healing|hit points)!?$")
Set mReTarget = NewRegExp("^\{caster\} casts \{spell\} (on|at|upon) you(,| for| and)? .*\{dmg\} damage!?$")
Set mReHit = NewRegExp("^(The |An? )?\{monster\} [a-z-]+ you for \{dmg\} damage!$")
Set mReHitOther = NewRegExp("^(The |An? )?\{monster\} [a-z-]+ \{target\} for \{dmg\} damage!$")

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
Set mReSelf = Nothing
Set mReTarget = Nothing
Set mReHit = Nothing
Set mReHitOther = Nothing

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

ReadMsgLine = RTrim(sTemp)

Exit Function
error:
Call HandleError("ReadMsgLine")
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

'sList = "|"-separated %s tokens in order, "" for none
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

NormalizeWith = NormalizeLine(sLine, sTokens, nTokens, bFlagged)

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
    ByVal sKind As String, ByVal sText As String, ByVal sEndText As String, ByVal sEffects As String, ByVal nDur As Long) As Long
Dim sKey As String, bTrunc As Boolean
On Error GoTo error:

AddRow = 0
If sText = "" And sEndText = "" Then Exit Function
AddRow = 1

If Len(sName) > 50 Then sName = Left$(sName, 50): bTrunc = True
If Len(sText) > 100 Then sText = Left$(sText, 100): bTrunc = True
If Len(sEndText) > 100 Then sEndText = Left$(sEndText, 100): bTrunc = True
If Len(sEffects) > 255 Then sEffects = Left$(sEffects, 255): bTrunc = True

sKey = sKind & "|" & sName & "|" & sText & "|" & sEndText & "|" & sEffects
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
rs.Fields("Duration") = nDur
rs.Update
nNextID = nNextID + 1

Exit Function
error:
Call HandleError("AddRow")
mbFailed = True
End Function

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
                Call TrapRow(rs, dictSeen, nNextID, Roomrec.Para3(x), "Trap", sSrc, sRole)
                Call TrapRow(rs, dictSeen, nNextID, Roomrec.Para4(x), "TrapFail", sSrc, sRole)
            Case 24:
                Call TrapRow(rs, dictSeen, nNextID, Roomrec.Para4(x), "Trap", sSrc, sRole)
                Call TrapRow(rs, dictSeen, nNextID, Roomrec.Para3(x), "TrapFail", sSrc, sRole)
        End Select
    End If
Next x

Exit Sub
error:
Call HandleError("ProcessRoom")
mbFailed = True
End Sub

Private Sub TrapRow(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, ByVal nMsg As Long, _
    ByVal sKind As String, ByVal sSrc As String, ByVal sRole As String)
Dim sRaw As String, sText As String, bFlag As Boolean
On Error GoTo error:

If nMsg <= 1 Then Exit Sub
sRaw = ReadMsgLine(nMsg, 1, sSrc, sRole)
If sRaw = "" Then Exit Sub

bFlag = False
sText = NormalizeWith(sRaw, "{dmg}", bFlag)
If bFlag Then Call LogUnmatched(sSrc & " " & sRole & " line 1", sRaw)
Call AddRow(rs, dictSeen, nNextID, "", sKind, sText, "", "Damage,Trap", 0)

Exit Sub
error:
Call HandleError("TrapRow")
mbFailed = True
End Sub

'---------------------------------------------------------------------------
' Pass 2: spells
'---------------------------------------------------------------------------

Private Sub ExportSpellMessages(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, _
    dictRoomSpells As Scripting.Dictionary, ByRef bStop As Boolean)
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

    Call ProcessSpell(rs, dictSeen, nNextID, dictRoomSpells)

    nStatus = BTRCALL(BGETNEXT, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

Exit Sub
error:
Call HandleError("ExportSpellMessages")
mbFailed = True
End Sub

Private Sub ProcessSpell(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, dictRoomSpells As Scripting.Dictionary)
Dim sName As String, sSrc As String, sEffects As String
Dim nDur As Integer, bStyle As Boolean, nRows As Long
Dim nDesc As Long, nConf As Long, nShock As Long
On Error GoTo error:

sName = RTrim(ClipNull(Spellrec.Name))
nDur = Spellrec.duration
bStyle = ((Spellrec.MsgStyle And 1) = 1)
sSrc = "Spell " & Spellrec.Number & " (" & sName & ")"

nDesc = SpellAbilityValue(115)
nConf = SpellAbilityValue(101)
nShock = SpellAbilityValue(137)
sEffects = SpellEffects(nDur, dictRoomSpells.Exists(CStr(Spellrec.Number)))

nRows = 0
nRows = nRows + SpellStatusRow(rs, dictSeen, nNextID, sName, sSrc, nDesc, sEffects, nDur)
nRows = nRows + SpellTargetRow(rs, dictSeen, nNextID, sName, sSrc, nDesc, bStyle, sEffects, nDur)
nRows = nRows + SpellSelfRow(rs, dictSeen, nNextID, sName, sSrc, bStyle, sEffects, nDur)
nRows = nRows + SpellFailRow(rs, dictSeen, nNextID, sName, sSrc, nConf, sEffects, nDur)
nRows = nRows + SpellShieldRows(rs, dictSeen, nNextID, sName, sSrc, nShock, nDur)

If sName = "" And nRows > 0 Then Call LogWarning("Spell " & Spellrec.Number & ": blank spell name")

Exit Sub
error:
Call HandleError("ProcessSpell")
mbFailed = True
End Sub

'DescMsg (115): Text = line 3, EndText = line 1
Private Function SpellStatusRow(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, ByVal sName As String, _
    ByVal sSrc As String, ByVal nDesc As Long, ByVal sEffects As String, ByVal nDur As Integer) As Long
Dim sRaw1 As String, sRaw3 As String, sText As String, sEnd As String, bFlag As Boolean
On Error GoTo error:

SpellStatusRow = 0
If nDesc <= 1 Then Exit Function

sRaw3 = ReadMsgLine(nDesc, 3, sSrc, "desc")
sRaw1 = ReadMsgLine(nDesc, 1, sSrc, "desc")
If sRaw3 = "" And sRaw1 = "" Then Exit Function

bFlag = False
sText = NormalizeWith(sRaw3, "", bFlag)
If bFlag Then Call LogUnmatched(sSrc & " desc line 3", sRaw3)

bFlag = False
sEnd = NormalizeWith(sRaw1, "{spell}", bFlag)
If bFlag Then Call LogUnmatched(sSrc & " desc line 1", sRaw1)

SpellStatusRow = AddRow(rs, dictSeen, nNextID, sName, "Status", sText, sEnd, sEffects, nDur)

Exit Function
error:
Call HandleError("SpellStatusRow")
mbFailed = True
End Function

'CastMsgB line 2 (what the target sees)
Private Function SpellTargetRow(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, ByVal sName As String, _
    ByVal sSrc As String, ByVal nDesc As Long, ByVal bStyle As Boolean, ByVal sEffects As String, ByVal nDur As Integer) As Long
Dim sRaw As String, sText As String, sList As String, bFlag As Boolean, bExport As Boolean
On Error GoTo error:

SpellTargetRow = 0
If Spellrec.CastMsgB <= 1 Then Exit Function

sRaw = ReadMsgLine(Spellrec.CastMsgB, 2, sSrc, "cast")
If sRaw = "" Then Exit Function

If bStyle Then
    sList = "{dmg}"
Else
    sList = "{caster}|{spell}|{dmg}"
End If
bFlag = False
sText = NormalizeWith(sRaw, sList, bFlag)

bExport = False
If InStr(1, sText, "{dmg}", vbBinaryCompare) > 0 Then
    If Not IsStandardForm(mReTarget, sText) Then bExport = True
End If
If bExport = False And nDesc <= 1 Then
    If SpellHasStatusAbility() Then bExport = True
End If
If bExport = False Then Exit Function

If bFlag Then Call LogUnmatched(sSrc & " cast line 2", sRaw)
SpellTargetRow = AddRow(rs, dictSeen, nNextID, sName, "Target", sText, "", sEffects, nDur)

Exit Function
error:
Call HandleError("SpellTargetRow")
mbFailed = True
End Function

'CastMsgB line 1 (what the caster sees)
Private Function SpellSelfRow(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, ByVal sName As String, _
    ByVal sSrc As String, ByVal bStyle As Boolean, ByVal sEffects As String, ByVal nDur As Integer) As Long
Dim sRaw As String, sText As String, sList As String, bFlag As Boolean
On Error GoTo error:

SpellSelfRow = 0
If Spellrec.CastMsgB <= 1 Then Exit Function

sRaw = ReadMsgLine(Spellrec.CastMsgB, 1, sSrc, "cast")
If sRaw = "" Then Exit Function

If bStyle Then
    sList = "{target}|{dmg}"
Else
    sList = "{spell}|{target}|{dmg}"
End If
bFlag = False
sText = NormalizeWith(sRaw, sList, bFlag)

If InStr(1, sText, "{dmg}", vbBinaryCompare) = 0 Then Exit Function
If IsStandardForm(mReSelf, sText) Then Exit Function

If bFlag Then Call LogUnmatched(sSrc & " cast line 1", sRaw)
SpellSelfRow = AddRow(rs, dictSeen, nNextID, sName, "Self", sText, "", sEffects, nDur)

Exit Function
error:
Call HandleError("SpellSelfRow")
mbFailed = True
End Function

'ConfuseMsg (101) line 1
Private Function SpellFailRow(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, ByVal sName As String, _
    ByVal sSrc As String, ByVal nConf As Long, ByVal sEffects As String, ByVal nDur As Integer) As Long
Dim sRaw As String, sText As String, bFlag As Boolean
On Error GoTo error:

SpellFailRow = 0
If nConf <= 1 Then Exit Function

sRaw = ReadMsgLine(nConf, 1, sSrc, "confuse")
If sRaw = "" Then Exit Function

bFlag = False
sText = NormalizeWith(sRaw, "", bFlag)
If bFlag Then Call LogUnmatched(sSrc & " confuse line 1", sRaw)
SpellFailRow = AddRow(rs, dictSeen, nNextID, sName, "Fail", sText, "", sEffects, nDur)

Exit Function
error:
Call HandleError("SpellFailRow")
mbFailed = True
End Function

'ShockMsg (137): line 1 = ShieldOwner, line 2 = Shield
Private Function SpellShieldRows(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, ByVal sName As String, _
    ByVal sSrc As String, ByVal nShock As Long, ByVal nDur As Integer) As Long
Dim sRaw As String, sText As String, bFlag As Boolean, nRows As Long
On Error GoTo error:

SpellShieldRows = 0
If nShock <= 1 Then Exit Function

sRaw = ReadMsgLine(nShock, 1, sSrc, "shock")
If Not sRaw = "" Then
    bFlag = False
    sText = NormalizeWith(sRaw, "{attacker}|{dmg}", bFlag)
    If bFlag Then Call LogUnmatched(sSrc & " shock line 1", sRaw)
    nRows = nRows + AddRow(rs, dictSeen, nNextID, sName, "ShieldOwner", sText, "", "Damage,DamageShield", nDur)
End If

sRaw = ReadMsgLine(nShock, 2, sSrc, "shock")
If Not sRaw = "" Then
    bFlag = False
    sText = NormalizeWith(sRaw, "{dmg}", bFlag)
    If bFlag Then Call LogUnmatched(sSrc & " shock line 2", sRaw)
    nRows = nRows + AddRow(rs, dictSeen, nNextID, sName, "Shield", sText, "", "Damage,DamageShield", nDur)
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
If nDur > 0 Then Call AddEffect(s, "Timed")
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

nDeath = Monsterrec.DeathMsg
If nDeath <= 1 And Monsterrec.AttackType(0) = 1 Then nDeath = Monsterrec.AttackHitMsg(0)
If nDeath <= 1 Then Exit Sub

sRaw = ReadMsgLine(nDeath, 3, sSrc, "death")
If sRaw = "" Then Exit Sub

bFlag = False
sText = NormalizeWith(sRaw, "", bFlag)
If bFlag Then Call LogUnmatched(sSrc & " death line 3", sRaw)
Call AddRow(rs, dictSeen, nNextID, sMon, "Death", sText, "", "Death", 0)

Exit Sub
error:
Call HandleError("ProcessMonster")
mbFailed = True
End Sub

'AttackHitMsg line 1 = Hit; line 2 = HitOther (only when line 1 is not the standard form)
Private Sub MonsterHitRows(rs As DAO.Recordset, dictSeen As Scripting.Dictionary, ByRef nNextID As Long, ByVal nMsg As Long, _
    ByVal sSrc As String, ByVal sRole As String)
Dim sRaw As String, sText As String, bFlag As Boolean
On Error GoTo error:

sRaw = ReadMsgLine(nMsg, 1, sSrc, sRole)
If sRaw = "" Then Exit Sub

bFlag = False
sText = NormalizeWith(sRaw, "{monster}|{dmg}", bFlag)
If IsStandardForm(mReHit, sText) Then Exit Sub
If bFlag Then Call LogUnmatched(sSrc & " " & sRole & " line 1", sRaw)
Call AddRow(rs, dictSeen, nNextID, "", "Hit", sText, "", "Damage", 0)

sRaw = ReadMsgLine(nMsg, 2, sSrc, sRole)
If sRaw = "" Then Exit Sub

bFlag = False
sText = NormalizeWith(sRaw, "{monster}|{target}|{dmg}", bFlag)
If IsStandardForm(mReHitOther, sText) Then Exit Sub
If bFlag Then Call LogUnmatched(sSrc & " " & sRole & " line 2", sRaw)
Call AddRow(rs, dictSeen, nNextID, "", "HitOther", sText, "", "Damage", 0)

Exit Sub
error:
Call HandleError("MonsterHitRows")
mbFailed = True
End Sub
