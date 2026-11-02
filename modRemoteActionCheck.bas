Attribute VB_Name = "modRemoteActionCheck"
Option Explicit
Option Base 0

'Remote action check for the Room Editor: for every hidden exit a room is part of (its own, or one its Remote
'Actions work), the whole state - how the exit is set, every remote action aimed at it, which of them open it
'now and in what order, and when that is wrong the values that make all of them count. Only reports. (issue #22)
'
'Rules (MajorMUD 1.11p wccmmud.dll, read from the code):
' - a Remote Action (exit type 12) runs action (Para2 \ 10) on exit (Para2 Mod 10) of its target room; a
'   textblock "remoteaction <room> <msg> <action> <exit>" does the same (@0x450784-0x4507ae, @0x471606)
' - a Hidden exit (type 6) needing N actions (Para2 = N, or -N for any order) keeps one blocked bit per action
'   in Para1: &H10 action 1, &H20 action 2 ... = 16 * (2^N - 1); the exit reset writes that back (@0x44f780)
' - action n clears its bit; with Para2 > 0 (specific order) it is ignored until action n+1 is done, so the
'   highest number goes first; action 0 clears every bit = opens the exit alone (@0x46c8de-0x46cafc)
' - doors (7), traps (9), gates (11) and spell traps (24) take actions too; other exit types ignore them

Private Const HIDDEN_EXIT As Integer = 6
Private Const REMOTE_ACTION As Integer = 12
Private Const BTR_READ_ONLY As Integer = -2 'BOPEN mode
Public Const REMOTE_REPORT_WIDTH As Long = 95 'report columns: longer lines are wrapped (see Wrap)

'index, built on first use from the dat files on position blocks and records of its own, so the editors'
'positions and Roomrec/Messagerec/TextblockRec are never touched:
' mHidden   "map/room/exit" -> "para1|para2"        hidden exits that need actions
' mActions  "map/room/exit" -> actions aimed at it, ";" separated, each "number,map,room,exit,para2,message"
'                              (a textblock: "number,0,<textblock>,-1,0,0")
' mTypes    "map/room"      -> the room's 10 exit types, "," separated
' mFromRoom "map/room"      -> targets of the room's own Remote Actions, ";" separated
' mNames    "map/room"      -> room name;  mCommands  message number -> its line 1 (the commands)
Private mHidden As Scripting.Dictionary
Private mActions As Scripting.Dictionary
Private mTypes As Scripting.Dictionary
Private mFromRoom As Scripting.Dictionary
Private mNames As Scripting.Dictionary
Private mCommands As Scripting.Dictionary
Private mBuilt As Boolean
Private mFailed As Boolean 'the files could not be read: not tried again until RemoteActionIndexReset

Private ChkRoomPos As RoomPosBlockType
Private ChkTextPos As TextblockPosBlockType
Private ChkMsgPos As MessagePosBlockType
Private ChkKey As String * 255

'one exit's remote actions, while a report is made
Private aNum() As Long, aMap() As Long, aRoom() As Long, aExit() As Long, aPara2() As Long, aMsg() As Long
Private aCount As Long

'---------------------------------------------------------------- helpers
Private Function DirName(ByVal nExit As Integer) As String
If nExit < 0 Or nExit > 9 Then
    DirName = "exit " & nExit
Else
    DirName = Choose(nExit + 1, "N", "S", "E", "W", "NE", "NW", "SE", "SW", "U", "D")
End If
End Function

Private Function ExitTypeName(ByVal nType As Integer) As String
If nType < 0 Or nType > 24 Then
    ExitTypeName = "type " & nType
Else
    ExitTypeName = Choose(nType + 1, "Normal", "Spell", "Key", "Item", "Toll", "Action", "Hidden", "Door", _
        "Map Change", "Trap", "Text", "Gate", "Remote Action", "Class", "Race", "Level", "Timed", "Ticket", _
        "User Count", "Block Guard", "Alignment", "Delay", "Cast", "Ability", "Spell Trap")
End If
End Function

Private Function NeededMask(ByVal nCount As Long) As Long
NeededMask = 16 * (2 ^ nCount - 1)
End Function

Private Function PadR(ByVal s As String, ByVal n As Long) As String
PadR = Left$(s & Space$(n), n)
End Function

Private Function PadL(ByVal s As String, ByVal n As Long) As String
PadL = Right$(Space$(n) & s, n)
End Function

'sText cut into lines of REMOTE_REPORT_WIDTH columns at most, the later ones indented by nIndent. it breaks
'after a "> " or "; " (between steps or fixes) when it can, else at a space
Private Function Wrap(ByVal sText As String, ByVal nIndent As Long) As String
Dim sOut As String, n As Long, nBreak As Long
Do While Len(sText) > REMOTE_REPORT_WIDTH
    nBreak = 0
    For n = REMOTE_REPORT_WIDTH To nIndent + 20 Step -1
        If Mid$(sText, n - 1, 2) = "> " Or Mid$(sText, n - 1, 2) = "; " Then nBreak = n: Exit For
    Next n
    If nBreak = 0 Then nBreak = InStrRev(sText, " ", REMOTE_REPORT_WIDTH + 1)
    If nBreak <= nIndent Then nBreak = REMOTE_REPORT_WIDTH 'one long word
    sOut = sOut & RTrim$(Left$(sText, nBreak)) & vbCrLf
    sText = Space$(nIndent) & LTrim$(Mid$(sText, nBreak + 1))
Loop
Wrap = sOut & sText
End Function

Private Function RoomName(ByVal sRoom As String) As String
If mNames.Exists(sRoom) Then RoomName = mNames(sRoom)
End Function

Private Function Where(ByVal i As Long) As String
If aExit(i) < 0 Then
    Where = "textblock " & aRoom(i)
Else
    Where = aMap(i) & "/" & aRoom(i) & " " & DirName(aExit(i))
End If
End Function

Private Function Commands(ByVal i As Long) As String
If aExit(i) < 0 Then
    Commands = "(textblock)"
ElseIf mCommands.Exists(CStr(aMsg(i))) Then
    Commands = Replace(mCommands(CStr(aMsg(i))), ",", ", ")
Else
    Commands = "message " & aMsg(i)
End If
End Function

Private Function FirstCommand(ByVal i As Long) As String
If aExit(i) < 0 Then
    FirstCommand = Where(i)
Else
    FirstCommand = Split(Commands(i) & ",", ",")(0) & " (" & Where(i) & ")"
End If
End Function

'---------------------------------------------------------------- index
'call after anything that changes rooms, textblocks or messages other than through RemoteActionReindexRoom:
'the index is built again the next time it is needed
Public Sub RemoteActionIndexReset()
mBuilt = False
mFailed = False
Set mHidden = Nothing: Set mActions = Nothing: Set mTypes = Nothing
Set mFromRoom = Nothing: Set mNames = Nothing: Set mCommands = Nothing
End Sub

Private Sub AddAction(ByVal sTarget As String, ByVal sEntry As String)
If mActions.Exists(sTarget) Then
    mActions(sTarget) = mActions(sTarget) & ";" & sEntry
Else
    mActions.Add sTarget, sEntry
End If
End Sub

'indexes room tRoom (removing what was indexed for it before)
Private Sub IndexRoom(ByRef tRoom As RoomRecType)
Dim x As Integer, sRoom As String, sTarget As String, sTypes As String, sOwn As String
Dim vOld As Variant, sKept As String, vEntry As Variant, f() As String

sRoom = tRoom.MapNumber & "/" & tRoom.RoomNumber

If mFromRoom.Exists(sRoom) Then 'drop the room's old remote actions from their targets
    For Each vOld In Split(mFromRoom(sRoom), ";")
        If mActions.Exists(vOld) Then
            sKept = ""
            For Each vEntry In Split(mActions(vOld), ";")
                f = Split(vEntry, ",")
                If Not (Val(f(1)) = tRoom.MapNumber And Val(f(2)) = tRoom.RoomNumber And Val(f(3)) >= 0) Then
                    sKept = sKept & IIf(sKept = "", "", ";") & vEntry
                End If
            Next vEntry
            If sKept = "" Then mActions.Remove vOld Else mActions(vOld) = sKept
        End If
    Next vOld
    mFromRoom.Remove sRoom
End If

For x = 0 To 9
    sTypes = sTypes & IIf(x = 0, "", ",") & tRoom.RoomType(x)
    sTarget = sRoom & "/" & x
    If mHidden.Exists(sTarget) Then mHidden.Remove sTarget
    If tRoom.RoomType(x) = HIDDEN_EXIT And tRoom.Para2(x) <> 0 Then
        mHidden.Add sTarget, tRoom.Para1(x) & "|" & tRoom.Para2(x)
    End If
    If tRoom.RoomType(x) = REMOTE_ACTION And tRoom.RoomExit(x) <> 0 And tRoom.Para2(x) >= 0 Then
        sTarget = tRoom.MapNumber & "/" & tRoom.RoomExit(x) & "/" & (tRoom.Para2(x) Mod 10)
        Call AddAction(sTarget, (tRoom.Para2(x) \ 10) & "," & tRoom.MapNumber & "," & tRoom.RoomNumber & "," _
            & x & "," & tRoom.Para2(x) & "," & tRoom.Para1(x))
        If InStr(1, ";" & sOwn & ";", ";" & sTarget & ";") = 0 Then sOwn = sOwn & IIf(sOwn = "", "", ";") & sTarget
    End If
Next x
If mTypes.Exists(sRoom) Then mTypes.Remove sRoom
mTypes.Add sRoom, sTypes
If mNames.Exists(sRoom) Then mNames.Remove sRoom
mNames.Add sRoom, Trim$(ClipNull(tRoom.Name))
If sOwn <> "" Then mFromRoom.Add sRoom, sOwn
End Sub

'after a room is written or inserted: Roomrec holds what was written
Public Sub RemoteActionReindexRoom()
If Not mBuilt Then Exit Sub
On Error GoTo error:
Call IndexRoom(Roomrec)
Exit Sub
error:
Call HandleError("RemoteActionReindexRoom")
RemoteActionIndexReset
End Sub

'after a message is written: Messagerec holds what was written
Public Sub RemoteActionReindexMessage()
If Not mBuilt Then Exit Sub
On Error GoTo error:
If mCommands.Exists(CStr(Messagerec.Number)) Then mCommands.Remove CStr(Messagerec.Number)
mCommands.Add CStr(Messagerec.Number), Trim$(ClipNull(Messagerec.MessageLine1))
Exit Sub
error:
Call HandleError("RemoteActionReindexMessage")
RemoteActionIndexReset
End Sub

Private Function BuildIndex() As Boolean
Dim nStatus As Integer, sPath As String, tRoom As RoomRecType, tText As TextblockRecType, tMsg As MessageRecType
Dim tRoomBuf As RoomDatabufType, tTextBuf As TextblockDataBufType, tMsgBuf As MessageDatabufType
Dim dictText As Scripting.Dictionary, dictMaps As Scripting.Dictionary
Dim vKey As Variant, sText As String, nPos As Long, sParts() As String, vMap As Variant
On Error GoTo error:

BuildIndex = False
mFailed = True
Set mHidden = New Scripting.Dictionary: Set mActions = New Scripting.Dictionary
Set mTypes = New Scripting.Dictionary: Set mFromRoom = New Scripting.Dictionary
Set mNames = New Scripting.Dictionary: Set mCommands = New Scripting.Dictionary
Set dictText = New Scripting.Dictionary: Set dictMaps = New Scripting.Dictionary
Screen.MousePointer = vbHourglass

'no DoEvents while building: a click that showed another room would start a second build on these position blocks
sPath = CStr(ReadINI("Settings", "WGPath" & IIf(WorksWithN = True, "_n", "")))
If Not Right(sPath, 1) = "\" Then sPath = sPath & "\"

'rooms
ChkKey = sPath & "w" & strDatCallLetters & "mp002.dat" & Chr(0)
nStatus = BTRCALL(BOPEN, ChkRoomPos, tRoomBuf, Len(tRoomBuf), ByVal ChkKey, KEY_BUF_LEN, BTR_READ_ONLY)
If Not nStatus = 0 Then GoTo done:
nStatus = BTRCALL(BGETFIRST, ChkRoomPos, tRoomBuf, Len(tRoomBuf), ByVal ChkKey, KEY_BUF_LEN, 0)
Do While nStatus = 0
    RowToStruct tRoomBuf.buf, RoomFldMap, tRoom, LenB(tRoom)
    Call IndexRoom(tRoom)
    vKey = CStr(tRoom.RoomNumber)
    If dictMaps.Exists(vKey) Then
        dictMaps(vKey) = dictMaps(vKey) & "," & tRoom.MapNumber
    Else
        dictMaps.Add vKey, CStr(tRoom.MapNumber)
    End If
    nStatus = BTRCALL(BGETNEXT, ChkRoomPos, tRoomBuf, Len(tRoomBuf), ByVal ChkKey, KEY_BUF_LEN, 0)
Loop
nStatus = BTRCALL(BCLOSE, ChkRoomPos, 0, 0, 0, 0, 0)

'messages: line 1 of a Remote Action's Para1 message lists its commands
ChkKey = sPath & "w" & strDatCallLetters & "msg2.dat" & Chr(0)
nStatus = BTRCALL(BOPEN, ChkMsgPos, tMsgBuf, Len(tMsgBuf), ByVal ChkKey, KEY_BUF_LEN, BTR_READ_ONLY)
If nStatus = 0 Then
    nStatus = BTRCALL(BGETFIRST, ChkMsgPos, tMsgBuf, Len(tMsgBuf), ByVal ChkKey, KEY_BUF_LEN, 0)
    Do While nStatus = 0
        RowToStruct tMsgBuf.buf, MessageFldMap, tMsg, LenB(tMsg)
        If Not mCommands.Exists(CStr(tMsg.Number)) Then mCommands.Add CStr(tMsg.Number), Trim$(ClipNull(tMsg.MessageLine1))
        nStatus = BTRCALL(BGETNEXT, ChkMsgPos, tMsgBuf, Len(tMsgBuf), ByVal ChkKey, KEY_BUF_LEN, 0)
    Loop
    nStatus = BTRCALL(BCLOSE, ChkMsgPos, 0, 0, 0, 0, 0)
End If

'textblocks: "remoteaction <room> <message> <action> <exit>". A textblock names no map (it runs on the map of
'the room that calls it), so it counts for every map that has the room. It only ever supplies actions.
ChkKey = sPath & "w" & strDatCallLetters & "text2.dat" & Chr(0)
nStatus = BTRCALL(BOPEN, ChkTextPos, tTextBuf, TextblockMaxBufSize, ByVal ChkKey, KEY_BUF_LEN, BTR_READ_ONLY)
If nStatus = 0 Then
    nStatus = BTRCALL(BGETFIRST, ChkTextPos, tTextBuf, TextblockMaxBufSize, ByVal ChkKey, KEY_BUF_LEN, 0)
    Do While nStatus = 0
        RowToStruct tTextBuf.buf, TextblockFldMap, tText, LenB(tText)
        vKey = CStr(tText.Number)
        sText = LCase(DecryptTextblock(tText.Data))
        If dictText.Exists(vKey) Then dictText(vKey) = dictText(vKey) & sText Else dictText.Add vKey, sText
        nStatus = BTRCALL(BGETNEXT, ChkTextPos, tTextBuf, TextblockMaxBufSize, ByVal ChkKey, KEY_BUF_LEN, 0)
    Loop
    nStatus = BTRCALL(BCLOSE, ChkTextPos, 0, 0, 0, 0, 0)
End If
For Each vKey In dictText.Keys
    sText = dictText(vKey)
    nPos = InStr(1, sText, "remoteaction ")
    Do While nPos > 0
        sText = Mid$(dictText(vKey), nPos + 13, 40)
        sText = Replace(Replace(Replace(sText, ":", " "), vbLf, " "), vbCr, " ")
        Do While InStr(1, sText, "  ") > 0
            sText = Replace(sText, "  ", " ")
        Loop
        sParts = Split(Trim$(sText), " ")
        If UBound(sParts) >= 3 Then
            If Val(sParts(3)) >= 0 And Val(sParts(3)) <= 9 And dictMaps.Exists(CStr(Val(sParts(0)))) Then
                For Each vMap In Split(dictMaps(CStr(Val(sParts(0)))), ",")
                    Call AddAction(vMap & "/" & Val(sParts(0)) & "/" & Val(sParts(3)), Val(sParts(2)) & ",0," & vKey & ",-1,0,0")
                Next vMap
            End If
        End If
        nPos = InStr(nPos + 13, dictText(vKey), "remoteaction ")
    Loop
Next vKey

mBuilt = True
mFailed = False
BuildIndex = True

done:
Screen.MousePointer = vbDefault
Set dictText = Nothing: Set dictMaps = Nothing
Exit Function
error:
Call HandleError("RemoteActionCheck BuildIndex")
nStatus = BTRCALL(BCLOSE, ChkRoomPos, 0, 0, 0, 0, 0)
nStatus = BTRCALL(BCLOSE, ChkMsgPos, 0, 0, 0, 0, 0)
nStatus = BTRCALL(BCLOSE, ChkTextPos, 0, 0, 0, 0, 0)
Resume done:
End Function

'---------------------------------------------------------------- one exit
'loads the actions aimed at sTarget into a*(), in the order the exit is meant to be worked: action 0 first
'(NMR's MME export and MegaMUD paths list it first), then the numbered ones highest first, then textblocks
Private Sub LoadActions(ByVal sTarget As String)
Dim vAct As Variant, f() As String, i As Long, j As Long, sKey() As String, t As String, n As Long
aCount = 0
If Not mActions.Exists(sTarget) Then Exit Sub
For Each vAct In Split(mActions(sTarget), ";")
    f = Split(vAct, ",")
    ReDim Preserve aNum(aCount): ReDim Preserve aMap(aCount): ReDim Preserve aRoom(aCount)
    ReDim Preserve aExit(aCount): ReDim Preserve aPara2(aCount): ReDim Preserve aMsg(aCount)
    ReDim Preserve sKey(aCount)
    aNum(aCount) = Val(f(0)): aMap(aCount) = Val(f(1)): aRoom(aCount) = Val(f(2))
    aExit(aCount) = Val(f(3)): aPara2(aCount) = Val(f(4)): aMsg(aCount) = Val(f(5))
    If aExit(aCount) < 0 Then
        sKey(aCount) = "2" & Format(aCount, "000")
    ElseIf aNum(aCount) = 0 Then
        sKey(aCount) = "0" & Format(aMap(aCount), "0000") & Format(aRoom(aCount), "000000") & aExit(aCount)
    Else
        sKey(aCount) = "1" & Format(99 - aNum(aCount), "00") & Format(aMap(aCount), "0000") _
            & Format(aRoom(aCount), "000000") & aExit(aCount)
    End If
    aCount = aCount + 1
Next vAct
For i = 1 To aCount - 1 'insertion sort on sKey
    j = i
    Do While j > 0
        If sKey(j - 1) <= sKey(j) Then Exit Do
        t = sKey(j): sKey(j) = sKey(j - 1): sKey(j - 1) = t
        n = aNum(j): aNum(j) = aNum(j - 1): aNum(j - 1) = n
        n = aMap(j): aMap(j) = aMap(j - 1): aMap(j - 1) = n
        n = aRoom(j): aRoom(j) = aRoom(j - 1): aRoom(j - 1) = n
        n = aExit(j): aExit(j) = aExit(j - 1): aExit(j - 1) = n
        n = aPara2(j): aPara2(j) = aPara2(j - 1): aPara2(j - 1) = n
        n = aMsg(j): aMsg(j) = aMsg(j - 1): aMsg(j - 1) = n
        j = j - 1
    Loop
Next i
End Sub

'how the engine opens an exit whose blocked actions are listed in sBlocked (",1,2,"): the indexes of the
'actions to do, "," separated, in order; "" when it cannot open
Private Function OpensWith(ByVal sBlocked As String) As String
Dim i As Long, n As Long, sOut As String, bFound As Boolean
For i = 0 To aCount - 1
    If aNum(i) = 0 Then OpensWith = CStr(i): Exit Function 'action 0 opens it alone
Next i
For n = 10 To 1 Step -1
    If InStr(1, sBlocked, "," & n & ",") > 0 Then
        bFound = False
        For i = 0 To aCount - 1
            If aNum(i) = n Then sOut = sOut & IIf(sOut = "", "", ",") & i: bFound = True: Exit For
        Next i
        If Not bFound Then Exit Function
    End If
Next n
OpensWith = sOut
End Function

Private Function StepsText(ByVal sList As String, ByVal bAlone As Boolean) As String
Dim v As Variant, sOut As String
For Each v In Split(sList, ",")
    sOut = sOut & IIf(sOut = "", "", " > ") & FirstCommand(Val(v))
Next v
StepsText = sOut & IIf(bAlone, " alone", "")
End Function

'The section of the report for one hidden exit; bOK says whether it is all right, sLine is its one-line summary;
'the rooms involved are added to sRooms ("map/room|...")
Private Function ExitSection(ByVal sTarget As String, ByRef bOK As Boolean, ByRef sLine As String, _
    ByRef sRooms As String) As String
Dim sInfo() As String, k() As String, nCount As Long, bOrdered As Boolean, nPara1 As Long, bPara1OK As Boolean
Dim sHere As String, sSetup As String, sNowBlocked As String, sAll As String, n As Long, i As Long
Dim sNow As String, sAfterReset As String, bByZero As Boolean, nZeros As Long, sNums As String
Dim bComplete As Boolean, bAllCount As Boolean, bNoRenumber As Boolean, nFinal As Long
Dim nNewPara2() As Long, sFix As String, sStatus As String, sOut As String, sRow As String, sNowText As String
Dim bFixCol As Boolean, sAfterFix As String, sPrefix As String

sInfo = Split(mHidden(sTarget), "|")
nPara1 = Val(sInfo(0)): nCount = Abs(Val(sInfo(1))): bOrdered = Val(sInfo(1)) > 0
k = Split(sTarget, "/")
sHere = DirName(Val(k(2))) & " exit of " & k(0) & "/" & k(1)
sSetup = "needs " & nCount & " action" & IIf(nCount = 1, "", "s") & IIf(nCount > 1, IIf(bOrdered, ", in order", ", any order"), "")
Call LoadActions(sTarget)

bPara1OK = ((nPara1 And &H3FF0&) = NeededMask(nCount)) Or ((nPara1 And 8) <> 0)
sAll = ","
For n = 1 To nCount: sAll = sAll & n & ",": Next n
If (nPara1 And 8) <> 0 Then
    sNowBlocked = sAll
Else
    sNowBlocked = ","
    For n = 1 To 10
        If (nPara1 And (&H10& * 2 ^ (n - 1))) <> 0 Then sNowBlocked = sNowBlocked & n & ","
    Next n
End If
sNow = OpensWith(sNowBlocked)
If Not bPara1OK Then sAfterReset = OpensWith(sAll)
sNums = ","
For i = 0 To aCount - 1
    sNums = sNums & aNum(i) & ","
    If aNum(i) = 0 Then nZeros = nZeros + 1
    If aExit(i) < 0 Then bNoRenumber = True 'textblocks: their numbers are in the text, the fix is left to the user
Next i
If aCount > 10 Then bNoRenumber = True 'an exit can wait for 10 actions at most
bByZero = (sNow <> "") And (InStr(1, sNums, ",0,") > 0)
bComplete = (sNow <> "") And Not bByZero And (UBound(Split(sNow, ",")) + 1 = nCount)
bAllCount = bComplete
For n = 1 To nCount
    If InStr(1, sNums, "," & n & ",") = 0 Then bAllCount = False
Next n
For i = 0 To aCount - 1
    If aNum(i) > nCount Then bAllCount = False
Next i
If nCount = 1 Then
    bOK = (aCount > 0) And bPara1OK And (sNow <> "")
Else
    bOK = (aCount > 0) And bPara1OK And bComplete And bAllCount
End If

sRooms = sRooms & k(0) & "/" & k(1) & "|"
For i = 0 To aCount - 1
    If aExit(i) >= 0 Then sRooms = sRooms & aMap(i) & "/" & aRoom(i) & "|"
Next i

If aCount = 0 Then
    sLine = "X  " & sHere & " - nothing aims at it, so it never opens"
    ExitSection = Wrap("X  " & sHere & " (" & RoomName(k(0) & "/" & k(1)) & ") - " & sSetup, 3) & vbCrLf _
        & "   Problem:  nothing aims at it, so it never opens" & vbCrLf _
        & "   Opens now:           never" & vbCrLf _
        & Wrap("   Fix:                 add a Remote Action aimed at it, or leave it if it is meant to stay shut", 24) & vbCrLf
    Exit Function
End If
If bOK Then
    sLine = Wrap("OK " & sHere & " - " & sSetup, 3)
    ExitSection = Wrap("OK " & sHere & " (" & RoomName(k(0) & "/" & k(1)) & ") - " & sSetup & " - opens with: " _
        & StepsText(sNow, False), 3) & vbCrLf
    Exit Function
End If

'the fix: every Remote Action aimed at the exit counts, numbered count..1 in the order loaded
ReDim nNewPara2(aCount - 1)
nFinal = nCount
If Not bNoRenumber And aCount <> nCount Then
    nFinal = aCount
    sFix = sFix & IIf(sFix = "", "", "; ") & k(0) & "/" & k(1) & " " & DirName(Val(k(2))) & ": #actions " _
        & IIf(bOrdered, nCount, -nCount) & " -> " & IIf(bOrdered, aCount, -aCount)
End If
If (nPara1 And &H3FF0&) <> NeededMask(nFinal) And (nPara1 And 8) = 0 Then
    sFix = sFix & IIf(sFix = "", "", "; ") & k(0) & "/" & k(1) & " " & DirName(Val(k(2))) & ": type (Para1) " _
        & nPara1 & " -> " & NeededMask(nFinal)
End If
For i = 0 To aCount - 1
    nNewPara2(i) = aPara2(i)
    If Not bNoRenumber And aExit(i) >= 0 Then
        If nFinal > 1 Or aNum(i) <> 0 Then nNewPara2(i) = (nFinal - i) * 10 + Val(k(2))
        If nNewPara2(i) <> aPara2(i) Then
            bFixCol = True
            sFix = sFix & IIf(sFix = "", "", "; ") & Where(i) & ": # " & aPara2(i) & " -> " & nNewPara2(i)
        End If
    End If
    If Not bNoRenumber Then sAfterFix = sAfterFix & IIf(sAfterFix = "", "", ",") & i
Next i

'status
If sNow = "" Then
    sStatus = "never opens: an action it needs is missing" & IIf(bOrdered, " (the others wait for it)", "")
ElseIf bByZero And nCount > 1 Then
    sStatus = "opens with 1 of its " & aCount & " actions"
ElseIf Not bPara1OK Then
    sStatus = "type (Para1) is " & nPara1 & ", not " & NeededMask(nCount) & " (" & nCount & " actions): until it resets it opens with " _
        & (UBound(Split(sNow, ",")) + 1) & " of its " & aCount
Else
    sStatus = "opens with " & (UBound(Split(sNow, ",")) + 1) & " of its " & aCount & " actions"
End If
sLine = Wrap("X  " & sHere & " - " & sStatus & IIf(sFix = "", "", " - fix: " & sFix), 3)

sOut = Wrap("X  " & sHere & " (" & RoomName(k(0) & "/" & k(1)) & ") - " & sSetup, 3) & vbCrLf _
    & Wrap("   Problem:  " & sStatus, 13) & vbCrLf & vbCrLf _
    & "   " & PadR("Remote action", 16) & PadR("Commands", 26) & PadL("#", 4) & "   " & PadR("Now", 34) _
    & IIf(bFixCol, PadL("Fix #", 6), "") & vbCrLf
For i = 0 To aCount - 1
    If sNow = "" Then
        sNowText = IIf(aNum(i) <> 0, "waits for an action that doesn't exist", "-")
    ElseIf bByZero And aNum(i) = 0 Then
        sNowText = IIf(nZeros > 1, "opens it alone (any one of them)", "opens it alone")
    ElseIf InStr(1, "," & sNow & ",", "," & i & ",") > 0 And Not bByZero Then
        If bOrdered And UBound(Split(sNow, ",")) > 0 Then
            sPrefix = Left$(sNow, InStr(1, "," & sNow & ",", "," & i & ",") - 1) 'the steps before this one
            n = Len(sPrefix) - Len(Replace(sPrefix, ",", "")) + 1
            sNowText = "step " & n & " of " & (UBound(Split(sNow, ",")) + 1)
        Else
            sNowText = "needed"
        End If
    ElseIf aNum(i) > nCount Then
        sNowText = "does nothing (above " & nCount & ")"
    Else
        sNowText = "not needed"
    End If
    sRow = "   " & PadR(Where(i), 16) & PadR(Commands(i), 26) _
        & PadL(IIf(aExit(i) < 0, "a" & aNum(i), CStr(aPara2(i))), 4) & "   " & PadR(sNowText, 34)
    If bFixCol Then
        sRow = sRow & PadL(CStr(nNewPara2(i)), 6) & IIf(nNewPara2(i) <> aPara2(i), " <<", "")
    End If
    sOut = sOut & sRow & vbCrLf
Next i
sOut = sOut & vbCrLf & Wrap("   Opens now:           " & IIf(sNow = "", "never", StepsText(sNow, bByZero And nCount > 1)), 24) & vbCrLf
If sAfterReset <> "" Then sOut = sOut & Wrap("   After its next reset: " & StepsText(sAfterReset, False), 25) & vbCrLf
If sAfterFix <> "" Then
    sOut = sOut & Wrap("   After the fix:       " & StepsText(sAfterFix, False) _
        & IIf(bOrdered And aCount > 1, ", in this order", ""), 24) & vbCrLf
End If
If sFix <> "" Then sOut = sOut & Wrap("   Fix:                 " & sFix, 24) & vbCrLf
ExitSection = sOut
End Function

'---------------------------------------------------------------- the room
'The report for room nMap/nRoom (Roomrec holds it, as it does after DispRoomInfo). Returns the text ("" when the
'room has no remote actions and no hidden exit that needs them); sSummary = one line per hidden exit and per
'problem, for the top of the report window; nProblems = exits that are wrong + actions aimed at missing rooms, nIdle = actions
'aimed at exits that ignore them, sRooms = the other rooms involved ("map/room|...").
Public Function RemoteActionReport(ByVal nMap As Long, ByVal nRoom As Long, ByRef sSummary As String, _
    ByRef nProblems As Long, ByRef nIdle As Long, ByRef sRooms As String) As String
Dim x As Integer, sOut As String, sTarget As String, sDone As String, sTypes() As String, nType As Integer
Dim bOK As Boolean, sNotes As String, sRoomKey As String, sSection As String, sAll As String, v As Variant
Dim sLine As String, nInfo As Long
On Error GoTo error:

sSummary = "": nProblems = 0: nIdle = 0: sRooms = ""
If Not mBuilt Then
    If mFailed Then Exit Function
    If Not BuildIndex() Then Exit Function
End If

sDone = "|"
For x = 0 To 9 'the room's own hidden exits
    sTarget = nMap & "/" & nRoom & "/" & x
    If mHidden.Exists(sTarget) Then
        sDone = sDone & sTarget & "|"
        sSection = ExitSection(sTarget, bOK, sLine, sAll)
        If Not bOK Then nProblems = nProblems + 1
        sOut = sOut & IIf(sOut = "", "", String$(REMOTE_REPORT_WIDTH, "-") & vbCrLf) & sSection
        sSummary = sSummary & sLine & vbCrLf
    End If
Next x
For x = 0 To 9 'its Remote Actions
    If Roomrec.RoomType(x) = REMOTE_ACTION And Roomrec.RoomExit(x) <> 0 And Roomrec.Para2(x) >= 0 Then
        sRoomKey = nMap & "/" & Roomrec.RoomExit(x)
        sTarget = sRoomKey & "/" & (Roomrec.Para2(x) Mod 10)
        If Not mTypes.Exists(sRoomKey) Then
            nProblems = nProblems + 1
            sNotes = sNotes & "   X  " & DirName(x) & ": works room " & sRoomKey & ", which doesn't exist" & vbCrLf
            sSummary = sSummary & "X  " & DirName(x) & " remote action works room " & sRoomKey & ", which doesn't exist" & vbCrLf
        Else
            sTypes = Split(mTypes(sRoomKey), ",")
            nType = Val(sTypes(Roomrec.Para2(x) Mod 10))
            If mHidden.Exists(sTarget) Then
                If InStr(1, sDone, "|" & sTarget & "|") = 0 Then
                    sDone = sDone & sTarget & "|"
                    sSection = ExitSection(sTarget, bOK, sLine, sAll)
                    If Not bOK Then nProblems = nProblems + 1
                    sOut = sOut & IIf(sOut = "", "", String$(REMOTE_REPORT_WIDTH, "-") & vbCrLf) & sSection
                    sSummary = sSummary & sLine & vbCrLf
                End If
            ElseIf nType = 6 Or nType = 7 Or nType = 9 Or nType = 11 Or nType = 24 Then
                nInfo = nInfo + 1
                sNotes = sNotes & "   i  " & DirName(x) & ": works the " & DirName(Roomrec.Para2(x) Mod 10) & " " _
                    & LCase(ExitTypeName(nType)) & " of " & sRoomKey _
                    & IIf(nType = 7 Or nType = 11, " (toggles it)", IIf(nType = 6, " (opens it)", " (springs it)")) & vbCrLf
            Else
                nIdle = nIdle + 1
                sNotes = sNotes & "   o  " & DirName(x) & ": works the " & DirName(Roomrec.Para2(x) Mod 10) & " exit of " _
                    & sRoomKey & ", a " & ExitTypeName(nType) & " exit: does nothing" & vbCrLf
            End If
        End If
    End If
Next x
If sNotes <> "" Then
    sOut = sOut & IIf(sOut = "", "", String$(REMOTE_REPORT_WIDTH, "-") & vbCrLf) & "Other remote actions of this room" & vbCrLf & sNotes
End If
If nInfo > 0 Then
    sSummary = sSummary & "i  " & nInfo & " remote action" & IIf(nInfo = 1, "", "s") & " of this room " _
        & IIf(nInfo = 1, "works a door, gate or trap", "work doors, gates or traps") & vbCrLf
End If
If nIdle > 0 Then
    sSummary = sSummary & "o  " & nIdle & " remote action" & IIf(nIdle = 1, "", "s") & " of this room " _
        & IIf(nIdle = 1, "works an exit that ignores it", "work exits that ignore them") & vbCrLf
End If
If Right$(sSummary, 2) = vbCrLf Then sSummary = Left$(sSummary, Len(sSummary) - 2)

'the other rooms involved, each once
For Each v In Split(sAll, "|")
    If v <> "" And v <> nMap & "/" & nRoom And InStr(1, "|" & sRooms, "|" & v & "|") = 0 Then sRooms = sRooms & v & "|"
Next v
RemoteActionReport = sOut
Exit Function
error:
Call HandleError("RemoteActionReport")
RemoteActionReport = ""
End Function

'name of a room for the report window's room list
Public Function RemoteActionRoomName(ByVal sRoom As String) As String
If mBuilt Then RemoteActionRoomName = RoomName(sRoom)
End Function
