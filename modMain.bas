Attribute VB_Name = "modMain"
Option Base 0
Option Explicit

Public bKeepSettingsOpen As Boolean

Public Const LongOffset = 4294967296#
Public Const MaxLong = 2147483647
Public Const IntOffset = 65536
Public Const MaxInt = 32767

Private Const CB_SHOWDROPDOWN = &H14F
'Private Const CB_GETITEMHEIGHT = &H154
'Private Const CB_SETDROPPEDWIDTH = &H160
'Private Const CB_GETDROPPEDWIDTH = &H15F
Private Const CB_SETDROPPEDCONTROLRECT = &H160
'Private Const DT_CALCRECT = &H400

Public Const STATE_SYSTEM_FOCUSABLE = &H100000
Public Const STATE_SYSTEM_INVISIBLE = &H8000
Public Const STATE_SYSTEM_OFFSCREEN = &H10000
Public Const STATE_SYSTEM_UNAVAILABLE = &H1
Public Const STATE_SYSTEM_PRESSED = &H8
Public Const CCHILDREN_TITLEBAR = 5
Public Const LB_GETITEMRECT = &H198
Public Const CB_GETDROPPEDCONTROLRECT = &H15F
Public Const CB_GETITEMHEIGHT = &H154
Public Const MF_BYPOSITION = &H400&
Public Const MF_DISABLED = &H2&
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Private Const CB_SETDROPPEDWIDTH = &H160
Private Const CB_GETDROPPEDWIDTH = &H15F
Private Const DT_CALCRECT = &H400

Public TITLEBAR_OFFSET As Integer
Public bUseCPU As Boolean
Public bAbilityDBOpen As Boolean
Public bDisableWriting As Boolean
Public bOnlyNames As Boolean
Public bOppositeListOrder As Boolean
Public sAppVersion As String
Public sMenuCaption As String
Public strDatCallLetters As String * 2
Public bStopControlBuild As Boolean
Public dbAbilities As Database
Public rsAbilities As Recordset

'------------------------------------------------------------------------------------------
' START: Placeholders for the DAT file suffixes: NT or non-NT versions (MBBS/WG)
'------------------------------------------------------------------------------------------
'these get overwritten with either NT or non-NT strings
Public strDatSuffix_ACTS As String
Public strDatSuffix_BANKS As String
Public strDatSuffix_CLASS As String
Public strDatSuffix_GANGS As String
Public strDatSuffix_ITEMS As String
Public strDatSuffix_KNMSR As String
Public strDatSuffix_MP As String
Public strDatSuffix_MSG As String
Public strDatSuffix_RACE As String
Public strDatSuffix_SHOPS As String
Public strDatSuffix_SPELS As String
Public strDatSuffix_TEXT As String
Public strDatSuffix_UPDAT As String
Public strDatSuffix_USERS As String
'non-NT DAT suffixes...
Public Const strDatSuffixNNT_ACTS As String = "ACTS.DAT"
Public Const strDatSuffixNNT_BANKS As String = "BANKS.DAT"
Public Const strDatSuffixNNT_CLASS As String = "CLASS.DAT"
Public Const strDatSuffixNNT_GANGS As String = "GANGS.DAT"
Public Const strDatSuffixNNT_ITEMS As String = "ITEMS.DAT"
Public Const strDatSuffixNNT_KNMSR As String = "KNMSR.DAT"
Public Const strDatSuffixNNT_MP As String = "MP001.DAT"
Public Const strDatSuffixNNT_MSG As String = "MSG.DAT"
Public Const strDatSuffixNNT_RACE As String = "RACE.DAT"
Public Const strDatSuffixNNT_SHOPS As String = "SHOPS.DAT"
Public Const strDatSuffixNNT_SPELS As String = "SPELS.DAT"
Public Const strDatSuffixNNT_TEXT As String = "TEXT.DAT"
Public Const strDatSuffixNNT_UPDAT As String = "UPDAT.DAT"
Public Const strDatSuffixNNT_USERS As String = "USERS.DAT"
'NT DAT suffixes...
Public Const strDatSuffixNT_ACTS As String = "acts2.dat"
Public Const strDatSuffixNT_BANKS As String = "bank2.dat"
Public Const strDatSuffixNT_CLASS As String = "clas2.dat"
Public Const strDatSuffixNT_GANGS As String = "gang2.dat"
Public Const strDatSuffixNT_ITEMS As String = "item2.dat"
Public Const strDatSuffixNT_KNMSR As String = "knms2.dat"
Public Const strDatSuffixNT_MP As String = "mp002.dat"
Public Const strDatSuffixNT_MSG As String = "msg2.dat"
Public Const strDatSuffixNT_RACE As String = "race2.dat"
Public Const strDatSuffixNT_SHOPS As String = "shop2.dat"
Public Const strDatSuffixNT_SPELS As String = "spel2.dat"
Public Const strDatSuffixNT_TEXT As String = "text2.dat"
Public Const strDatSuffixNT_UPDAT As String = "upda2.dat"
Public Const strDatSuffixNT_USERS As String = "user2.dat"
'------------------------------------------------------------------------------------------
' END
'------------------------------------------------------------------------------------------

Type MGILType
    nNumber(10) As Long
    'sName(20) As String
End Type
Public MGIL() As MGILType 'MGIL=Monster Group Index List
Public ControlRoomList As New Dictionary
Public Races() As ArrayRec
Public Classes() As ArrayRec

Public nMonsterSingleAttackCopy(9) As Long
Public nMonsterAllAttackCopy(49) As Long

Public nRoomCopyPaste(15) As Long
Public sRoomCopyPaste As String
Public sRoomCopyDesc(6) As String

Public sSpellCopyPaste(11) As String

Public Enum enm_wCommand
    HELP_CONTEXT = &H1&
    HELP_QUIT = &H2&
    HELP_CONTENTS = &H3&
    HELP_INDEX = &H3&
    HELP_HELPONHELP = &H4&
    HELP_SETCONTENTS = &H5&
    HELP_SETINDEX = &H5&
    HELP_CONTEXTPOPUP = &H8&
    HELP_FORCEFILE = &H9&
    HELP_CONTEXTMENU = &HA&
    HELP_FINDER = &HB&
    HELP_WM_HELP = &HC&
    HELP_SETPOPUP_POS = &HD&
    HELP_FORCE_GID = &HE&
    
    HELP_TAB = &HF&
    HELP_KEY = &H101&
    HELP_COMMAND = &H102&
    HELP_PARTIALKEY = &H105&
    HELP_MULTIKEY = &H201&
    HELP_SETWINPOS = &H203&
End Enum


Public WinVal As Long

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type ArrayRec
    Number As Long
    Name As String * 29
End Type

Public Enum ListDataType
    ldtString = 0
    ldtNumber = 1
    ldtDateTime = 2
End Enum

Type TITLEBARINFO
    cbSize As Long
    rcTitleBar As RECT
    rgstate(CCHILDREN_TITLEBAR) As Long
End Type

Public Enum QBColorCode
    black = 0
    Blue = 1
    Green = 2
    Cyan = 3
    Red = 4
    Magenta = 5
    Yellow = 6
    white = 7
    Grey = 8
    BrightBlue = 9
    BrightGreen = 10
    BrightCyan = 11
    BrightRed = 12
    BrightMagenta = 13
    BrightYellow = 14
    BrightWhite = 15
End Enum

Public Enum DatVersion '***NEWMUDVER***
    v111h = 0
    v111i = 1
    v111j = 2
    v111k = 3
    v111l = 4
    v111m = 5
    v111n = 6
    v111o = 7
    v111p13 = 8
    v111p = 9
End Enum
Public eDatFileVersion As DatVersion

Public Enum eExpandBy
    Percent50 = 0
    Percent75 = 1
    DoubleWidth = 2
    TripleWidth = 3
    QuadWidth = 4
    NoExpand = 5
End Enum
Public Enum eExpandType
    WidthOnly = 0
    HeightOnly = 1
    HeightAndWidth = 2
End Enum
Public Type POINTAPI
   x As Long
   y As Long
End Type


Public Declare Function MoveWindow Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal x As Long, ByVal y As Long, _
   ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal bRepaint As Long) As Long
   
Public Declare Function GetWindowRect Lib "user32" _
  (ByVal hwnd As Long, _
   lpRect As RECT) As Long
   
Public Declare Function ScreenToClient Lib "user32" _
  (ByVal hwnd As Long, _
   lpPoint As POINTAPI) As Long
   
Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
   
Public Declare Function SendMessageLong Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long
        
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Public Declare Function WinHelpString Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As enm_wCommand, ByVal strData As String) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetTitleBarInfo Lib "user32" (ByVal hwnd As Long, ByRef pti As TITLEBARINFO) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function DrawText Lib "user32" Alias _
    "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, _
    ByVal nCount As Long, lpRect As RECT, ByVal wFormat _
    As Long) As Long
    
Declare Function SetWindowPos Lib "user32" _
      (ByVal hwnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long

Public Function IsDimmed(Arr As Variant) As Boolean
On Error GoTo ReturnFalse
  IsDimmed = UBound(Arr) >= LBound(Arr)
ReturnFalse:
End Function

Public Function RandomNumber(startNum As Integer, endNum As Integer) As Integer
    Randomize
    RandomNumber = Int(((endNum - startNum + 1) * Rnd) + startNum)
End Function

Public Sub ExpandCombo(ByRef Combo As ComboBox, ByVal ExpandType As eExpandType, _
    ByVal ExpandBy As eExpandBy, Optional ByVal hFrame As Long)

    Dim lRet As Long
    Dim pt As POINTAPI
    Dim rc As RECT
    Dim lComboWidth As Long
    Dim lNewHeight As Long
    Dim lItemHeight As Long
    
    If ExpandType <> HeightOnly Then
        lComboWidth = (Combo.Width / Screen.TwipsPerPixelX)
        Select Case ExpandBy
            Case 0:
                lComboWidth = lComboWidth + (lComboWidth * 0.5)
            Case 1:
                lComboWidth = lComboWidth + (lComboWidth * 0.75)
            Case 2:
                lComboWidth = lComboWidth * 2
            Case 3:
                lComboWidth = lComboWidth * 3
            Case 4:
                lComboWidth = lComboWidth * 4
        End Select
        lRet = SendMessage(Combo.hwnd, CB_SETDROPPEDCONTROLRECT, lComboWidth, 0)
        
    End If
    
    If ExpandType <> WidthOnly Then
        lComboWidth = Combo.Width / Screen.TwipsPerPixelX
        lItemHeight = SendMessage(Combo.hwnd, CB_GETITEMHEIGHT, 0, 0)
        Select Case ExpandBy
            Case 1:
                'lComboWidth = lComboWidth + (lComboWidth * 0.75)
                lNewHeight = lItemHeight * 16
            Case 2:
                'lComboWidth = lComboWidth * 2
                lNewHeight = lItemHeight * 18
            Case 3:
                'lComboWidth = lComboWidth * 3
                lNewHeight = lItemHeight * 26
            Case 4:
                'lComboWidth = lComboWidth * 4
                lNewHeight = lItemHeight * 32
            Case Else:
                lNewHeight = lItemHeight * 14
                'lComboWidth = lComboWidth + (lComboWidth * 0.5)
        End Select
        Call GetWindowRect(Combo.hwnd, rc)
        pt.x = rc.Left
        pt.y = rc.Top
        Call ScreenToClient(hFrame, pt)
        Call MoveWindow(Combo.hwnd, pt.x, pt.y, lComboWidth, lNewHeight, True)
    End If
    
End Sub

Public Function GetAbilityName(ByVal nNumber As Long) As String
On Error GoTo error:

GetAbilityName = "Ability " & nNumber
If bAbilityDBOpen = False Then Exit Function

rsAbilities.Index = "PrimaryKey"
rsAbilities.Seek "=", nNumber
If Not rsAbilities.NoMatch Then
    GetAbilityName = rsAbilities.Fields("Name")
End If

out:
Exit Function
error:
Call HandleError("GetAbilityName")
Resume out:
End Function
Public Function ExtractValueFromString(ByVal sWholeString As String, ByVal sSearchText As String) As Long
Dim x As Long, y As Long, sChar As String * 1

On Error GoTo error:

x = InStr(1, sWholeString, sSearchText)
If x > 0 Then
    x = x + Len(sSearchText) 'position x just after the search text
    y = x
    Do Until y > Len(sWholeString)
        sChar = Mid(sWholeString, y, 1)
        Select Case sChar
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
            Case " ":
                If y > x Then
                    Exit Do
                Else
                    x = x + 1
                End If
            Case Else: Exit Do
        End Select
        y = y + 1
    Loop
    If y > x Then ExtractValueFromString = Val(Mid(sWholeString, x, y - x))
    'If ExtractValueFromString = "0" Then ExtractValueFromString = ""
End If

out:
Exit Function
error:
Call HandleError("ExtractValueFromString")
Resume out:

End Function

Public Function GetBinaryValue(ByRef BinArray() As Byte, nOffset As Integer, nLength As Integer) As Variant
On Error GoTo error:

Select Case nLength
    Case 1:
        GetBinaryValue = (BinArray(nOffset))
    Case 2:
        GetBinaryValue = (BinArray(nOffset + 1) * 256) + BinArray(nOffset)
    Case 4:
        GetBinaryValue = (((((BinArray(nOffset + 3) * 256) + BinArray(nOffset + 2)) * 256) _
            + BinArray(nOffset + 1)) * 256) + BinArray(nOffset)
End Select

Exit Function
error:
Call HandleError("GetBinaryValue")

End Function

Public Function GetBinaryString(ByRef BinArray() As Byte, nOffset As Integer, nLength As Integer) As String
On Error GoTo error:
Dim x As Integer

For x = nOffset To nOffset + nLength - 1
    GetBinaryString = GetBinaryString & Chr(BinArray(x))
Next x

Exit Function
error:
Call HandleError("GetBinaryString")

End Function

Public Sub SetBinaryValue(nValue As Variant, ByRef BinArray() As Byte, nOffset As Integer, nLength As Integer)
On Error GoTo error:

Select Case nLength
    Case 1:
        BinArray(nOffset) = (nValue)
    Case 2:
        BinArray(nOffset + 1) = (Fix(nValue / 256) Mod 256)
        BinArray(nOffset) = (nValue Mod 256)
    Case 4:
        BinArray(nOffset + 3) = Fix(nValue / 16777216)
        BinArray(nOffset + 2) = (Fix(nValue / 65536) Mod 256)
        BinArray(nOffset + 1) = (Fix(nValue / 256) Mod 256)
        BinArray(nOffset) = (nValue Mod 256)
End Select

Exit Sub
error:
Call HandleError("SetBinaryValue")

End Sub

Public Sub SetBinaryString(ByVal strString As String, ByRef BinArray() As Byte, nOffset As Integer, nLength As Integer)
On Error GoTo error:
Dim x As Integer

If Len(strString) < nLength Then
    strString = strString & String(nLength - Len(strString), Chr(0))
End If

For x = nOffset To nOffset + nLength - 1
    BinArray(x) = Asc(Mid(strString, Abs(x - nOffset + 1), 1))
Next x

Exit Sub
error:
Call HandleError("SetBinaryString")

End Sub

Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) _
   As Long

   If Topmost = True Then 'Make the window topmost
      SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, _
         0, FLAGS)
   Else
      SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, _
         0, 0, FLAGS)
      SetTopMostWindow = False
   End If
End Function

'The call to bring up the help topic page is;
'WinHelpString hwnd, "C:\Myprog\MyHelpFile.hlp", HELP_COMMAND, "JumpId(""C:/Myprog/MyHelpFile.hlp"",""TopicID"")"

Public Function PutCommas(ByVal sNumber As String) As String
On Error GoTo error:
Dim x As Integer, y As Integer, z As Integer

If Len(sNumber) < 4 Then
    PutCommas = sNumber
    Exit Function
End If

z = 1
y = Len(sNumber)
For x = 1 To y
    PutCommas = Mid(sNumber, y - x + 1, 1) & PutCommas
    
    If z > 2 And Not z = y Then
        If z Mod 3 = 0 Then PutCommas = "," & PutCommas
    End If
    
    z = z + 1
Next

Exit Function
error:
HandleError
End Function

Public Function FormIsLoaded(ByVal FormName As String) As Boolean
Dim frmForm As Form

For Each frmForm In Forms
    If LCase(frmForm.Name) = LCase(FormName) Then FormIsLoaded = True: Exit For
    Set frmForm = Nothing
Next

Set frmForm = Nothing
End Function

Public Sub CopyRoomForm()
Dim frmRoomCopy As Form
Set frmRoomCopy = New frmRoom
frmRoomCopy.RoomCopy = True
Load frmRoomCopy
frmRoomCopy.Top = frmRoomCopy.Top + 300
frmRoomCopy.Left = frmRoomCopy.Left + 300
End Sub

Public Sub CopyItemForm()
Dim frmItemCopy As New frmItem
Load frmItemCopy
frmItemCopy.Top = frmItemCopy.Top + 300
frmItemCopy.Left = frmItemCopy.Left + 300
End Sub

Public Sub CopySpellForm()
Dim frmSpellCopy As New frmSpell
Load frmSpellCopy
frmSpellCopy.Top = frmSpellCopy.Top + 300
frmSpellCopy.Left = frmSpellCopy.Left + 300
End Sub

Public Sub CopyActionForm()
Dim frmActionCopy As New frmAction
Load frmActionCopy
frmActionCopy.Top = frmActionCopy.Top + 300
frmActionCopy.Left = frmActionCopy.Left + 300
End Sub

Public Sub CopyMonsterForm()
Dim frmMonsterCopy As New frmMonster
Load frmMonsterCopy
frmMonsterCopy.Top = frmMonsterCopy.Top + 300
frmMonsterCopy.Left = frmMonsterCopy.Left + 300
End Sub

Public Sub CopyRaceForm()
Dim frmRaceCopy As New frmRace
Load frmRaceCopy
frmRaceCopy.Top = frmRaceCopy.Top + 300
frmRaceCopy.Left = frmRaceCopy.Left + 300
End Sub

Public Sub CopyClassForm()
Dim frmClassCopy As New frmClass
Load frmClassCopy
frmClassCopy.Top = frmClassCopy.Top + 300
frmClassCopy.Left = frmClassCopy.Left + 300
End Sub

Public Sub CopyShopForm()
Dim frmShopCopy As New frmShop
Load frmShopCopy
frmShopCopy.Top = frmShopCopy.Top + 300
frmShopCopy.Left = frmShopCopy.Left + 300
End Sub

Public Sub CopyUserForm()
Dim frmUserCopy As New frmUser
Load frmUserCopy
frmUserCopy.Top = frmUserCopy.Top + 300
frmUserCopy.Left = frmUserCopy.Left + 300
End Sub

Public Sub CopyMessageForm()
Dim frmMessageCopy As New frmMessage
Load frmMessageCopy
frmMessageCopy.Top = frmMessageCopy.Top + 300
frmMessageCopy.Left = frmMessageCopy.Left + 300
End Sub

Public Sub CopyTextblockForm()
Dim frmTextblockCopy As New frmTextblock
Load frmTextblockCopy
frmTextblockCopy.Top = frmTextblockCopy.Top + 300
frmTextblockCopy.Left = frmTextblockCopy.Left + 300
End Sub

Public Sub LockMenus()

frmMain.mnuFile.Enabled = False
frmMain.mnuEdit.Enabled = False
frmMain.mnuTools.Enabled = False
frmMain.mnuWindow.Enabled = False
frmMain.mnuHelp.Enabled = False

End Sub

Public Sub UnLockMenus()

frmMain.mnuFile.Enabled = True
frmMain.mnuEdit.Enabled = True
frmMain.mnuTools.Enabled = True
frmMain.mnuWindow.Enabled = True
frmMain.mnuHelp.Enabled = True

End Sub

Public Function GetShortName(ByVal sLongFileName As String) As String
Dim lRetVal As Long, sShortPathName As String, iLen As Integer

'Set up buffer area for API function call return
sShortPathName = Space(255)
iLen = Len(sShortPathName)

'Call the function
lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
'Strip away unwanted characters.
GetShortName = Left(sShortPathName, lRetVal)

End Function

Public Function GetLongDirName(ByVal short_name As String) As String
Dim pos As Integer
Dim result As String
Dim long_name As String

    ' Start after the drive letter if any.
    If Mid$(short_name, 2, 1) = ":" Then
        result = Left$(short_name, 2)
        pos = 3
    Else
        result = ""
        pos = 1
    End If

    ' Consider each section in the file name.
    Do While pos > 0
        ' Find the next \.
        pos = InStr(pos + 1, short_name, "\")

        ' Get the next piece of the path.
        If Not pos = 0 Then
            long_name = Dir$(Left$(short_name, pos - 1), _
                vbNormal + vbHidden + vbSystem + _
                vbDirectory)
            result = result & "\" & long_name
        End If
        
    Loop

    GetLongDirName = result
End Function

Public Function GetLongFileName(ByVal short_name As String) As String
Dim pos As Integer
Dim result As String
Dim long_name As String

    ' Start after the drive letter if any.
    If Mid$(short_name, 2, 1) = ":" Then
        result = Left$(short_name, 2)
        pos = 3
    Else
        result = ""
        pos = 1
    End If

    ' Consider each section in the file name.
    Do While pos > 0
        ' Find the next \.
        pos = InStr(pos + 1, short_name, "\")

        ' Get the next piece of the path.
        If pos = 0 Then
            long_name = Dir$(short_name, vbNormal + _
                vbHidden + vbSystem + vbDirectory)
        Else
            long_name = Dir$(Left$(short_name, pos - 1), _
                vbNormal + vbHidden + vbSystem + _
                vbDirectory)
        End If
        result = result & "\" & long_name
    Loop

    GetLongFileName = result
End Function

Public Sub HandleError(Optional ByVal sSource As String)

If Not sSource = "" Then
    sSource = "Error " & Err.Number & " in [" & sSource & "]:" & vbCrLf
Else
    sSource = "Error " & Err.Number & ": "
End If

Select Case Err.Number
    Case 6: MsgBox sSource & "Overflow (One of the values entered was too large)."
    Case 53: MsgBox sSource & "File not found." & vbCrLf & vbCrLf _
        & "If you recieve this error on load it MAY be because the btrieve engine can not initiate." & vbCrLf _
        & "Try copying wbtrv32.dll, w32mkset.dll, w32mkrc.dll, and w32mkde.exe from" & vbCrLf _
        & "the NMR directory to your system (Win9x) or system32 (WinNT+) directory.", vbCritical
    Case 70: MsgBox sSource & "File is locked by another process!"
    Case Else: MsgBox sSource & Err.Source & vbCrLf & Err.Description, vbExclamation
End Select

Err.clear
End Sub

Public Sub TransparentForm(frm As Form)
    frm.ScaleMode = vbPixels
    Const RGN_DIFF = 4
    Const RGN_OR = 2

    Dim outer_rgn As Long
    Dim inner_rgn As Long
    Dim wid As Single
    Dim hgt As Single
    Dim border_width As Single
    Dim title_height As Single
    Dim ctl_left As Single
    Dim ctl_top As Single
    Dim ctl_right As Single
    Dim ctl_bottom As Single
    Dim control_rgn As Long
    Dim combined_rgn As Long
    Dim ctl As Control

    If frm.WindowState = vbMinimized Then Exit Sub

    ' Create the main form region.
    wid = frm.ScaleX(frm.Width, vbTwips, vbPixels)
    hgt = frm.ScaleY(frm.Height, vbTwips, vbPixels)
    outer_rgn = CreateRectRgn(0, 0, wid, hgt)

    border_width = (wid - frm.ScaleWidth) / 2
    title_height = hgt - border_width - frm.ScaleHeight
    inner_rgn = CreateRectRgn(border_width, title_height, wid - border_width, _
        hgt - border_width)

    ' Subtract the inner region from the outer.
    combined_rgn = CreateRectRgn(0, 0, 0, 0)
    CombineRgn combined_rgn, outer_rgn, inner_rgn, RGN_DIFF

    ' Create the control regions.
    For Each ctl In frm.Controls
        If ctl.Container Is frm Then
            ctl_left = frm.ScaleX(ctl.Left, frm.ScaleMode, vbPixels) _
                + border_width
            ctl_top = frm.ScaleX(ctl.Top, frm.ScaleMode, vbPixels) + title_height
            ctl_right = frm.ScaleX(ctl.Width, frm.ScaleMode, vbPixels) + ctl_left
            ctl_bottom = frm.ScaleX(ctl.Height, frm.ScaleMode, vbPixels) + ctl_top
            control_rgn = CreateRectRgn(ctl_left, ctl_top, ctl_right, ctl_bottom)
            CombineRgn combined_rgn, combined_rgn, control_rgn, RGN_OR
        End If
    Next ctl

    'Restrict the window to the region.
    SetWindowRgn frm.hwnd, combined_rgn, True
End Sub

Public Function RemoveCharacter(ByVal DataToTest As String, ByVal sChar As String) As String
Dim temp As Variant, x As Integer

    For x = 1 To Len(DataToTest)
        If Not Mid(DataToTest, x, 1) = sChar Then
            temp = temp & Mid(DataToTest, x, 1)
        End If
    Next x
    RemoveCharacter = temp

End Function

Public Function ClipNull(ByVal DataToClip As String, Optional ByVal nLen As Integer) As String
On Error GoTo error:
Dim x As Integer

If nLen = 0 Then nLen = Len(DataToClip)

For x = 1 To nLen
    If Mid(DataToClip, x, 1) = Chr(0) Then
        ClipNull = Mid(DataToClip, 1, x - 1)
        Exit Function
    End If
Next x
ClipNull = DataToClip

Exit Function
error:
Call HandleError
End Function

Public Function AddChr(ByVal DataToFinish As String, ByVal ChrToAdd As String, length As Integer) As String
    Dim i As Integer
    For i = 1 To length - Len(DataToFinish)
        DataToFinish = DataToFinish & ChrToAdd
    Next i
    AddChr = DataToFinish
End Function

Public Sub DBStatRowToStruct(row() As Byte)
RowToStruct row, DBStatFldMap, DBStat, LenB(DBStat)
End Sub

Public Sub DBStatStructToRow(row() As Byte)
StructToRow row, DBStatFldMap, DBStat, LenB(DBStat)
End Sub

'Public Sub GetNextExStructToRow(row() As Byte)
'StructToRow row, GetNextExFldMap, GetNextExRec, LenB(GetNextExRec)
'End Sub

'Public Sub GetNextExRowToStruct(row() As Byte)
'RowToStruct row, GetNextExFldMap, GetNextExRec, LenB(GetNextExRec)
'End Sub

'Public Sub RoomFilterStructToRow(row() As Byte)
'StructToRow row, RoomFilterFldMap, RoomFilterRec, LenB(RoomFilterRec)
'End Sub

'Public Sub RoomFilterRowToStruct(row() As Byte)
'RowToStruct row, RoomFilterFldMap, RoomFilterRec, LenB(RoomFilterRec)
'End Sub

Public Sub RoomRowToStruct(row() As Byte)
RowToStruct row, RoomFldMap, Roomrec, LenB(Roomrec)
End Sub

Public Sub RoomStructToRow(row() As Byte)
StructToRow row, RoomFldMap, Roomrec, LenB(Roomrec)
End Sub

Public Sub TextblockStructToRow(row() As Byte)
StructToRow row, TextblockFldMap, TextblockRec, LenB(TextblockRec)
End Sub

Public Sub TextblockRowToStruct(row() As Byte)
RowToStruct row, TextblockFldMap, TextblockRec, LenB(TextblockRec)
End Sub

Public Sub ShopStructToRow(row() As Byte)
StructToRow row, ShopFldMap, Shoprec, LenB(Shoprec)
End Sub

Public Sub ShopRowToStruct(row() As Byte)
RowToStruct row, ShopFldMap, Shoprec, LenB(Shoprec)
End Sub

Public Sub ActionStructToRow(row() As Byte)
StructToRow row, ActionFldMap, Actionrec, LenB(Actionrec)
End Sub

Public Sub ActionRowToStruct(row() As Byte)
RowToStruct row, ActionFldMap, Actionrec, LenB(Actionrec)
End Sub

Public Sub RaceStructToRow(row() As Byte)
StructToRow row, RaceFldMap, Racerec, LenB(Racerec)
End Sub

Public Sub RaceRowToStruct(row() As Byte)
RowToStruct row, RaceFldMap, Racerec, LenB(Racerec)
End Sub

Public Sub ClassStructToRow(row() As Byte)
StructToRow row, ClassFldMap, Classrec, LenB(Classrec)
End Sub

Public Sub ClassRowToStruct(row() As Byte)
RowToStruct row, ClassFldMap, Classrec, LenB(Classrec)
End Sub

Public Sub ItemStructToRow(row() As Byte)
StructToRow row, ItemFldMap, Itemrec, LenB(Itemrec)
End Sub

Public Sub ItemRowToStruct(row() As Byte)
RowToStruct row, ItemFldMap, Itemrec, LenB(Itemrec)
End Sub

Public Sub SpellStructToRow(row() As Byte)
StructToRow row, SpellFldMap, Spellrec, LenB(Spellrec)
End Sub

Public Sub SpellRowToStruct(row() As Byte)
RowToStruct row, SpellFldMap, Spellrec, LenB(Spellrec)
End Sub

Public Sub MessageStructToRow(row() As Byte)
StructToRow row, MessageFldMap, Messagerec, LenB(Messagerec)
End Sub

Public Sub MessageRowToStruct(row() As Byte)
RowToStruct row, MessageFldMap, Messagerec, LenB(Messagerec)
End Sub

Public Sub MonsterStructToRow(row() As Byte)
StructToRow row, MonsterFldMap, Monsterrec, LenB(Monsterrec)
End Sub

Public Sub MonsterRowToStruct(row() As Byte)
RowToStruct row, MonsterFldMap, Monsterrec, LenB(Monsterrec)
End Sub

Public Sub BankStructToRow(row() As Byte)
StructToRow row, BankFldMap, Bankrec, LenB(Bankrec)
End Sub

Public Sub BankRowToStruct(row() As Byte)
RowToStruct row, BankFldMap, Bankrec, LenB(Bankrec)
End Sub

Public Sub GangStructToRow(row() As Byte)
StructToRow row, GangFldMap, Gangrec, LenB(Gangrec)
End Sub

Public Sub GangRowToStruct(row() As Byte)
RowToStruct row, GangFldMap, Gangrec, LenB(Gangrec)
End Sub
Public Sub UserStructToRow(row() As Byte)
StructToRow row, UserFldMap, Userrec, LenB(Userrec)
End Sub

Public Sub UserRowToStruct(row() As Byte)
RowToStruct row, UserFldMap, Userrec, LenB(Userrec)
End Sub

Public Sub BankKeyStructToRow(row() As Byte)
StructToRow row, BankKeyFldMap, BankKey, LenB(BankKey)
End Sub

Public Sub BankKeyRowToStruct(row() As Byte)
RowToStruct row, BankKeyFldMap, BankKey, LenB(BankKey)
End Sub

'set and return here so we can avoid having to make a separate 'set' call before-hand.
Public Function TextblockKeyStructToRow() As TextblockKeyDataBufType
StructToRow TextblockKeyDataBuf.bytes, TextblockKeyFldMap, TextblockKey, LenB(TextblockKey)
TextblockKeyStructToRow = TextblockKeyDataBuf
End Function

Private Function SetDatVersion() As Boolean

eDatFileVersion = ReadINI("Settings", "eDatFileVersion" & IIf(WorksWithN = True, "_n", IIf(WorksWithWG = True, "_wg", "")))
Call SetDatSuffixStrings

If Not WorksWithWG Then
    Select Case eDatFileVersion '***NEWMUDVER***
        Case Is <= 6: 'v1.11h to v1.11n
            If Not WorksWithN Then
                MsgBox "This version of NMR only supports MajorMUD v1.11o and greater." & vbCrLf _
                    & "Please choose a valid dat file version from the settings window.", vbExclamation
                Exit Function
            End If
        Case Is >= 7: 'v1.11o+
            If WorksWithN Then
                MsgBox "This version of NMR only supports MajorMUD v1.11i through v1.11n." & vbCrLf _
                    & "Please choose a valid dat file version from the settings window.", vbExclamation
                Exit Function
            End If
        Case Else:
            MsgBox "'Dat File Version' setting invalid, please set it on the settings screen."
            Exit Function
    End Select
End If

'BankKey.nothing1 = &H0
'BankKey.nothing2 = &H0

If WorksWithWG Then
    TextblockKey.PartNum = 0
    TextblockKey.LeadIn(1) = &H63
    TextblockKey.LeadIn(2) = &H20
    TextblockKey.LeadIn(3) = &H63
    TextblockKey.LeadIn(4) = &H6C
    TextblockKey.LeadIn(5) = &H73
    TextblockKey.LeadIn(6) = &HD
    TextblockKey.LeadIn(7) = &H9
    TextblockKey.LeadIn(8) = &HD
    TextblockKey.LeadIn(9) = &H0
    TextblockKey.LeadIn(10) = &H0
    TextblockKey.LeadIn(11) = &H0
    TextblockKey.LeadIn(12) = &H0
Else
    Select Case eDatFileVersion '***NEWMUDVER***
        Case 2: 'J
            TextblockKey.PartNum = 0
            TextblockKey.LeadIn(1) = &H63
            TextblockKey.LeadIn(2) = &H20
            TextblockKey.LeadIn(3) = &H63
            TextblockKey.LeadIn(4) = &H6C
            TextblockKey.LeadIn(5) = &H0
            TextblockKey.LeadIn(6) = &H0
            TextblockKey.LeadIn(7) = &H73
            TextblockKey.LeadIn(8) = &HD
            TextblockKey.LeadIn(9) = &H98
            TextblockKey.LeadIn(10) = &H94
            TextblockKey.LeadIn(11) = &H40
            TextblockKey.LeadIn(12) = &H59
            TextblockKey.LeadIn(13) = &H35
            TextblockKey.LeadIn(14) = &H0
        Case 4: 'L
            TextblockKey.PartNum = 0
            TextblockKey.LeadIn(1) = &H63
            TextblockKey.LeadIn(2) = &H20
            TextblockKey.LeadIn(3) = &H63
            TextblockKey.LeadIn(4) = &H6C
            TextblockKey.LeadIn(5) = &H0
            TextblockKey.LeadIn(6) = &H0
            TextblockKey.LeadIn(7) = &H73
            TextblockKey.LeadIn(8) = &HD
            TextblockKey.LeadIn(9) = &H70
            TextblockKey.LeadIn(10) = &H72
            TextblockKey.LeadIn(11) = &H65
            TextblockKey.LeadIn(12) = &H6E
            TextblockKey.LeadIn(13) = &H35
            TextblockKey.LeadIn(14) = &H0
        Case 6: 'N
            TextblockKey.PartNum = 0
            TextblockKey.LeadIn(1) = &H63
            TextblockKey.LeadIn(2) = &H20
            TextblockKey.LeadIn(3) = &H63
            TextblockKey.LeadIn(4) = &H6C
            TextblockKey.LeadIn(5) = &H0
            TextblockKey.LeadIn(6) = &H0
            TextblockKey.LeadIn(7) = &H73
            TextblockKey.LeadIn(8) = &HD
            TextblockKey.LeadIn(9) = &H65
            TextblockKey.LeadIn(10) = &H63
            TextblockKey.LeadIn(11) = &H6F
            TextblockKey.LeadIn(12) = &H72
            TextblockKey.LeadIn(13) = &H3A
            TextblockKey.LeadIn(14) = &H0
        Case 7 To 8: 'O & Pb13
            TextblockKey.PartNum = 0
            TextblockKey.LeadIn(1) = &H63
            TextblockKey.LeadIn(2) = &H20
            TextblockKey.LeadIn(3) = &H63
            TextblockKey.LeadIn(4) = &H6C
            TextblockKey.LeadIn(5) = &H0
            TextblockKey.LeadIn(6) = &H0
            TextblockKey.LeadIn(7) = &H73
            TextblockKey.LeadIn(8) = &HD
            TextblockKey.LeadIn(9) = &H1
            TextblockKey.LeadIn(10) = &HB1
            TextblockKey.LeadIn(11) = &H2
            TextblockKey.LeadIn(12) = &H59
            TextblockKey.LeadIn(13) = &H3B
            TextblockKey.LeadIn(14) = &H0
        Case 9: 'p final
            TextblockKey.PartNum = 0
            TextblockKey.LeadIn(1) = &H63
            TextblockKey.LeadIn(2) = &H20
            TextblockKey.LeadIn(3) = &H63
            TextblockKey.LeadIn(4) = &H6C
            TextblockKey.LeadIn(5) = &H0
            TextblockKey.LeadIn(6) = &H0
            TextblockKey.LeadIn(7) = &H73
            TextblockKey.LeadIn(8) = &HD
            TextblockKey.LeadIn(9) = &H9
            TextblockKey.LeadIn(10) = &HD
            TextblockKey.LeadIn(11) = &H0
            TextblockKey.LeadIn(12) = &H0
            TextblockKey.LeadIn(13) = &H0
            TextblockKey.LeadIn(14) = &H0
        Case Else: 'h, i, k, m
            TextblockKey.PartNum = 0
            TextblockKey.LeadIn(1) = &H63
            TextblockKey.LeadIn(2) = &H20
            TextblockKey.LeadIn(3) = &H63
            TextblockKey.LeadIn(4) = &H6C
            TextblockKey.LeadIn(5) = &H0
            TextblockKey.LeadIn(6) = &H0
            TextblockKey.LeadIn(7) = &H73
            TextblockKey.LeadIn(8) = &HD
            TextblockKey.LeadIn(9) = &H0
            TextblockKey.LeadIn(10) = &H0
            TextblockKey.LeadIn(11) = &H0
            TextblockKey.LeadIn(12) = &H0
            TextblockKey.LeadIn(13) = &H0
            TextblockKey.LeadIn(14) = &H0
    End Select
End If

SetDatVersion = True
End Function

Private Function CheckDatVersion() As Boolean
On Error GoTo error:
Dim nStatus As Integer, sVer As String, bMatch As Boolean, bMonCheck As Boolean


nStatus = BTRCALL(BGETFIRST, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Warning, unable to verify dat file version setting.  This is done by accessing the first textblock." & vbCrLf _
        & "If your textblock database is empty or missing, that would cause this error.", vbExclamation
    Exit Function
End If
TextblockRowToStruct TextblockDataBuf.buf

CheckDatVersion = True

'Dim x As Integer
'For x = 1 To 14
'    Debug.Print "&H" & Hex(TextblockRec.LeadIn(x)) & vbCrLf
'Next x

Select Case TextblockRec.LeadIn(10) 'checks the first textblock ***NEWMUDVER***
    Case &H94: sVer = "v1.11j"
    Case &H72: sVer = "v1.11L"
    Case &H63: sVer = "v1.11n"
    Case &HB1: sVer = "v1.11o (or v1.11p13)"
    Case &HD: sVer = "v1.11p"
    Case &H0:
        If (TextblockRec.LeadIn(9) = &H0) And (TextblockRec.LeadIn(7) = &H73) Then
            sVer = "v1.11h, v1.11i, v1.11k, or v1.11m"
        ElseIf (TextblockRec.LeadIn(9) = &H0) And (TextblockRec.LeadIn(7) = &H9) Then
            sVer = "v1.11p-WG"
        Else
            sVer = "v?.??"
        End If
    Case Else: sVer = "v?.??"
End Select

bMatch = True
If WorksWithWG Then
    If TextblockRec.LeadIn(12) <> &H0 Then bMatch = False
    If TextblockRec.LeadIn(10) <> &H0 Then bMatch = False
    If TextblockRec.LeadIn(8) <> &HD Then bMatch = False
    If TextblockRec.LeadIn(7) <> &H9 Then bMatch = False
Else
    Select Case eDatFileVersion '***NEWMUDVER***
        Case 2: 'J
            If TextblockRec.LeadIn(12) <> &H59 Then bMatch = False
            If TextblockRec.LeadIn(13) <> &H35 Then bMatch = False
            If TextblockRec.LeadIn(10) <> &H94 Then bMatch = False
        Case 4: 'L
            If TextblockRec.LeadIn(12) <> &H6E Then bMatch = False
            If TextblockRec.LeadIn(10) <> &H72 Then bMatch = False
        Case 6: 'N
            If TextblockRec.LeadIn(12) <> &H72 Then bMatch = False
            If TextblockRec.LeadIn(10) <> &H63 Then bMatch = False
        Case 7 To 8: 'o and p
            If TextblockRec.LeadIn(12) <> &H59 Then bMatch = False
            If TextblockRec.LeadIn(13) <> &H3B Then bMatch = False
            If TextblockRec.LeadIn(10) <> &HB1 Then bMatch = False
        Case 9: 'p final
            If TextblockRec.LeadIn(8) <> &HD Then bMatch = False
            If TextblockRec.LeadIn(9) <> &H9 Then bMatch = False
            If TextblockRec.LeadIn(10) <> &HD Then bMatch = False
        Case Else: 'h, i, k, m
            If TextblockRec.LeadIn(12) <> &H0 Then bMatch = False
            If TextblockRec.LeadIn(10) <> &H0 Then bMatch = False
            If TextblockRec.LeadIn(8) <> &HD Then bMatch = False
            If TextblockRec.LeadIn(7) <> &H73 Then bMatch = False
    End Select
End If

bMonCheck = True

If eDatFileVersion >= v111j Then
    If bMatch = True Then 'if textblocks matched, also do a monster check for 0 in the exp multi
        nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            MonsterRowToStruct Monsterdatabuf.buf
            If Monsterrec.ExpMulti = 0 Then bMonCheck = False
        End If
    End If
End If

'check room database size
nStatus = BTRCALL(BSTAT, RoomPosBlock, DBStatDatabuf, Len(DBStatDatabuf), 0, KEY_BUF_LEN, 0)
If nStatus = 0 Then
    DBStatRowToStruct DBStatDatabuf.buf
    If DBStat.RecLen = 1512 Then
        If Not eDatFileVersion <= v111n Then bMatch = False
    ElseIf DBStat.RecLen = 1544 Then
        If Not eDatFileVersion >= v111o Then bMatch = False
    ElseIf DBStat.RecLen = 1528 Then
        If Not WorksWithWG Then bMatch = False
    Else
        CheckDatVersion = False
        MsgBox "Warning: Invalid room database record size returned while verifying dat file version setting.", vbExclamation
    End If
Else
    CheckDatVersion = False
    MsgBox "Warning: Unable to check room record size while verifying dat file version setting.", vbExclamation
End If

If bMatch = False Then
    MsgBox "Warning, current 'Dat File Version' setting does not seem to match the loaded dat files!" & vbCrLf _
        & "These dats appear to be " & sVer & "." & vbCrLf _
        & " " & vbCrLf _
        & "Please check the option on the settings screen.", vbExclamation
    CheckDatVersion = False
    'If nStatus = vbYes Then Load frmSettings
    frmMain.stsStatusBar.Panels(3).Text = frmMain.stsStatusBar.Panels(3).Text & " <-- ** appears invalid! **"
Else
    If bMonCheck = False Then
        MsgBox "Warning: Current 'Dat File Version' setting does not seem to match the loaded dat files!" & vbCrLf _
        & "These dats appear to be older than v1.11J." & vbCrLf & vbCrLf _
        & "(This message is caused by having set the first monster to have an exp multiplier value of zero)" & vbCrLf _
        & "Please check the option on the settings screen.", vbExclamation
    CheckDatVersion = False
    'If nStatus = vbYes Then Load frmSettings
    frmMain.stsStatusBar.Panels(3).Text = frmMain.stsStatusBar.Panels(3).Text & " <-- ** appears invalid! **"
    End If
End If

Exit Function
error:
Call HandleError
End Function

Public Function FriendlyDatVersion(ByVal nNum As Integer) As String
If WorksWithWG Then
    FriendlyDatVersion = "v1.11p-WG"
Else
    Select Case nNum '***NEWMUDVER***
        Case 0: FriendlyDatVersion = "v1.11h"
        Case 1: FriendlyDatVersion = "v1.11i"
        Case 2: FriendlyDatVersion = "v1.11j"
        Case 3: FriendlyDatVersion = "v1.11k"
        Case 4: FriendlyDatVersion = "v1.11L"
        Case 5: FriendlyDatVersion = "v1.11m"
        Case 6: FriendlyDatVersion = "v1.11n"
        Case 7: FriendlyDatVersion = "v1.11o"
        Case 8: FriendlyDatVersion = "v1.11p-b13"
        Case 9: FriendlyDatVersion = "v1.11p"
        Case Else:  FriendlyDatVersion = "v???"
    End Select
End If
End Function

Public Sub InitTaskbar()
Dim nTemp As Integer

On Error GoTo error:

frmMain.tbTaskBar.Visible = False
nTemp = Val(ReadINI("Settings", "TaskBarPos"))
If nTemp = 0 Then
    frmMain.tbTaskBar.Align = 2 'bottom
    frmMain.tbTaskBar.Visible = True
ElseIf nTemp = 1 Then
    frmMain.tbTaskBar.Align = 1 'top
    frmMain.tbTaskBar.Visible = True
End If
nTemp = Val(ReadINI("Settings", "TaskBarDelay"))
If nTemp > 0 Then
    frmMain.tbTaskBar.AutoHideWait = nTemp * 1000
Else
    frmMain.tbTaskBar.AutoHideWait = 1000
End If
If Val(ReadINI("Settings", "TaskBarAutoHide")) > 0 Then
    frmMain.tbTaskBar.AutoHide = True
Else
    frmMain.tbTaskBar.AutoHide = False
End If

out:
Exit Sub
error:
Call HandleError("InitTaskbar")
Resume out:

End Sub

Private Sub SetDatSuffixStrings()
    If WorksWithWG Then
        strDatSuffix_ACTS = strDatSuffixNNT_ACTS
        strDatSuffix_BANKS = strDatSuffixNNT_BANKS
        strDatSuffix_CLASS = strDatSuffixNNT_CLASS
        strDatSuffix_GANGS = strDatSuffixNNT_GANGS
        strDatSuffix_ITEMS = strDatSuffixNNT_ITEMS
        strDatSuffix_KNMSR = strDatSuffixNNT_KNMSR
        strDatSuffix_MP = strDatSuffixNNT_MP
        strDatSuffix_MSG = strDatSuffixNNT_MSG
        strDatSuffix_RACE = strDatSuffixNNT_RACE
        strDatSuffix_SHOPS = strDatSuffixNNT_SHOPS
        strDatSuffix_SPELS = strDatSuffixNNT_SPELS
        strDatSuffix_TEXT = strDatSuffixNNT_TEXT
        strDatSuffix_UPDAT = strDatSuffixNNT_UPDAT
        strDatSuffix_USERS = strDatSuffixNNT_USERS
    Else
        strDatSuffix_ACTS = strDatSuffixNT_ACTS
        strDatSuffix_BANKS = strDatSuffixNT_BANKS
        strDatSuffix_CLASS = strDatSuffixNT_CLASS
        strDatSuffix_GANGS = strDatSuffixNT_GANGS
        strDatSuffix_ITEMS = strDatSuffixNT_ITEMS
        strDatSuffix_KNMSR = strDatSuffixNT_KNMSR
        strDatSuffix_MP = strDatSuffixNT_MP
        strDatSuffix_MSG = strDatSuffixNT_MSG
        strDatSuffix_RACE = strDatSuffixNT_RACE
        strDatSuffix_SHOPS = strDatSuffixNT_SHOPS
        strDatSuffix_SPELS = strDatSuffixNT_SPELS
        strDatSuffix_TEXT = strDatSuffixNT_TEXT
        strDatSuffix_UPDAT = strDatSuffixNT_UPDAT
        strDatSuffix_USERS = strDatSuffixNT_USERS
    End If
End Sub

Public Sub Startup()
On Error GoTo error:
Dim nYesNo As Integer, WGPath As String, nStatus As Integer, bTest As Boolean

bKeepSettingsOpen = True

frmSplash.lblStatus.Caption = "Loading ..."
DoEvents

'initialize the mgil array (monster group/index list)
Erase MGIL()
ReDim MGIL(1, 1)

Call GetTitleBarOffset
Call CheckINIReadOnly
Call InitTaskbar
nStatus = 0

If Val(ReadINI("Settings", "OppositeListOrder")) > 0 Then
    bOppositeListOrder = True
Else
    bOppositeListOrder = False
End If

If Val(ReadINI("Settings", "OnlyLoadNames")) > 0 Then
    bOnlyNames = True
Else
    bOnlyNames = False
End If

frmSplash.lblStatus.Caption = "Opening Ability DB ..."
Call OpenAbilityDB

strDatCallLetters = ReadINI("Settings", "DatCallLetters" & IIf(WorksWithN = True, "_n", IIf(WorksWithWG = True, "_wg", "")))
Call SetDatSuffixStrings

If Val(ReadINI("Settings", "UseCPU")) > 0 Then
    bUseCPU = True
Else
    bUseCPU = False
End If

frmSplash.lblStatus.Caption = "Initializing Field Maps ..."
DoEvents
Call IntFieldMaps

If ReadINI("Settings", "FirstRun" & IIf(WorksWithN = True, "_n", IIf(WorksWithWG = True, "_wg", ""))) = "0" Then
    eDatFileVersion = IIf(WorksWithN = True, 6, IIf(WorksWithWG = True, 0, 9))
    Call SetDatSuffixStrings
    MsgBox "This appears to be your first time launching" & _
        " Nightmare Redux" & IIf(WorksWithN = True, " for vN.", ".") & _
        " Please set the path and version of your MajorMUD *.dat files on the following" & _
        " settings screen.", vbInformation
    GoTo load_settings:
End If

bTest = SetDatVersion
If bTest = False Then GoTo load_settings:

frmSplash.lblStatus.Caption = "Opening Files ..."
DoEvents
WGPath = ReadINI("Settings", "WGPath" & IIf(WorksWithN = True, "_n", IIf(WorksWithWG = True, "_wg", "")))
If Not Right(WGPath, 1) = "\" Then
    WGPath = WGPath & "\"
    Call WriteINI("Settings", "WGPath" & IIf(WorksWithN = True, "_n", IIf(WorksWithWG = True, "_wg", "")), WGPath)
End If
If DirectoryExists(WGPath) = False Then
    MsgBox "Nightmare cannot find your MajorMUD *.dat files.  Please choose the location of them on the settings screen.", vbExclamation
    GoTo load_settings:
End If
If FileExists(WGPath & "w" & strDatCallLetters & strDatSuffix_RACE) = False Or _
    FileExists(WGPath & "w" & strDatCallLetters & strDatSuffix_CLASS) = False Then
    MsgBox "w" & strDatCallLetters & strDatSuffix_RACE & " or w" & strDatCallLetters & strDatSuffix_CLASS & " was not found." _
    & "Please locate your dat files on the settings screen.", vbExclamation
    GoTo load_settings:
End If

RaceKeyBuffer = WGPath & "w" & strDatCallLetters & strDatSuffix_RACE & Chr(0)
ClassKeyBuffer = WGPath & "w" & strDatCallLetters & strDatSuffix_CLASS & Chr(0)
SpellKeyBuffer = WGPath & "w" & strDatCallLetters & strDatSuffix_SPELS & Chr(0)
MonsterKeyBuffer = WGPath & "w" & strDatCallLetters & strDatSuffix_KNMSR & Chr(0)
ItemKeyBuffer = WGPath & "w" & strDatCallLetters & strDatSuffix_ITEMS & Chr(0)
ShopKeyBuffer = WGPath & "w" & strDatCallLetters & strDatSuffix_SHOPS & Chr(0)
RoomKeyBuffer = WGPath & "w" & strDatCallLetters & strDatSuffix_MP & Chr(0)
MessageKeyBuffer = WGPath & "w" & strDatCallLetters & strDatSuffix_MSG & Chr(0)
TextblockKeyBuffer = WGPath & "w" & strDatCallLetters & strDatSuffix_TEXT & Chr(0)
UserKeyBuffer = WGPath & "w" & strDatCallLetters & strDatSuffix_USERS & Chr(0)
ActionKeyBuffer = WGPath & "w" & strDatCallLetters & strDatSuffix_ACTS & Chr(0)
BankKeyBuffer = WGPath & "w" & strDatCallLetters & strDatSuffix_BANKS & Chr(0)
GangKeyBuffer = WGPath & "w" & strDatCallLetters & strDatSuffix_GANGS & Chr(0)

frmMain.stsStatusBar.Panels(2).Text = "call letters: " & strDatCallLetters
If bUseCPU = True Then
    frmMain.stsStatusBar.Panels(4).Text = "Use CPU: On"
Else
    frmMain.stsStatusBar.Panels(4).Text = "Use CPU: Off"
End If
frmMain.stsStatusBar.Panels(5).Text = WGPath
'frmLogo.lblWGPath.Caption = WGPath

'frmLogo.lblVersion = FriendlyDatVersion(eDatFileVersion)
frmMain.stsStatusBar.Panels(3).Text = "file version: " & FriendlyDatVersion(eDatFileVersion)

frmSplash.lblStatus.Caption = "Opening Races ..."
DoEvents
nStatus = BTRCALL(BOPEN, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0) ', -64)
If Not nStatus = 0 Then
    nYesNo = MsgBox("Error opening Race file -- Btrieve Error: " & vbCrLf & RTrim(RemoveCharacter(RaceKeyBuffer, vbNullChar)) _
        & vbCrLf & BtrieveErrorCode(nStatus) & vbCrLf & vbCrLf & "Continue Loading?", vbYesNo + vbDefaultButton2 + vbCritical)
    If Not nYesNo = vbYes Then GoTo done:
Else
    Call LoadRaceArray
End If

frmSplash.lblStatus.Caption = "Opening Classes ..."
nStatus = BTRCALL(BOPEN, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0) ', -64)
If Not nStatus = 0 Then
    nYesNo = MsgBox("Error opening Class file -- Btrieve Error: " & vbCrLf & RTrim(RemoveCharacter(ClassKeyBuffer, vbNullChar)) _
        & vbCrLf & BtrieveErrorCode(nStatus) & vbCrLf & vbCrLf & "Continue Loading?", vbYesNo + vbDefaultButton2 + vbCritical)
    If Not nYesNo = vbYes Then GoTo done:
Else
    Call LoadClassArray
End If

frmSplash.lblStatus.Caption = "Opening Spells ..."
DoEvents
nStatus = BTRCALL(BOPEN, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0) ', -64)
If Not nStatus = 0 Then
    nYesNo = MsgBox("Error opening Spell file -- Btrieve Error: " & vbCrLf & RTrim(RemoveCharacter(SpellKeyBuffer, vbNullChar)) _
        & vbCrLf & BtrieveErrorCode(nStatus) & vbCrLf & vbCrLf & "Continue Loading?", vbYesNo + vbDefaultButton2 + vbCritical)
    If Not nYesNo = vbYes Then GoTo done:
End If

frmSplash.lblStatus.Caption = "Opening Monsters ..."
DoEvents
nStatus = BTRCALL(BOPEN, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0) ', -64)
If Not nStatus = 0 Then
    nYesNo = MsgBox("Error opening Monster file -- Btrieve Error: " & vbCrLf & RTrim(RemoveCharacter(MonsterKeyBuffer, vbNullChar)) _
        & vbCrLf & BtrieveErrorCode(nStatus) & vbCrLf & vbCrLf & "Continue Loading?", vbYesNo + vbDefaultButton2 + vbCritical)
    If Not nYesNo = vbYes Then GoTo done:
End If

frmSplash.lblStatus.Caption = "Opening Items ..."
DoEvents
nStatus = BTRCALL(BOPEN, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0) ', -64)
If Not nStatus = 0 Then
    nYesNo = MsgBox("Error opening Item file -- Btrieve Error: " & vbCrLf & RTrim(RemoveCharacter(ItemKeyBuffer, vbNullChar)) _
        & vbCrLf & BtrieveErrorCode(nStatus) & vbCrLf & vbCrLf & "Continue Loading?", vbYesNo + vbDefaultButton2 + vbCritical)
    If Not nYesNo = vbYes Then GoTo done:
End If

frmSplash.lblStatus.Caption = "Opening Shops ..."
DoEvents
nStatus = BTRCALL(BOPEN, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0) ', -64)
If Not nStatus = 0 Then
    nYesNo = MsgBox("Error opening Shop file -- Btrieve Error: " & vbCrLf & RTrim(RemoveCharacter(ShopKeyBuffer, vbNullChar)) _
        & vbCrLf & BtrieveErrorCode(nStatus) & vbCrLf & vbCrLf & "Continue Loading?", vbYesNo + vbDefaultButton2 + vbCritical)
    If Not nYesNo = vbYes Then GoTo done:
End If

frmSplash.lblStatus.Caption = "Opening Rooms ..."
DoEvents
nStatus = BTRCALL(BOPEN, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0) ', -64)
If Not nStatus = 0 Then
    nYesNo = MsgBox("Error opening Room file -- Btrieve Error: " & vbCrLf & RTrim(RemoveCharacter(RoomKeyBuffer, vbNullChar)) _
        & vbCrLf & BtrieveErrorCode(nStatus) & vbCrLf & vbCrLf & "Continue Loading?", vbYesNo + vbDefaultButton2 + vbCritical)
    If Not nYesNo = vbYes Then GoTo done:
End If

frmSplash.lblStatus.Caption = "Opening Messages ..."
DoEvents
nStatus = BTRCALL(BOPEN, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0) ', -64)
If Not nStatus = 0 Then
    nYesNo = MsgBox("Error opening Message file -- Btrieve Error: " & vbCrLf & RTrim(RemoveCharacter(MessageKeyBuffer, vbNullChar)) _
        & vbCrLf & BtrieveErrorCode(nStatus) & vbCrLf & vbCrLf & "Continue Loading?", vbYesNo + vbDefaultButton2 + vbCritical)
    If Not nYesNo = vbYes Then GoTo done:
End If

frmSplash.lblStatus.Caption = "Opening Textblocks ..."
DoEvents
nStatus = BTRCALL(BOPEN, TextblockPosBlock, TextblockDataBuf, Len(TextblockDataBuf), ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0) ', -64)
If Not nStatus = 0 Then
    nYesNo = MsgBox("Error opening Textblock file -- Btrieve Error: " & vbCrLf & RTrim(RemoveCharacter(TextblockKeyBuffer, vbNullChar)) _
        & vbCrLf & BtrieveErrorCode(nStatus) & vbCrLf & vbCrLf & "Continue Loading?", vbYesNo + vbDefaultButton2 + vbCritical)
    If Not nYesNo = vbYes Then GoTo done:
End If

frmSplash.lblStatus.Caption = "Opening Actions ..."
DoEvents
nStatus = BTRCALL(BOPEN, ActionPosBlock, ActionDatabuf, Len(ActionDatabuf), ByVal ActionKeyBuffer, KEY_BUF_LEN, 0) ', -64)
If Not nStatus = 0 Then
    nYesNo = MsgBox("Error opening Action file -- Btrieve Error: " & vbCrLf & RTrim(RemoveCharacter(ActionKeyBuffer, vbNullChar)) _
        & vbCrLf & BtrieveErrorCode(nStatus) & vbCrLf & vbCrLf & "Continue Loading?", vbYesNo + vbDefaultButton2 + vbCritical)
    If Not nYesNo = vbYes Then GoTo done:
End If

frmSplash.lblStatus.Caption = "Opening Users ..."
DoEvents
nStatus = BTRCALL(BOPEN, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0) ', -64)
If Not nStatus = 0 Then
    nYesNo = MsgBox("Error opening User file -- Btrieve Error: " & vbCrLf & RTrim(RemoveCharacter(UserKeyBuffer, vbNullChar)) _
        & vbCrLf & BtrieveErrorCode(nStatus) & vbCrLf & vbCrLf & "Continue Loading?", vbYesNo + vbDefaultButton2 + vbCritical)
    If Not nYesNo = vbYes Then GoTo done:
End If

frmSplash.lblStatus.Caption = "Opening Bankbooks ..."
DoEvents
nStatus = BTRCALL(BOPEN, BankPosBlock, BankDatabuf, Len(BankDatabuf), ByVal BankKeyBuffer, KEY_BUF_LEN, 0) ', -64)
If Not nStatus = 0 Then
    nYesNo = MsgBox("Error opening Bank file -- Btrieve Error: " & vbCrLf & RTrim(RemoveCharacter(BankKeyBuffer, vbNullChar)) _
        & vbCrLf & BtrieveErrorCode(nStatus) & vbCrLf & vbCrLf & "Continue Loading?", vbYesNo + vbDefaultButton2 + vbCritical)
    If Not nYesNo = vbYes Then GoTo done:
End If

frmSplash.lblStatus.Caption = "Opening Gangs ..."
DoEvents
nStatus = BTRCALL(BOPEN, GangPosBlock, GangDatabuf, Len(GangDatabuf), ByVal GangKeyBuffer, KEY_BUF_LEN, 0) ', -64)
If Not nStatus = 0 Then
    nYesNo = MsgBox("Error opening Gang file -- Btrieve Error: " & vbCrLf & RTrim(RemoveCharacter(GangKeyBuffer, vbNullChar)) _
        & vbCrLf & BtrieveErrorCode(nStatus) & vbCrLf & vbCrLf & "Continue Loading?", vbYesNo + vbDefaultButton2 + vbCritical)
    If Not nYesNo = vbYes Then GoTo done:
End If

bTest = CheckDatVersion
If bTest = False Then GoTo load_settings:

If ReadINI("Settings", "AutoMonsterIndex") = "1" Then
    frmSplash.lblStatus.Caption = "Building Monster Index ..."
    DoEvents
    Call CreateMGIL
End If


done:
bKeepSettingsOpen = False
DoEvents
Exit Sub

load_settings:
Load frmSettings
frmSettings.bReload = True
DoEvents
Exit Sub

error:
bKeepSettingsOpen = False
Call HandleError("Startup")
End Sub

Public Sub LoadRaceArray()
Dim nStatus As Integer
Static counter As Long
On Error GoTo error:

counter = 1
ReDim Races(0)
Races(0).Name = "None"

nStatus = BTRCALL(BGETFIRST, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "LoadRaceArray, BGETFIRST, Race, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

Do While nStatus = 0
    RaceRowToStruct Racedatabuf.buf
    If Racerec.Number > counter Then
        Do While Racerec.Number > counter
            ReDim Preserve Races(counter)
            Races(counter).Number = counter
            Races(counter).Name = "Race #" & CStr(counter)
            counter = counter + 1
        Loop
    End If
    ReDim Preserve Races(Racerec.Number)
    Races(Racerec.Number).Number = Racerec.Number
    Races(Racerec.Number).Name = Racerec.Name
    nStatus = BTRCALL(BGETNEXT, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
    counter = counter + 1
Loop

out:
Exit Sub
error:
Call HandleError("LoadRaceArray")
Resume out:

End Sub

Public Sub LoadClassArray()
Dim nStatus As Integer
Static counter As Long
On Error GoTo error:

counter = 1
ReDim Classes(0)
Classes(0).Name = "None"

nStatus = BTRCALL(BGETFIRST, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "LoadClassArray, BGETFIRST, Class, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

Do While nStatus = 0
    ClassRowToStruct Classdatabuf.buf
    If Classrec.Number > counter Then
        Do While Classrec.Number > counter
            ReDim Preserve Classes(counter)
            Classes(counter).Number = counter
            Classes(counter).Name = "Class #" & CStr(counter)
            counter = counter + 1
        Loop
    End If
    ReDim Preserve Classes(Classrec.Number)
    Classes(Classrec.Number).Number = Classrec.Number
    Classes(Classrec.Number).Name = Classrec.Name
    nStatus = BTRCALL(BGETNEXT, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
    counter = counter + 1
Loop

out:
Exit Sub
error:
Call HandleError("LoadClassArray")
Resume out:

End Sub

Public Sub Add2ClassArray(ByVal nClass As Integer)
Dim nStatus As Integer
Static counter As Long
On Error GoTo error:

counter = 1
ReDim Classes(0)
Classes(0).Name = "None"

nStatus = BTRCALL(BGETFIRST, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "LoadClassArray, BGETFIRST, Class, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

Do While nStatus = 0
    ClassRowToStruct Classdatabuf.buf
    If Classrec.Number > counter Then
        Do While Classrec.Number > counter
            ReDim Preserve Classes(counter)
            Classes(counter).Number = counter
            Classes(counter).Name = "Class #" & CStr(counter)
            counter = counter + 1
        Loop
    End If
    ReDim Preserve Classes(Classrec.Number)
    Classes(Classrec.Number).Number = Classrec.Number
    Classes(Classrec.Number).Name = Classrec.Name
    nStatus = BTRCALL(BGETNEXT, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
    counter = counter + 1
Loop

If counter <= nClass Then
    Do While counter <= nClass
        ReDim Preserve Classes(counter)
        If counter <> nClass Then
            Classes(counter).Number = counter
            Classes(counter).Name = "Class #" & CStr(counter)
        Else
            Classes(nClass).Number = nClass
            Classes(nClass).Name = "Class #" & CStr(nClass)
        End If
        counter = counter + 1
    Loop
End If

out:
Exit Sub
error:
Call HandleError("Add2ClassArray")
Resume out:

End Sub

Public Sub Add2RaceArray(ByVal nRace As Integer)
Dim nStatus As Integer
Static counter As Long
On Error GoTo error:

counter = 1
ReDim Races(0)
Races(0).Name = "None"

nStatus = BTRCALL(BGETFIRST, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "LoadRaceArray, BGETFIRST, Race, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

Do While nStatus = 0
    RaceRowToStruct Racedatabuf.buf
    If Racerec.Number > counter Then
        Do While Racerec.Number > counter
            ReDim Preserve Races(counter)
            Races(counter).Number = counter
            Races(counter).Name = "Race #" & CStr(counter)
            counter = counter + 1
        Loop
    End If
    ReDim Preserve Races(Racerec.Number)
    Races(Racerec.Number).Number = Racerec.Number
    Races(Racerec.Number).Name = Racerec.Name
    nStatus = BTRCALL(BGETNEXT, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
    counter = counter + 1
Loop

If counter <= nRace Then
    Do While counter <= nRace
        ReDim Preserve Races(counter)
        If counter <> nRace Then
            Races(counter).Number = counter
            Races(counter).Name = "Race #" & CStr(counter)
        Else
            Races(nRace).Number = nRace
            Races(nRace).Name = "Race #" & CStr(nRace)
        End If
        counter = counter + 1
    Loop
End If

out:
Exit Sub
error:
Call HandleError("Add2RaceArray")
Resume out:

End Sub

Public Function ULong2SLong(ByVal Value As Variant) As Long
On Error GoTo error:
    If Value >= LongOffset Then Value = LongOffset - 1
    If Value < (0 - MaxLong) - 1 Then Value = (0 - MaxLong) - 1
    If Value <= MaxLong Then
        ULong2SLong = Value
    Else
        ULong2SLong = Value - LongOffset
    End If
Exit Function
error:
Call HandleError
End Function

Public Function SLong2ULong(ByVal Value As Variant) As Double
On Error GoTo error:
    If Value > MaxLong Then Value = MaxLong
    If Value < 0 Then
        SLong2ULong = Value + LongOffset
    Else
        SLong2ULong = Value
    End If
Exit Function
error:
Call HandleError
End Function

Public Function UInt2SInt(ByVal Value As Variant) As Integer
On Error GoTo error:
    If Value >= IntOffset Then Value = IntOffset - 1
    If Value < (0 - MaxInt) - 1 Then Value = (0 - MaxInt) - 1
    If Value <= MaxInt Then
        UInt2SInt = Value
    Else
        UInt2SInt = Value - IntOffset
    End If
Exit Function
error:
Call HandleError
End Function

Public Function SInt2UInt(ByVal Value As Variant) As Long
On Error GoTo error:
    If Value > MaxInt Then Value = MaxInt
    If Value < 0 Then
        SInt2UInt = Value + IntOffset
    Else
        SInt2UInt = Value
    End If
Exit Function
error:
Call HandleError
End Function

Public Function Dec2Bin(ByVal mynum As Variant) As String
Dim loopcounter As Integer

If mynum >= 2 ^ 31 Then
    Dec2Bin = "Number too big"
    Exit Function
End If

Do
    If (mynum And 2 ^ loopcounter) = 2 ^ loopcounter Then
        Dec2Bin = "1" & Dec2Bin
    Else
        Dec2Bin = "0" & Dec2Bin
    End If

    loopcounter = loopcounter + 1

Loop Until 2 ^ loopcounter > mynum

End Function

Public Function Bin2Dec(ByVal binValue As String) As Long
Dim lngValue As Long
Dim x As Long
Dim k As Long

k = Len(binValue) ' will only work with 32 or fewer "bits"
For x = k To 1 Step -1 ' work backwards down string
  If Mid$(binValue, x, 1) = "1" Then
    If k - x > 30 Then ' bit 31 is the sign bit
      lngValue = lngValue Or -2147483648# ' avoid overflow error
    Else
      lngValue = lngValue + 2 ^ (k - x)
    End If
  End If
Next x
Bin2Dec = lngValue
End Function

Public Function DOSDate2Date(ByVal Value As Long) As String
Dim BinaryDate As String, Month As String, Day As String, Year As String
Dim temp As String

If Value = 0 Then GoTo NoDate:

BinaryDate = Dec2Bin(Value)

If Len(BinaryDate) < 16 Then
    temp = 16 - Len(BinaryDate)
    temp = String(Val(temp), "0")
    BinaryDate = temp & BinaryDate
End If

Year = 1980 + Bin2Dec(Left(BinaryDate, 7))
If Year = "1980" Then Year = "0000"
Day = Bin2Dec(Right(BinaryDate, 5))
Month = Bin2Dec(Right(Left(BinaryDate, 11), 4))

If Len(Year) < 4 Then Year = String(4 - Len(Year), "0") & Year
If Len(Day) < 2 Then Day = String(2 - Len(Day), "0") & Day
If Len(Month) < 2 Then Month = String(2 - Len(Month), "0") & Month

DOSDate2Date = Month & "/" & Day & "/" & Year

Exit Function
NoDate:
DOSDate2Date = "00/00/0000"

End Function

Public Function DOSTime2Time(ByVal Value As Long) As String
Dim BinaryTime As String, Hour As String, Minute As String, Second As String
Dim temp As String

If Value = 0 Then GoTo NoTime:

BinaryTime = Dec2Bin(Value)

If Len(BinaryTime) < 16 Then
    temp = 16 - Len(BinaryTime)
    temp = String(Val(temp), "0")
    BinaryTime = temp & BinaryTime
End If

Hour = Bin2Dec(Left(BinaryTime, 5))
Minute = Bin2Dec(Right(Left(BinaryTime, 11), 6))
Second = 2 * Bin2Dec(Right(BinaryTime, 5))

If Len(Hour) < 2 Then Hour = String(2 - Len(Hour), "0") & Hour
If Len(Minute) < 2 Then Minute = String(2 - Len(Minute), "0") & Minute
If Len(Second) < 2 Then Second = String(2 - Len(Second), "0") & Second

DOSTime2Time = Hour & ":" & Minute & ":" & Second
Exit Function
NoTime:
DOSTime2Time = "00:00:00"
End Function

Public Function Time2DOSTime(ByVal Value As String) As Long
On Error GoTo error:
Dim BinaryTime As String, Hour As String, Minute As String, Second As String
Dim temp As String

If Value = "00:00:00" Then GoTo NoTime:
If Len(Value) < 8 Or Len(Value) > 8 Then GoTo BadFormat:

Hour = Dec2Bin(Val(Left(Value, 2)))
Minute = Dec2Bin(Val(Right(Left(Value, 5), 2)))
Second = Dec2Bin((Val(Right(Value, 2)) / 2))

If Len(Hour) < 5 Then Hour = String(5 - Len(Hour), "0") & Hour
If Len(Minute) < 6 Then Minute = String(6 - Len(Minute), "0") & Minute
If Len(Second) < 5 Then Second = String(5 - Len(Second), "0") & Second

If Len(Hour) > 5 Then Hour = "11111"
If Len(Minute) > 6 Then Minute = "111111"
If Len(Second) > 5 Then Second = "11111"

Time2DOSTime = Bin2Dec(Hour & Minute & Second)
Exit Function

NoTime:
Time2DOSTime = 0
Exit Function

BadFormat:
MsgBox "Incorrect time format!", vbExclamation
Time2DOSTime = -1
Exit Function

error:
MsgBox "Incorrect time format!", vbExclamation
Call HandleError
Time2DOSTime = -1
End Function

Public Function Date2DOSDate(ByVal Value As String) As Long
On Error GoTo error:
Dim BinaryDate As String, Year As String, Month As String, Day As String
Dim temp As String

If Value = "00/00/0000" Then GoTo NoDate:
If Len(Value) < 10 Or Len(Value) > 10 Then GoTo BadFormat:

Year = Dec2Bin((Val(Right(Value, 4) - 1980)))
Month = Dec2Bin(Val(Left(Value, 2)))
Day = Dec2Bin(Val(Right(Left(Value, 5), 2)))

If Len(Year) < 7 Then Year = String(7 - Len(Year), "0") & Year
If Len(Month) < 4 Then Month = String(4 - Len(Month), "0") & Month
If Len(Day) < 5 Then Day = String(5 - Len(Day), "0") & Day

If Len(Year) > 7 Then Year = "1111111"
If Len(Month) > 4 Then Month = "1111"
If Len(Day) > 5 Then Day = "11111"

Date2DOSDate = Bin2Dec(Year & Month & Day)
Exit Function

NoDate:
Date2DOSDate = 0
Exit Function

BadFormat:
MsgBox "Incorrect Date format!", vbExclamation
Date2DOSDate = -1
Exit Function

error:
MsgBox "Incorrect Date format!", vbExclamation
Call HandleError
Date2DOSDate = -1
End Function

Public Sub UnloadForms(ByVal noFrm As String)
Dim frm As Form
On Error Resume Next
'Not frm.Name = frmLogo.Name

Unload frmMapEditor
DoEvents
For Each frm In Forms
    If Not frm.Name = noFrm And Not frm.Name = "frmMain" Then
        Unload frm
    End If
    Set frm = Nothing
Next

End Sub

Private Sub OpenAbilityDB()
On Error GoTo error:

bAbilityDBOpen = True

Set dbAbilities = OpenDatabase(App.Path + "\ability.mdb")
Set rsAbilities = dbAbilities.OpenRecordset("TABLE")
rsAbilities.MoveFirst

Exit Sub

error:
Select Case Err
     Case 3051:
         MsgBox "The Ability Database (ability.mdb) is marked read-only," & vbCrLf & "or a user has it opened in exclusive mode.  Unable to open it.", vbExclamation + vbOKOnly
         bAbilityDBOpen = False
     Case 3024:
         MsgBox "The Ability Database (ability.mdb) is missing!", vbExclamation + vbOKOnly
         bAbilityDBOpen = False
     Case Else
         MsgBox "Ability Database not opened, Error: " & Err.Number & vbCrLf & "Source: " & Err.Source & vbCrLf & "Description: " & Err.Description, vbExclamation + vbOKOnly
         bAbilityDBOpen = False
End Select

End Sub

Public Function NumberKeysOnly(ByVal KeyAscii As Integer) As Integer
NumberKeysOnly = KeyAscii
If KeyAscii = 3 Or KeyAscii = 22 Then Exit Function 'control+v, control+c
If KeyAscii < 48 Or KeyAscii > 57 Then NumberKeysOnly = 0
If KeyAscii = 8 Then NumberKeysOnly = KeyAscii
If KeyAscii = 45 Then NumberKeysOnly = KeyAscii
End Function

Public Sub SelectAll(oTxt As TextBox)
oTxt.SelStart = 0
oTxt.SelLength = Len(oTxt.Text)
End Sub

Public Function UpdateItem() As Integer
ItemStructToRow Itemdatabuf.buf
If bDisableWriting = True Then Exit Function
UpdateItem = BTRCALL(bUpdate, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
End Function

Public Function UpdateClass() As Integer
ClassStructToRow Classdatabuf.buf
If bDisableWriting = True Then Exit Function
UpdateClass = BTRCALL(bUpdate, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
End Function

Public Function UpdateSpell() As Integer
SpellStructToRow Spelldatabuf.buf
If bDisableWriting = True Then Exit Function
UpdateSpell = BTRCALL(bUpdate, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
End Function

Public Function UpdateTextblock() As Integer
TextblockStructToRow TextblockDataBuf.buf
If bDisableWriting = True Then Exit Function
UpdateTextblock = BTRCALL(bUpdate, TextblockPosBlock, TextblockDataBuf, Len(TextblockDataBuf), ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
End Function

Public Function UpdateMonster() As Integer
MonsterStructToRow Monsterdatabuf.buf
If bDisableWriting = True Then Exit Function
UpdateMonster = BTRCALL(bUpdate, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
End Function

Public Function UpdateMessage() As Integer
MessageStructToRow Messagedatabuf.buf
If bDisableWriting = True Then Exit Function
UpdateMessage = BTRCALL(bUpdate, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
End Function

Public Function UpdateRoom() As Integer
RoomStructToRow Roomdatabuf.buf
If bDisableWriting = True Then Exit Function
UpdateRoom = BTRCALL(bUpdate, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
End Function

Public Function UpdateUser() As Integer
UserStructToRow Userdatabuf.buf
If bDisableWriting = True Then Exit Function
UpdateUser = BTRCALL(bUpdate, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0)
End Function

Public Function UpdateShop() As Integer
ShopStructToRow Shopdatabuf.buf
If bDisableWriting = True Then Exit Function
UpdateShop = BTRCALL(bUpdate, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
End Function

Public Function UpdateBank() As Integer
BankStructToRow BankDatabuf.buf
If bDisableWriting = True Then Exit Function
UpdateBank = BTRCALL(bUpdate, BankPosBlock, BankDatabuf, Len(BankDatabuf), ByVal BankKeyBuffer, KEY_BUF_LEN, 0)
End Function

Public Function UpdateGang() As Integer
GangStructToRow GangDatabuf.buf
If bDisableWriting = True Then Exit Function
UpdateGang = BTRCALL(bUpdate, GangPosBlock, GangDatabuf, Len(GangDatabuf), ByVal GangKeyBuffer, KEY_BUF_LEN, 0)
End Function

Public Function UpdateAction() As Integer
ActionStructToRow ActionDatabuf.buf
If bDisableWriting = True Then Exit Function
UpdateAction = BTRCALL(bUpdate, ActionPosBlock, ActionDatabuf, Len(ActionDatabuf), ByVal ActionKeyBuffer, KEY_BUF_LEN, 0)
End Function

Public Function UpdateRace() As Integer
RaceStructToRow Racedatabuf.buf
If bDisableWriting = True Then Exit Function
UpdateRace = BTRCALL(bUpdate, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
End Function

Public Sub StopBtrieve()
On Error GoTo error:
Dim nStatus As Integer

Call BtrieveCloseRace
Call BtrieveCloseClass
Call BtrieveCloseItem
Call BtrieveCloseMonster
Call BtrieveCloseSpell
Call BtrieveCloseMessage
Call BtrieveCloseTextblock
Call BtrieveCloseUser
Call BtrieveCloseBank
Call BtrieveCloseGang
Call BtrieveCloseRoom
Call BtrieveCloseAction
Call BtrieveCloseShop
Call BtrieveRelease

Exit Sub
error:
Call HandleError("StopBtrieve")
Resume Next
End Sub

Private Sub BtrieveCloseRace()
Dim nStatus As Integer

frmMain.stsStatusBar.Panels(5).Text = "Closing Race Database ..."
DoEvents
nStatus = BTRCALL(BCLOSE, RacePosBlock, 0, 0, 0, 0, 0)

End Sub

Private Sub BtrieveCloseClass()
Dim nStatus As Integer

frmMain.stsStatusBar.Panels(5).Text = "Closing Class Database ..."
DoEvents
nStatus = BTRCALL(BCLOSE, ClassPosBlock, 0, 0, 0, 0, 0)

End Sub
Private Sub BtrieveCloseSpell()
Dim nStatus As Integer
frmMain.stsStatusBar.Panels(5).Text = "Closing Spell Database ..."
DoEvents
nStatus = BTRCALL(BCLOSE, SpellPosBlock, 0, 0, 0, 0, 0)
End Sub
Private Sub BtrieveCloseMonster()
Dim nStatus As Integer
frmMain.stsStatusBar.Panels(5).Text = "Closing Monster Database ..."
DoEvents
nStatus = BTRCALL(BCLOSE, MonsterPosBlock, 0, 0, 0, 0, 0)
End Sub
Private Sub BtrieveCloseItem()
Dim nStatus As Integer
frmMain.stsStatusBar.Panels(5).Text = "Closing Item Database ..."
DoEvents
nStatus = BTRCALL(BCLOSE, ItemPosBlock, 0, 0, 0, 0, 0)
End Sub
Private Sub BtrieveCloseShop()
Dim nStatus As Integer
frmMain.stsStatusBar.Panels(5).Text = "Closing Shop Database ..."
DoEvents
nStatus = BTRCALL(BCLOSE, ShopPosBlock, 0, 0, 0, 0, 0)
End Sub
Private Sub BtrieveCloseRoom()
Dim nStatus As Integer
frmMain.stsStatusBar.Panels(5).Text = "Closing Room Database ..."
DoEvents
nStatus = BTRCALL(BCLOSE, RoomPosBlock, 0, 0, 0, 0, 0)
End Sub
Private Sub BtrieveCloseMessage()
Dim nStatus As Integer
frmMain.stsStatusBar.Panels(5).Text = "Closing Message Database ..."
DoEvents
nStatus = BTRCALL(BCLOSE, MessagePosBlock, 0, 0, 0, 0, 0)
End Sub
Private Sub BtrieveCloseTextblock()
Dim nStatus As Integer
frmMain.stsStatusBar.Panels(5).Text = "Closing Textblock Database ..."
DoEvents
nStatus = BTRCALL(BCLOSE, TextblockPosBlock, 0, 0, 0, 0, 0)
End Sub
Private Sub BtrieveCloseUser()
Dim nStatus As Integer
frmMain.stsStatusBar.Panels(5).Text = "Closing User Database ..."
DoEvents
nStatus = BTRCALL(BCLOSE, UserPosBlock, 0, 0, 0, 0, 0)
End Sub
Private Sub BtrieveCloseAction()
Dim nStatus As Integer
frmMain.stsStatusBar.Panels(5).Text = "Closing Action Database ..."
DoEvents
nStatus = BTRCALL(BCLOSE, ActionPosBlock, 0, 0, 0, 0, 0)
End Sub
Private Sub BtrieveCloseBank()
Dim nStatus As Integer
frmMain.stsStatusBar.Panels(5).Text = "Closing Bank Database ..."
DoEvents
nStatus = BTRCALL(BCLOSE, BankPosBlock, 0, 0, 0, 0, 0)
End Sub
Private Sub BtrieveCloseGang()
Dim nStatus As Integer
frmMain.stsStatusBar.Panels(5).Text = "Closing Gang Database ..."
DoEvents
nStatus = BTRCALL(BCLOSE, GangPosBlock, 0, 0, 0, 0, 0)
End Sub
Private Sub BtrieveRelease()
Dim nStatus As Integer
frmMain.stsStatusBar.Panels(5).Text = "Releasing Btrieve Resources ..."
DoEvents
nStatus = BTRCALL(BRESET, 0, 0, 0, 0, 0, 0)
DoEvents
nStatus = BTRCALL(BSTOP, 0, 0, 0, 0, 0, 0)
DoEvents

End Sub

'*******************************************************************************
' Sort a ListView by String, Number, or DateTime -- modified by syntax53
'
' Parameters:
'
'   ListView    Reference to the ListView control to be sorted.
'   Index       Index of the column in the ListView to be sorted. The first
'               column in a ListView has an index value of 1.
'   DataType    Sets whether the data in the column is to be sorted
'               alphabetically, numerically, or by date.
'   Ascending   Sets the direction of the sort. True sorts A-Z (Ascending),
'               and False sorts Z-A (descending)
'-------------------------------------------------------------------------------

Public Sub SortListView(ListView As ListView, ByVal Index As Integer, ByVal dataType As ListDataType, ByVal Ascending As Boolean)

    On Error Resume Next
    Dim i As Integer
    Dim l As Long
    Dim strFormat As String
    
    ' Display the hourglass cursor whilst sorting
    
    Dim lngCursor As Long
    lngCursor = ListView.MousePointer
    ListView.MousePointer = vbHourglass
    
    ' Prevent the ListView control from updating on screen - this is to hide
    ' the changes being made to the listitems, and also to speed up the sort
    
    'LockWindowUpdate ListView.hwnd
    
    Dim blnRestoreFromTag As Boolean
    
    Select Case dataType
    Case ldtString
        
        ' Sort alphabetically. This is the only sort provided by the
        ' MS ListView control (at this time), and as such we don't really
        ' need to do much here
    
        blnRestoreFromTag = False
        
    Case ldtNumber
    
        ' Sort Numerically
    
        strFormat = String$(20, "0") & "." & String$(10, "0")
        
        ' Loop through the values in this column. Re-format the values so
        ' as they can be sorted alphabetically, having already stored their
        ' text values in the tag, along with the tag's original value
    
        With ListView.ListItems
            If (Index = 1) Then
                For l = 1 To .Count
                    With .Item(l)
                        .Tag = .Text & Chr$(0) & .Tag
'                        If IsNumeric(.Text) Then
                            If CDbl(Val(.Text)) >= 0 Then
                                .Text = format(CDbl(Val(.Text)), strFormat)
                            Else
                                .Text = "&" & InvNumber(format(0 - CDbl(Val(.Text)), strFormat))
                            End If
'                        Else
'                            .Text = ""
'                        End If
                    End With
                Next l
            Else
                For l = 1 To .Count
                    With .Item(l).ListSubItems(Index - 1)
                        .Tag = .Text & Chr$(0) & .Tag
'                        If IsNumeric(.Text) Then
                            If CDbl(Val(.Text)) >= 0 Then
                                .Text = format(CDbl(Val(.Text)), strFormat)
                            Else
                                .Text = "&" & InvNumber(format(0 - CDbl(Val(.Text)), strFormat))
                            End If
'                        Else
'                            .Text = ""
'                        End If
                    End With
                Next l
            End If
        End With
        
        blnRestoreFromTag = True
    
    Case ldtDateTime
    
        ' Sort by date.
        
        strFormat = "YYYYMMDDHhNnSs"
        
        Dim dte As Date
    
        ' Loop through the values in this column. Re-format the dates so as they
        ' can be sorted alphabetically, having already stored their visible
        ' values in the tag, along with the tag's original value
    
        With ListView.ListItems
            If (Index = 1) Then
                For l = 1 To .Count
                    With .Item(l)
                        .Tag = .Text & Chr$(0) & .Tag
                        dte = CDate(.Text)
                        .Text = format$(dte, strFormat)
                    End With
                Next l
            Else
                For l = 1 To .Count
                    With .Item(l).ListSubItems(Index - 1)
                        .Tag = .Text & Chr$(0) & .Tag
                        dte = CDate(.Text)
                        .Text = format$(dte, strFormat)
                    End With
                Next l
            End If
        End With
        
        blnRestoreFromTag = True
        
    End Select
    
    ' Sort the ListView Alphabetically
    
    ListView.SortOrder = IIf(Ascending, lvwAscending, lvwDescending)
    ListView.SortKey = Index - 1
    ListView.Sorted = True
    
    ' Restore the Text Values if required
    
    If blnRestoreFromTag Then
        
        ' Restore the previous values to the 'cells' in this column of the list
        ' from the tags, and also restore the tags to their original values
        
        With ListView.ListItems
            If (Index = 1) Then
                For l = 1 To .Count
                    With .Item(l)
                        i = InStr(.Tag, Chr$(0))
                        .Text = Left$(.Tag, i - 1)
                        .Tag = Mid$(.Tag, i + 1)
                    End With
                Next l
            Else
                For l = 1 To .Count
                    With .Item(l).ListSubItems(Index - 1)
                        i = InStr(.Tag, Chr$(0))
                        .Text = Left$(.Tag, i - 1)
                        .Tag = Mid$(.Tag, i + 1)
                    End With
                Next l
            End If
        End With
    End If
    
    ' Unlock the list window so that the OCX can update it
    
    'LockWindowUpdate 0&
    
    ' Restore the previous cursor
    
    ListView.MousePointer = lngCursor
    

End Sub

'*******************************************************************************
' Modifies a numeric string to allow it to be sorted alphabetically
'-------------------------------------------------------------------------------

Private Function InvNumber(ByVal Number As String) As String
    Static i As Integer
    For i = 1 To Len(Number)
        Select Case Mid$(Number, i, 1)
            Case "-": Mid$(Number, i, 1) = " "
            Case "0": Mid$(Number, i, 1) = "9"
            Case "1": Mid$(Number, i, 1) = "8"
            Case "2": Mid$(Number, i, 1) = "7"
            Case "3": Mid$(Number, i, 1) = "6"
            Case "4": Mid$(Number, i, 1) = "5"
            Case "5": Mid$(Number, i, 1) = "4"
            Case "6": Mid$(Number, i, 1) = "3"
            Case "7": Mid$(Number, i, 1) = "2"
            Case "8": Mid$(Number, i, 1) = "1"
            Case "9": Mid$(Number, i, 1) = "0"
        End Select
    Next
    InvNumber = Number
End Function

'*******************************************************************************
'
'-------------------------------------------------------------------------------

Public Sub GetTitleBarOffset()
Dim TitleInfo As TITLEBARINFO, OSVer As cnWin32Ver

OSVer = Win32Ver
If OSVer <= win95 Then GoTo win95:

TitleInfo.cbSize = Len(TitleInfo)
GetTitleBarInfo frmMain.hwnd, TitleInfo

TITLEBAR_OFFSET = (TitleInfo.rcTitleBar.Bottom - TitleInfo.rcTitleBar.Top) * 15
If TITLEBAR_OFFSET > 285 Then
    TITLEBAR_OFFSET = TITLEBAR_OFFSET - 285
Else
    TITLEBAR_OFFSET = 0
End If

Exit Sub

win95:

TITLEBAR_OFFSET = 0

End Sub

Public Function ChangeCallLetters(ByVal sLetters As String, ByVal sChange As String) As String
Dim x As Long

ChangeCallLetters = sChange

For x = 1 To Len(ChangeCallLetters)
    If UCase(Mid(ChangeCallLetters, x, 3)) = UCase("W" & sLetters) Then
        ChangeCallLetters = Mid(ChangeCallLetters, 1, x - 1) & "W" & UCase(strDatCallLetters) & Mid(ChangeCallLetters, x + 3)
    End If
Next x

End Function

Public Function DirectoryExists(ByVal sDirectory As String) As Boolean
On Error GoTo error:
Dim fso As FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FolderExists(sDirectory) Then DirectoryExists = True

Set fso = Nothing
Exit Function
error:
Call HandleError
Set fso = Nothing
End Function

Public Function FileExists(ByVal sFileName As String) As Boolean
On Error GoTo error:
Dim fso As FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(sFileName) Then FileExists = True

Set fso = Nothing
Exit Function
error:
Call HandleError
Set fso = Nothing
End Function

Public Function IsNum(ByVal sChar As String) As Boolean
On Error GoTo error:

Select Case sChar
    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9": IsNum = True
End Select

Exit Function

error:
Call HandleError
End Function

Public Function FindDll(sDll As String) As String
Dim fso As FileSystemObject, fldr As String, fldr1 As String, fldr2 As String

Set fso = CreateObject("Scripting.FileSystemObject")

fldr = App.Path
fldr1 = fso.GetSpecialFolder(SystemFolder)
fldr2 = fso.GetSpecialFolder(WindowsFolder)

If fso.FileExists(fldr & "\" & sDll) = True Then: FindDll = fldr & "\" & sDll: Exit Function
If fso.FileExists(fldr1 & "\" & sDll) = True Then: FindDll = fldr1 & "\" & sDll: Exit Function
If fso.FileExists(fldr2 & "\" & sDll) = True Then: FindDll = fldr2 & "\" & sDll: Exit Function

FindDll = sDll

Set fso = Nothing

End Function

Public Function AutoSizeDropDownWidth(Combo As Object) As Boolean
'**************************************************************
'PURPOSE: Automatically size the combo box drop down width
'         based on the width of the longest item in the combo box

'PARAMETERS: Combo - ComboBox to size

'RETURNS: True if successful, false otherwise

'ASSUMPTIONS: 1. Form's Scale Mode is vbTwips, which is why
'                conversion from twips to pixels are made.
'                API functions require units in pixels
'
'             2. Combo Box's parent is a form or other
'                container that support the hDC property

'EXAMPLE: AutoSizeDropDownWidth Combo1
'****************************************************************
Dim lRet As Long
Dim lCurrentWidth As Single
Dim rectCboText As RECT
Dim lParentHDC As Long
Dim lListCount As Long
Dim lCtr As Long
Dim lTempWidth As Long
Dim lWidth As Long
Dim sSavedFont As String
Dim sngSavedSize As Single
Dim bSavedBold As Boolean
Dim bSavedItalic As Boolean
Dim bSavedUnderline As Boolean
Dim bFontSaved As Boolean

On Error GoTo ErrorHandler

If Not TypeOf Combo Is ComboBox Then Exit Function
lParentHDC = Combo.Parent.hdc
If lParentHDC = 0 Then Exit Function
lListCount = Combo.ListCount
If lListCount = 0 Then Exit Function


'Change font of parent to combo box's font
'Save first so it can be reverted when finished
'this is necessary for drawtext API Function
'which is used to determine longest string in combo box
With Combo.Parent

    sSavedFont = .FontName
    sngSavedSize = .FontSize
    bSavedBold = .FontBold
    bSavedItalic = .FontItalic
    bSavedUnderline = .FontUnderline
    
    .FontName = Combo.FontName
    .FontSize = Combo.FontSize
    .FontBold = Combo.FontBold
    .FontItalic = Combo.FontItalic
    .FontUnderline = Combo.FontItalic

End With

bFontSaved = True

'Get the width of the largest item
For lCtr = 0 To lListCount
   DrawText lParentHDC, Combo.List(lCtr), -1, rectCboText, _
        DT_CALCRECT
   'adjust the number added (20 in this case to
   'achieve desired right margin
   lTempWidth = rectCboText.Right - rectCboText.Left + 20

   If (lTempWidth > lWidth) Then
      lWidth = lTempWidth
   End If
Next
 
lCurrentWidth = SendMessageLong(Combo.hwnd, CB_GETDROPPEDWIDTH, _
    0, 0)

If lCurrentWidth > lWidth Then 'current drop-down width is
'                               sufficient

    AutoSizeDropDownWidth = True
    GoTo ErrorHandler
    Exit Function
End If
 
'don't allow drop-down width to
'exceed screen.width
 
   If lWidth > Screen.Width \ Screen.TwipsPerPixelX - 20 Then _
    lWidth = Screen.Width \ Screen.TwipsPerPixelX - 20

lRet = SendMessageLong(Combo.hwnd, CB_SETDROPPEDWIDTH, lWidth, 0)

AutoSizeDropDownWidth = lRet > 0
ErrorHandler:
On Error Resume Next
If bFontSaved Then
'restore parent's font settings
  With Combo.Parent
    .FontName = sSavedFont
    .FontSize = sngSavedSize
    .FontUnderline = bSavedUnderline
    .FontBold = bSavedBold
    .FontItalic = bSavedItalic
 End With
End If

End Function

Public Sub ReloadApp()
On Error GoTo error:
Dim nStatus As Integer
Dim frm As Form

Unload frmMapEditor
DoEvents
For Each frm In Forms
    If Not frm.Name = "frmSettings" And Not frm.Name = "frmMain" Then
        Unload frm
    End If
    Set frm = Nothing
Next

DoEvents

Load frmSplash
'frmSplash.Left = frmMain.Width / 4
'frmSplash.Top = frmMain.Height / 4
frmSplash.Show

'frmSettings.Hide
DoEvents

Call StopBtrieve

DoEvents
Call Startup

Unload frmSplash
If Not bKeepSettingsOpen Then Unload frmSettings
DoEvents

Exit Sub
error:
Call HandleError

End Sub

Public Sub FindAbilityNumber(txtNum As TextBox, txtName As TextBox) 'Optional ByVal sSearchText As String)
On Error GoTo error:

If bAbilityDBOpen = False Then
    txtNum.Text = RemoveCharacter(txtNum.Text, "+")
    txtNum.Text = RemoveCharacter(txtNum.Text, "-")
    txtNum.Text = RemoveCharacter(txtNum.Text, "=")
    txtName.ForeColor = vbRed
    txtName.FontBold = True
    txtName.FontItalic = False
    txtName.Text = "Ability DB not open!"
    txtName.ToolTipText = "Ability DB is not opened!"
    Exit Sub
End If

If txtNum.Text = "" Then
    txtNum.Text = "0"
    txtNum.SelStart = 0
    txtNum.SelLength = 1
End If
txtNum.Text = RemoveCharacter(txtNum.Text, "+")
txtNum.Text = RemoveCharacter(txtNum.Text, "-")
txtNum.Text = RemoveCharacter(txtNum.Text, "=")

rsAbilities.Index = "PrimaryKey"
rsAbilities.Seek "=", Val(txtNum.Text)
If Not rsAbilities.NoMatch Then
    txtName.ForeColor = RGB(rsAbilities.Fields("RED"), rsAbilities.Fields("GREEN"), rsAbilities.Fields("BLUE"))
    txtName.FontBold = True
    txtName.FontItalic = False
    txtName.Text = rsAbilities.Fields("Name")
    txtName.ToolTipText = rsAbilities.Fields("Description")
Else
    txtName.ForeColor = vbRed
    txtName.FontBold = True
    txtName.FontItalic = True
    txtName.Text = "UNKNOWN"
    txtName.ToolTipText = "The Currently Value is unknown and could be incorrect!"
End If

out:
Exit Sub
error:
Call HandleError("FindAbilityNumber")
Resume out:
    
End Sub
Public Sub FindAbilityName(ByVal nKeyCode As Integer, txtNum As TextBox, txtName As TextBox) 'Optional ByVal sSearchText As String)
Dim nALen As Integer
Dim sAName As String
On Error GoTo error:

'If nKeyCode = vbKeyBack Then Exit Sub
If nKeyCode = vbKeyShift Then Exit Sub
If nKeyCode = vbKeyHome Then Exit Sub
If nKeyCode = vbKeyUp Then Exit Sub
If nKeyCode = vbKeyDown Then Exit Sub
If nKeyCode = vbKeyLeft Then Exit Sub
If nKeyCode = vbKeyEscape Then Exit Sub
If nKeyCode = vbKeyTab Then Exit Sub

If bAbilityDBOpen = False Then
    txtNum.Text = RemoveCharacter(txtNum.Text, "+")
    txtNum.Text = RemoveCharacter(txtNum.Text, "-")
    txtNum.Text = RemoveCharacter(txtNum.Text, "=")
    txtName.ForeColor = vbRed
    txtName.FontBold = True
    txtName.FontItalic = False
    txtName.Text = "Ability DB not open!"
    txtName.ToolTipText = "Ability DB is not opened!"
    Exit Sub
End If

rsAbilities.MoveFirst
'If sSearchText = "" Then
'    If txtName.SelLength > 0 Then
'        sAName = LCase(Mid(txtName.Text, 1, txtName.SelStart))
'    Else
'        sAName = LCase(txtName.Text)
'    End If
'    sAName = sAName & LCase(Chr(nKeyCode))
'Else
'    sAName = sSearchText
'End If
sAName = LCase(txtName.Text)
nALen = Len(sAName)
'Debug.Print sAName

If nKeyCode = vbKeyRight Or nKeyCode = vbKeyDown Then
    rsAbilities.Index = "PrimaryKey"
    rsAbilities.Seek "=", Val(txtNum.Text)
    If rsAbilities.NoMatch Then
        rsAbilities.MoveFirst
    Else
        rsAbilities.MoveNext
    End If
End If

Do While Not rsAbilities.EOF
    If InStr(1, LCase(rsAbilities.Fields("Name")), sAName) > 0 Then
        txtName.ForeColor = RGB(rsAbilities.Fields("RED"), rsAbilities.Fields("GREEN"), rsAbilities.Fields("BLUE"))
        txtName.FontBold = True
        txtName.FontItalic = False
        'txtName.Text = rsAbilities.Fields("Name")
        txtName.ToolTipText = rsAbilities.Fields("Name") & ": " & rsAbilities.Fields("Description")
        'txtName.SelStart = nALen
        'txtName.SelLength = Len(rsAbilities.Fields("Name")) - nALen
        txtNum.Text = rsAbilities.Fields("Number")
        Exit Do
    Else
        rsAbilities.MoveNext
    End If
Loop

out:
Exit Sub
error:
Call HandleError("FindAbilityName")
Resume out:

End Sub


Public Sub BuildControlRoomList()
On Error GoTo error:
Dim nStatus As Integer, x As Integer, nMaxRooms As Long
Dim sControlRoom As String, sRefRoom As String, sArr() As String, sArr2() As String

ControlRoomList.RemoveAll

nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then MsgBox "Error getting first room, error: " & BtrieveErrorCode(nStatus): Exit Sub

nStatus = BTRCALL(BSTAT, RoomPosBlock, DBStatDatabuf, Len(Roomdatabuf), 0, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    nMaxRooms = 30000
Else
    DBStatRowToStruct DBStatDatabuf.buf
    nMaxRooms = DBStat.nRecords
End If

bStopControlBuild = False

frmProgressBar.sCaption = "Building Control Room List"
frmProgressBar.lblCaption = "Building Control Room List..."
frmProgressBar.cmdCancel.Enabled = True

Call frmProgressBar.SetRange(nMaxRooms)

frmProgressBar.lblNote.Visible = False
frmProgressBar.lblPanel(0).Caption = ""
frmProgressBar.lblPanel(1).Caption = ""
frmProgressBar.Show
DoEvents

frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_MP
frmProgressBar.lblPanel(1).Caption = "Scanning Rooms..."

nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
Do While nStatus = 0 And bStopControlBuild = False
    RoomRowToStruct Roomdatabuf.buf
    
    sRefRoom = Roomrec.MapNumber & "/" & Roomrec.RoomNumber
    'frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_MP
    'frmProgressBar.lblPanel(1).Caption = sRefRoom
    DoEvents
    
    If Roomrec.ControlRoom > 0 Then
        
'        sControlRoom = Roomrec.MapNumber & "/" & Roomrec.ControlRoom
'
'        If ControlRoomList.Exists(sControlRoom) Then
'            If InStr(1, ControlRoomList.Item(sControlRoom), "more]", vbTextCompare) = 0 Then
'                If InStr(1, ControlRoomList.Item(sControlRoom), ",", vbTextCompare) Then
'                    sArr = Split(ControlRoomList.Item(sControlRoom), ",")
'                    If UBound(sArr()) > 14 Then
'                        ControlRoomList.Item(sControlRoom) = ControlRoomList.Item(sControlRoom) & " [+1 more]"
'                        'Debug.Print ControlRoomList.Item(sControlRoom)
'                    Else
'                        ControlRoomList.Item(sControlRoom) = ControlRoomList.Item(sControlRoom) & ", " & Roomrec.RoomNumber
'                    End If
'                Else
'                    ControlRoomList.Item(sControlRoom) = ControlRoomList.Item(sControlRoom) & ", " & Roomrec.RoomNumber
'                End If
'            Else
'                sArr = Split(ControlRoomList.Item(sControlRoom), "+")
'                If UBound(sArr()) = 1 Then
'                    sArr2 = Split(sArr(1))
'                    If UBound(sArr2()) = 1 Then
'                        x = Val(sArr2(0))
'                        ControlRoomList.Item(sControlRoom) = sArr(0) & "+" & (x + 1) & " more]"
'                    End If
'                End If
'            End If
'        Else
'            ControlRoomList.add sControlRoom, sRefRoom
'        End If
        Call AddControlRoom(Roomrec.MapNumber, Roomrec.ControlRoom, Roomrec.RoomNumber)
    End If
    
    Call frmProgressBar.IncreaseProgress
    nStatus = BTRCALL(BGETNEXT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

frmProgressBar.ProgressBar.Value = frmProgressBar.ProgressBar.Max
DoEvents

kill:
On Error Resume Next
Unload frmProgressBar
Exit Sub
error:
Call HandleError("BuildControlRoomList")
bStopControlBuild = True
Resume kill:
End Sub

Public Sub AddControlRoom(nMap As Long, nControlRoom As Long, nRefRoom As Long)
Dim sControlRoom As String
On Error GoTo error:

sControlRoom = nMap & "/" & nControlRoom

If ControlRoomList.Exists(sControlRoom) Then
    If InStr(1, ControlRoomList.Item(sControlRoom), "(" & nRefRoom & ")", vbTextCompare) = 0 Then
        ControlRoomList.Item(sControlRoom) = ControlRoomList.Item(sControlRoom) & "," & "(" & nRefRoom & ")"
    End If
Else
    ControlRoomList.add sControlRoom, "(" & nRefRoom & ")"
End If


kill:
On Error Resume Next
Exit Sub
error:
Call HandleError("AddControlRoom")
Resume kill:
End Sub

Public Function GetControlRoomListByRoom(nMap As Long, nControlRoom As Long, Optional nMaxInList As Integer = 20, Optional bCountAlwaysFirst As Boolean = False) As String
Dim sControlRoom As String, sArr() As String, sArr2() As String
Dim sRoomList As String, x As Long, nTotalRefs As Long
On Error GoTo error:

sControlRoom = nMap & "/" & nControlRoom

If ControlRoomList.Exists(sControlRoom) Then
    sRoomList = ControlRoomList.Item(sControlRoom)
    
    If Not InStr(1, sRoomList, ",", vbTextCompare) = 0 Then
        
        sArr = Split(sRoomList, ",")
        nTotalRefs = UBound(sArr()) + 1
        
        If nTotalRefs > nMaxInList Then
            sRoomList = ""
            For x = 0 To nMaxInList - 1
                If x > 0 Then sRoomList = sRoomList & ", "
                sRoomList = sRoomList & sArr(x)
            Next x
            
            sRoomList = sRoomList & " [+" & (nTotalRefs - nMaxInList) & " more"
            If bCountAlwaysFirst Then
                sRoomList = sRoomList & "]"
            Else
                sRoomList = sRoomList & ", " & nTotalRefs & " total]"
            End If
        Else
            sRoomList = Replace(sRoomList, ",", ", ")
        End If
    Else
        nTotalRefs = 1
    End If
    
    GetControlRoomListByRoom = Replace(Replace(sRoomList, "(", "", 1, -1, vbTextCompare), ")", "", 1, -1, vbTextCompare)
    If bCountAlwaysFirst Then
        GetControlRoomListByRoom = nTotalRefs & " total: " & GetControlRoomListByRoom
    End If
Else
    GetControlRoomListByRoom = "N/A"
End If

kill:
On Error Resume Next
Exit Function
error:
Call HandleError("GetControlRoomListByRoom")
Resume kill:
End Function

Public Function StringOfNumbersToArray(sNumberString As String) As String()
Dim x As Long, sRet() As String
On Error GoTo error:

If InStr(1, sNumberString, ",", vbTextCompare) = 0 Then
    ReDim sRet(0)
    sRet(0) = Val(Replace(Replace(sNumberString, "(", "", 1, -1, vbTextCompare), ")", "", 1, -1, vbTextCompare))
Else
    sRet = Split(sNumberString, ",")
    For x = 0 To UBound(sRet())
        sRet(x) = Val(Replace(Replace(sRet(x), "(", "", 1, -1, vbTextCompare), ")", "", 1, -1, vbTextCompare))
    Next x
End If

StringOfNumbersToArray = sRet

out:
On Error Resume Next
Exit Function
error:
Call HandleError("NumberStringToArray")
Resume out:
End Function

Public Function MergeStringArrays(sArr1() As String, sArr2() As String) As String()
Dim x As Long, y As Long, bMatch As Boolean, sRet() As String
On Error GoTo error:

If UBound(sArr1()) > UBound(sArr2()) Then
    ReDim sRet(UBound(sArr1()))
    sRet = sArr1
    For x = LBound(sArr2()) To UBound(sArr2())
        bMatch = False
        For y = LBound(sRet()) To UBound(sRet())
            If sArr2(x) = sRet(y) Then
                bMatch = True
                Exit For
            End If
        Next y
        If Not bMatch Then
            ReDim Preserve sRet(UBound(sRet()) + 1)
            sRet(UBound(sRet())) = sArr2(x)
        End If
    Next x
Else
    ReDim sRet(UBound(sArr2()))
    sRet = sArr2
    For x = LBound(sArr1()) To UBound(sArr1())
        bMatch = False
        For y = LBound(sRet()) To UBound(sRet())
            If sArr1(x) = sRet(y) Then
                bMatch = True
                Exit For
            End If
        Next y
        If Not bMatch Then
            ReDim Preserve sRet(UBound(sRet()) + 1)
            sRet(UBound(sRet())) = sArr1(x)
        End If
    Next x
End If

MergeStringArrays = sRet

out:
On Error Resume Next
Exit Function
error:
Call HandleError("MergeStringArrays")
Resume out:
End Function

