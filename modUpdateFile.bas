Attribute VB_Name = "modUpdateFile"
Option Base 0
Option Explicit

Dim UpdateFileLoaded As Boolean
Dim recnum As Long
Dim nPerCalc As Long
Dim InitialWindowState As Integer

Public Sub CompileUpdatefile(Optional bBlankFile As Boolean)
On Error GoTo error:
Dim nYesNo As Integer, x As Integer, nTotalRecs As Long
Dim StartTime As Variant, nTotalTime As Double, sTotalTime As String

InitialWindowState = frmMain.WindowState

UpdateFileLoaded = False

If bBlankFile Then
    nYesNo = MsgBox("This will delete the current update file and compile a blank w" & strDatCallLetters _
        & strDatSuffix_UPDAT & ", are you sure?", _
        vbYesNo + vbQuestion + vbDefaultButton2, "Compile blank update file?")
    If nYesNo <> vbYes Then
        Call StopUpdate
        Exit Sub
    End If
Else
    nYesNo = MsgBox("This will close all open forms and compile a w" _
        & strDatCallLetters & strDatSuffix_UPDAT & ", are you sure?" & vbCrLf & vbCrLf _
        & "NOTE: This will take anywhere from 1-20 minutes.", _
        vbYesNo + vbQuestion + vbDefaultButton2, "Compile an update file?")
    If nYesNo <> vbYes Then
        Call StopUpdate
        Exit Sub
    End If
    UnloadForms (frmMain.Name)
    
    frmMain.WindowState = vbMinimized
    DoEvents
    
    frmProgressBar.sCaption = "Compiling Update File"
    frmProgressBar.lblCaption = "Compiling update file..."
    frmProgressBar.cmdCancel.Enabled = True
    
    Call frmProgressBar.SetRange(CalcTotalRecords)
    
    frmProgressBar.lblNote.Visible = True
    frmProgressBar.lblPanel(0).Caption = ""
    frmProgressBar.lblPanel(1).Caption = ""
    frmProgressBar.Show
    DoEvents
End If

nPerCalc = 0

frmMain.Enabled = False

Call CreateUpdateFile
If UpdateFileLoaded = False Then
    Call StopUpdate
    Exit Sub
End If

If bBlankFile Then
    Call InsertBlank
    Call StopUpdate
    Exit Sub
End If

StartTime = Timer

Updaterec.recnumber = 1

ClearUpdatebuf
If UpdateFileLoaded = True Then Call InsertMessages
ClearUpdatebuf
DoEvents
If UpdateFileLoaded = True Then Call InsertTextblocks
ClearUpdatebuf
DoEvents
If UpdateFileLoaded = True Then Call InsertRaces
ClearUpdatebuf
DoEvents
If UpdateFileLoaded = True Then Call InsertClasses
ClearUpdatebuf
DoEvents
If UpdateFileLoaded = True Then Call InsertSpells
ClearUpdatebuf
DoEvents
If UpdateFileLoaded = True Then Call InsertItems
ClearUpdatebuf
DoEvents
If UpdateFileLoaded = True Then Call InsertShops
ClearUpdatebuf
DoEvents
If UpdateFileLoaded = True Then Call InsertMonsters
ClearUpdatebuf
DoEvents
If UpdateFileLoaded = True Then Call InsertRooms
ClearUpdatebuf
DoEvents
If UpdateFileLoaded = True Then Call InsertActions
DoEvents

'frmProgressBar.sCaption = " 100% - NMR: Compiling Update File"
frmProgressBar.ProgressBar.Value = frmProgressBar.ProgressBar.Max
nTotalTime = Timer - StartTime
sTotalTime = CStr(Round(CDbl(nTotalTime / 60), 2))
sTotalTime = Left(sTotalTime, InStr(1, sTotalTime, ".") + 2)
DoEvents

If UpdateFileLoaded = True Then MsgBox "Update file compiled successfully." & vbCrLf & vbCrLf & "Total time: " & sTotalTime & " minutes.", vbInformation

Call StopUpdate

Exit Sub
error:
Call HandleError
Call StopUpdate
End Sub
Private Sub InsertBlank()
Dim x As Integer, nStatus As Integer

Updaterec.filenum = 0
Updaterec.recnumber = 0
For x = 1 To UpdateDataBufSize
    Updaterec.Data(x) = &H0
Next

nStatus = BTRCALL(BINSERT, UpdatePosBlock, Updaterec, Len(Updaterec), ByVal UpdateKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error inserting blank record, error: " & BtrieveErrorCode(nStatus)
    Exit Sub
Else
    MsgBox "Blank update file created successfully.", vbOKOnly + vbInformation
End If

End Sub
Private Sub InsertMessages()
Dim nStatus As Integer, nYesNo As Integer, x As Long

frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_MSG
recnum = 1
DoEvents

Updaterec.filenum = 8

ClearUpdatebuf
nStatus = BTRCALL(BGETFIRST, MessagePosBlock, Updatebuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)

Do While nStatus = 0 And UpdateFileLoaded = True

    For x = 1 To Len(Messagedatabuf)
        Updaterec.Data(x) = Updatebuf.Data(x)
    Next
    
    nStatus = BTRCALL(BINSERT, UpdatePosBlock, Updaterec, Len(Updaterec), ByVal UpdateKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Update Insert Error: " & BtrieveErrorCode(nStatus)
        nYesNo = MsgBox("Do you want to stop the process?", vbYesNo, "Stop update file creation?")
        If nYesNo = vbYes Then
            StopUpdate
            Exit Sub
        End If
    End If
    
    'ClearUpdatebuf (Len(Messagedatabuf))
    
    nStatus = BTRCALL(BGETNEXT, MessagePosBlock, Updatebuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
    
    Updaterec.recnumber = Updaterec.recnumber + 1
    recnum = recnum + 1
    frmProgressBar.lblPanel(1).Caption = recnum
    Call frmProgressBar.IncreaseProgress
    If Not bUseCPU Then DoEvents
    
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    nYesNo = MsgBox("Error exporting Messages, Btrieve Error: " & BtrieveErrorCode(nStatus, True) _
        & vbCrLf & "Continue?", vbYesNo + vbDefaultButton2 + vbQuestion)
    If nYesNo = vbNo Then Call StopUpdate
End If

End Sub

Private Sub InsertTextblocks()
Dim nStatus As Integer, nYesNo As Integer, x As Long

frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_TEXT
recnum = 1
DoEvents

Updaterec.filenum = 9

ClearUpdatebuf
nStatus = BTRCALL(BGETFIRST, TextblockPosBlock, Updatebuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)

Do While nStatus = 0 And UpdateFileLoaded = True
    
    For x = 1 To TextblockMaxBufSize
        Updaterec.Data(x) = Updatebuf.Data(x)
    Next
    
    nStatus = BTRCALL(BINSERT, UpdatePosBlock, Updaterec, Len(Updaterec), ByVal UpdateKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Update Insert Error: " & BtrieveErrorCode(nStatus)
        nYesNo = MsgBox("Do you want to stop the process?", vbYesNo, "Stop update file creation?")
        If nYesNo = vbYes Then
            StopUpdate
            Exit Sub
        End If
    End If
    
    'ClearUpdatebuf
    
    nStatus = BTRCALL(BGETNEXT, TextblockPosBlock, Updatebuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
    
    Updaterec.recnumber = Updaterec.recnumber + 1
    recnum = recnum + 1
    frmProgressBar.lblPanel(1).Caption = recnum
    Call frmProgressBar.IncreaseProgress
    If Not bUseCPU Then DoEvents
    
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    nYesNo = MsgBox("Error exporting Textblocks, Btrieve Error: " & BtrieveErrorCode(nStatus, True) _
        & vbCrLf & "Continue?", vbYesNo + vbDefaultButton2 + vbQuestion)
    If nYesNo = vbNo Then Call StopUpdate
End If

End Sub

Private Sub InsertRaces()
Dim nStatus As Integer, nYesNo As Integer, x As Long

frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_RACE
recnum = 1
DoEvents

Updaterec.filenum = 1

ClearUpdatebuf
nStatus = BTRCALL(BGETFIRST, RacePosBlock, Updatebuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)

Do While nStatus = 0 And UpdateFileLoaded = True
    
    For x = 1 To Len(Racedatabuf)
        Updaterec.Data(x) = Updatebuf.Data(x)
    Next
    
    nStatus = BTRCALL(BINSERT, UpdatePosBlock, Updaterec, Len(Updaterec), ByVal UpdateKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Update Insert Error: " & BtrieveErrorCode(nStatus)
        nYesNo = MsgBox("Do you want to stop the process?", vbYesNo, "Stop update file creation?")
        If nYesNo = vbYes Then
            StopUpdate
            Exit Sub
        End If
    End If
    
    'ClearUpdatebuf
    
    nStatus = BTRCALL(BGETNEXT, RacePosBlock, Updatebuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
    
    Updaterec.recnumber = Updaterec.recnumber + 1
    recnum = recnum + 1
    frmProgressBar.lblPanel(1).Caption = recnum
    Call frmProgressBar.IncreaseProgress
    If Not bUseCPU Then DoEvents
    
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    nYesNo = MsgBox("Error exporting Races, Btrieve Error: " & BtrieveErrorCode(nStatus, True) _
        & vbCrLf & "Continue?", vbYesNo + vbDefaultButton2 + vbQuestion)
    If nYesNo = vbNo Then Call StopUpdate
End If

End Sub

Private Sub InsertClasses()
Dim nStatus As Integer, nYesNo As Integer, x As Long

frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_CLASS
recnum = 1
DoEvents

Updaterec.filenum = 2

ClearUpdatebuf
nStatus = BTRCALL(BGETFIRST, ClassPosBlock, Updatebuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)

Do While nStatus = 0 And UpdateFileLoaded = True
    
    For x = 1 To Len(Classdatabuf)
        Updaterec.Data(x) = Updatebuf.Data(x)
    Next
    
    nStatus = BTRCALL(BINSERT, UpdatePosBlock, Updaterec, Len(Updaterec), ByVal UpdateKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Update Insert Error: " & BtrieveErrorCode(nStatus)
        nYesNo = MsgBox("Do you want to stop the process?", vbYesNo, "Stop update file creation?")
        If nYesNo = vbYes Then
            StopUpdate
            Exit Sub
        End If
    End If
    
    'ClearUpdatebuf
    
    nStatus = BTRCALL(BGETNEXT, ClassPosBlock, Updatebuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
    
    Updaterec.recnumber = Updaterec.recnumber + 1
    recnum = recnum + 1
    frmProgressBar.lblPanel(1).Caption = recnum
    Call frmProgressBar.IncreaseProgress
    If Not bUseCPU Then DoEvents

Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    nYesNo = MsgBox("Error exporting Classes, Btrieve Error: " & BtrieveErrorCode(nStatus, True) _
        & vbCrLf & "Continue?", vbYesNo + vbDefaultButton2 + vbQuestion)
    If nYesNo = vbNo Then Call StopUpdate
End If

End Sub

Private Sub InsertSpells()
Dim nStatus As Integer, nYesNo As Integer, x As Long

frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_SPELS
recnum = 1
DoEvents

Updaterec.filenum = 6

ClearUpdatebuf
nStatus = BTRCALL(BGETFIRST, SpellPosBlock, Updatebuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)

Do While nStatus = 0 And UpdateFileLoaded = True

    For x = 1 To Len(Spelldatabuf)
        Updaterec.Data(x) = Updatebuf.Data(x)
    Next
    
    nStatus = BTRCALL(BINSERT, UpdatePosBlock, Updaterec, Len(Updaterec), ByVal UpdateKeyBuffer, KEY_BUF_LEN, 0)
    
    If Not nStatus = 0 Then
        MsgBox "Update Insert Error: " & BtrieveErrorCode(nStatus)
        nYesNo = MsgBox("Do you want to stop the process?", vbYesNo, "Stop update file creation?")
        If nYesNo = vbYes Then
            StopUpdate
            Exit Sub
        End If
    End If
    
    'ClearUpdatebuf
    
    nStatus = BTRCALL(BGETNEXT, SpellPosBlock, Updatebuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
    
    Updaterec.recnumber = Updaterec.recnumber + 1
    recnum = recnum + 1
    frmProgressBar.lblPanel(1).Caption = recnum
    Call frmProgressBar.IncreaseProgress
    If Not bUseCPU Then DoEvents

Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    nYesNo = MsgBox("Error exporting Spells, Btrieve Error: " & BtrieveErrorCode(nStatus, True) _
        & vbCrLf & "Continue?", vbYesNo + vbDefaultButton2 + vbQuestion)
    If nYesNo = vbNo Then Call StopUpdate
End If

End Sub

Private Sub InsertItems()
Dim nStatus As Integer, nYesNo As Integer, x As Long

frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_ITEMS
recnum = 1
DoEvents

Updaterec.filenum = 5

ClearUpdatebuf
nStatus = BTRCALL(BGETFIRST, ItemPosBlock, Updatebuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)

Do While nStatus = 0 And UpdateFileLoaded = True
    
    For x = 1 To Len(Itemdatabuf)
        Updaterec.Data(x) = Updatebuf.Data(x)
    Next
    
    nStatus = BTRCALL(BINSERT, UpdatePosBlock, Updaterec, Len(Updaterec), ByVal UpdateKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Update Insert Error: " & BtrieveErrorCode(nStatus)
        nYesNo = MsgBox("Do you want to stop the process?", vbYesNo, "Stop update file creation?")
        If nYesNo = vbYes Then
            StopUpdate
            Exit Sub
        End If
    End If
    
    'ClearUpdatebuf
    
    nStatus = BTRCALL(BGETNEXT, ItemPosBlock, Updatebuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
    
    Updaterec.recnumber = Updaterec.recnumber + 1
    recnum = recnum + 1
    frmProgressBar.lblPanel(1).Caption = recnum
    Call frmProgressBar.IncreaseProgress
    If Not bUseCPU Then DoEvents
    
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    nYesNo = MsgBox("Error exporting Items, Btrieve Error: " & BtrieveErrorCode(nStatus, True) _
        & vbCrLf & "Continue?", vbYesNo + vbDefaultButton2 + vbQuestion)
    If nYesNo = vbNo Then Call StopUpdate
End If

End Sub

Private Sub InsertShops()
Dim nStatus As Integer, nYesNo As Integer, x As Long

frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_SHOPS
recnum = 1
DoEvents

Updaterec.filenum = 4

ClearUpdatebuf
nStatus = BTRCALL(BGETFIRST, ShopPosBlock, Updatebuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)

Do While nStatus = 0 And UpdateFileLoaded = True

    For x = 1 To Len(Shopdatabuf)
        Updaterec.Data(x) = Updatebuf.Data(x)
    Next
    
    nStatus = BTRCALL(BINSERT, UpdatePosBlock, Updaterec, Len(Updaterec), ByVal UpdateKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Update Insert Error: " & BtrieveErrorCode(nStatus)
        nYesNo = MsgBox("Do you want to stop the process?", vbYesNo, "Stop update file creation?")
        If nYesNo = vbYes Then
            StopUpdate
            Exit Sub
        End If
    End If
    
    'ClearUpdatebuf
    
    nStatus = BTRCALL(BGETNEXT, ShopPosBlock, Updatebuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
    
    Updaterec.recnumber = Updaterec.recnumber + 1
    recnum = recnum + 1
    frmProgressBar.lblPanel(1).Caption = recnum
    Call frmProgressBar.IncreaseProgress
    If Not bUseCPU Then DoEvents

Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    nYesNo = MsgBox("Error exporting Shops, Btrieve Error: " & BtrieveErrorCode(nStatus, True) _
        & vbCrLf & "Continue?", vbYesNo + vbDefaultButton2 + vbQuestion)
    If nYesNo = vbNo Then Call StopUpdate
End If

End Sub

Private Sub InsertMonsters()
Dim nStatus As Integer, nYesNo As Integer, x As Long

frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_KNMSR
recnum = 1
DoEvents

Updaterec.filenum = 7

ClearUpdatebuf
nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Updatebuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)

Do While nStatus = 0 And UpdateFileLoaded = True
    
    For x = 1 To Len(Monsterdatabuf)
        Updaterec.Data(x) = Updatebuf.Data(x)
    Next
    
    nStatus = BTRCALL(BINSERT, UpdatePosBlock, Updaterec, Len(Updaterec), ByVal UpdateKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Update Insert Error: " & BtrieveErrorCode(nStatus)
        nYesNo = MsgBox("Do you want to stop the process?", vbYesNo, "Stop update file creation?")
            If nYesNo = vbYes Then
                StopUpdate
                Exit Sub
            End If
    End If
    
    'ClearUpdatebuf
    
    nStatus = BTRCALL(BGETNEXT, MonsterPosBlock, Updatebuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
    
    Updaterec.recnumber = Updaterec.recnumber + 1
    recnum = recnum + 1
    frmProgressBar.lblPanel(1).Caption = recnum
    Call frmProgressBar.IncreaseProgress
    If Not bUseCPU Then DoEvents
    
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    nYesNo = MsgBox("Error exporting Monsters, Btrieve Error: " & BtrieveErrorCode(nStatus, True) _
        & vbCrLf & "Continue?", vbYesNo + vbDefaultButton2 + vbQuestion)
    If nYesNo = vbNo Then Call StopUpdate
End If

End Sub

Private Sub InsertRooms()
Dim nStatus As Integer, nYesNo As Integer, x As Long

frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_MP
recnum = 1
DoEvents

Updaterec.filenum = 3

ClearUpdatebuf
nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Updatebuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)

Do While nStatus = 0 And UpdateFileLoaded = True
    
    For x = 1 To Len(Roomdatabuf)
        Updaterec.Data(x) = Updatebuf.Data(x)
    Next
    
    nStatus = BTRCALL(BINSERT, UpdatePosBlock, Updaterec, Len(Updaterec), ByVal UpdateKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Update Insert Error: " & BtrieveErrorCode(nStatus)
        nYesNo = MsgBox("Do you want to stop the process?", vbYesNo, "Stop update file creation?")
        If nYesNo = vbYes Then
            StopUpdate
            Exit Sub
        End If
    End If
    
    'ClearUpdatebuf
    
    nStatus = BTRCALL(BGETNEXT, RoomPosBlock, Updatebuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    
    Updaterec.recnumber = Updaterec.recnumber + 1
    recnum = recnum + 1
    frmProgressBar.lblPanel(1).Caption = recnum
    Call frmProgressBar.IncreaseProgress
    If Not bUseCPU Then DoEvents

Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    nYesNo = MsgBox("Error exporting Rooms, Btrieve Error: " & BtrieveErrorCode(nStatus, True) _
        & vbCrLf & "Continue?", vbYesNo + vbDefaultButton2 + vbQuestion)
    If nYesNo = vbNo Then Call StopUpdate
End If

End Sub

Private Sub InsertActions()
Dim nStatus As Integer, nYesNo As Integer, x As Long

frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & "acts.dat"
recnum = 1
DoEvents

Updaterec.filenum = 10

ClearUpdatebuf
nStatus = BTRCALL(BGETFIRST, ActionPosBlock, Updatebuf, Len(ActionDatabuf), ByVal ActionKeyBuffer, KEY_BUF_LEN, 0)

Do While nStatus = 0 And UpdateFileLoaded = True
    
    For x = 1 To Len(ActionDatabuf)
        Updaterec.Data(x) = Updatebuf.Data(x)
    Next
    
    nStatus = BTRCALL(BINSERT, UpdatePosBlock, Updaterec, Len(Updaterec), ByVal UpdateKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Update Insert Error: " & BtrieveErrorCode(nStatus)
        nYesNo = MsgBox("Do you want to stop the process?", vbYesNo, "Stop update file creation?")
        If nYesNo = vbYes Then
            StopUpdate
            Exit Sub
        End If
    End If
    
    'ClearUpdatebuf
    
    nStatus = BTRCALL(BGETNEXT, ActionPosBlock, Updatebuf, Len(ActionDatabuf), ByVal ActionKeyBuffer, KEY_BUF_LEN, 0)
    
    Updaterec.recnumber = Updaterec.recnumber + 1
    recnum = recnum + 1
    frmProgressBar.lblPanel(1).Caption = recnum
    Call frmProgressBar.IncreaseProgress
    If Not bUseCPU Then DoEvents

Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    nYesNo = MsgBox("Error exporting Actions, Btrieve Error: " & BtrieveErrorCode(nStatus, True) _
        & vbCrLf & "Continue?", vbYesNo + vbDefaultButton2 + vbQuestion)
    If nYesNo = vbNo Then Call StopUpdate
End If

End Sub

Private Sub ClearUpdatebuf() 'nSize As Integer)
Dim x As Integer

'If nSize < 0 Then
    For x = 1 To Len(Updatebuf)
        Updatebuf.Data(x) = &H0 'Asc(vbNullChar)
    Next
'Else
'    For x = 1 To nSize
'        Updatebuf.Data(x) = &H0 'Asc(vbNullChar)
'    Next
'End If

End Sub
Public Sub StopUpdate()
Dim nStatus As Integer

If UpdateFileLoaded = True Then
    nStatus = BTRCALL(BCLOSE, UpdatePosBlock, Updaterec, Len(Updaterec), ByVal UpdateKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then MsgBox "BCLOSE, Update file Error: " & BtrieveErrorCode(nStatus)
    UpdateFileLoaded = False
End If

frmMain.Enabled = True
frmMain.WindowState = InitialWindowState

frmProgressBar.lblNote.Visible = False
Unload frmProgressBar

End Sub
Private Sub CreateUpdateFile()
Dim WGPath As String, nStatus As Integer
Dim fso As FileSystemObject, fil1 As File

Set fso = CreateObject("Scripting.FileSystemObject")

WGPath = ReadINI("Settings", "WGPath" & IIf(WorksWithN = True, "_n", IIf(WorksWithWG = True, "_wg", "")))
UpdateKeyBuffer = WGPath & "w" & strDatCallLetters & strDatSuffix_UPDAT

If fso.FileExists(UpdateKeyBuffer) = True Then
    Set fil1 = fso.GetFile(UpdateKeyBuffer)
    fil1.Delete True
End If

Set fso = Nothing
Set fil1 = Nothing

frmProgressBar.lblPanel(1).Caption = "Creating w" & strDatCallLetters & strDatSuffix_UPDAT

nStatus = InitUpdateFile
If Not nStatus = 0 Then
    MsgBox "Error creating update file: " & BtrieveErrorCode(nStatus)
    StopUpdate
    Exit Sub
End If

UpdateKeyBuffer = WGPath & "w" & strDatCallLetters & strDatSuffix_UPDAT
frmProgressBar.lblPanel(1).Caption = "Opening w" & strDatCallLetters & strDatSuffix_UPDAT

nStatus = BTRCALL(BOPEN, UpdatePosBlock, Updaterec, UpdateDataBufSize, ByVal UpdateKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error opening update file: " & BtrieveErrorCode(nStatus)
    StopUpdate
    Exit Sub
Else
    UpdateFileLoaded = True
End If

Set fso = Nothing
Set fil1 = Nothing

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

nStatus = BTRCALL(BSTAT, SpellPosBlock, DBStatDatabuf, Len(Spelldatabuf), 0, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    CalcTotalRecords = CalcTotalRecords + 1300
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
    
nStatus = BTRCALL(BSTAT, ActionPosBlock, DBStatDatabuf, Len(ActionDatabuf), 0, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    CalcTotalRecords = CalcTotalRecords + 100
Else
    DBStatRowToStruct DBStatDatabuf.buf
    CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
End If

nStatus = BTRCALL(BSTAT, MonsterPosBlock, DBStatDatabuf, Len(Monsterdatabuf), 0, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    CalcTotalRecords = CalcTotalRecords + 1100
Else
    DBStatRowToStruct DBStatDatabuf.buf
    CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
End If

nStatus = BTRCALL(BSTAT, TextblockPosBlock, DBStatDatabuf, Len(TextblockDataBuf), 0, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    CalcTotalRecords = CalcTotalRecords + 2600
Else
    DBStatRowToStruct DBStatDatabuf.buf
    CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
End If

nStatus = BTRCALL(BSTAT, MessagePosBlock, DBStatDatabuf, Len(Messagedatabuf), 0, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    CalcTotalRecords = CalcTotalRecords + 3700
Else
    DBStatRowToStruct DBStatDatabuf.buf
    CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
End If

nStatus = BTRCALL(BSTAT, RacePosBlock, DBStatDatabuf, Len(Racedatabuf), 0, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    CalcTotalRecords = CalcTotalRecords + 30
Else
    DBStatRowToStruct DBStatDatabuf.buf
    CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
End If

nStatus = BTRCALL(BSTAT, ClassPosBlock, DBStatDatabuf, Len(Classdatabuf), 0, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    CalcTotalRecords = CalcTotalRecords + 30
Else
    DBStatRowToStruct DBStatDatabuf.buf
    CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
End If

nStatus = BTRCALL(BSTAT, RoomPosBlock, DBStatDatabuf, Len(Roomdatabuf), 0, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    CalcTotalRecords = CalcTotalRecords + 30000
Else
    DBStatRowToStruct DBStatDatabuf.buf
    CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
End If

If CalcTotalRecords <= 0 Then CalcTotalRecords = 1

Exit Function

error:
Call HandleError
End Function


