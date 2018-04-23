VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDatabaseMerge 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Merge Users"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   Icon            =   "frmDatabaseMerge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5730
   Begin VB.TextBox txtSingleUser 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2220
      MaxLength       =   30
      TabIndex        =   15
      Top             =   1860
      Width           =   2115
   End
   Begin VB.CheckBox chkSingleUser 
      Caption         =   "Single User BBS ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CheckBox chkPrompt 
      Caption         =   "Prompt on Duplicate Records"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.ComboBox cmbDBType 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmDatabaseMerge.frx":08CA
      Left            =   3060
      List            =   "frmDatabaseMerge.frx":08F2
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   1755
   End
   Begin MSComctlLib.StatusBar stsStatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   3180
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7488
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Left            =   4560
      TabIndex        =   12
      Top             =   2700
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   315
      Left            =   4560
      TabIndex        =   11
      Top             =   2340
      Width           =   1095
   End
   Begin VB.CommandButton cmdBrowseDest 
      Caption         =   "Browse ..."
      Height          =   315
      Left            =   4560
      TabIndex        =   9
      Top             =   1380
      Width           =   1095
   End
   Begin VB.CommandButton cmdBrowseSource 
      Caption         =   "Browse ..."
      Height          =   315
      Left            =   4560
      TabIndex        =   6
      Top             =   660
      Width           =   1095
   End
   Begin VB.TextBox txtDest 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   4215
   End
   Begin VB.TextBox txtSource 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "Database Type:"
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
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblNote 
      Caption         =   $"frmDatabaseMerge.frx":095E
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   2340
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Database to import INTO:"
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
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Database to import FROM:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   4215
   End
End
Attribute VB_Name = "frmDatabaseMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Private Type SourcePosBlockType
    buf(1 To 128) As Byte
End Type
Dim SourcePosBlock As SourcePosBlockType

Private Type DestPosBlockType
    buf(1 To 128) As Byte
End Type

Dim DestPosBlock As DestPosBlockType

Dim SourceKeyBuffer As String * 255
Dim DestKeyBuffer As String * 255
Dim Source As String
Dim Dest As String
Dim LogFile As String
Dim ts As TextStream

Private Sub chkSingleUser_Click()
If chkSingleUser.Value = 1 Then
    txtSingleUser.Enabled = True
Else
    txtSingleUser.Enabled = False
End If
End Sub

Private Sub Form_Load()
'cmbDBType.ListIndex = 11
Me.Show
Me.SetFocus
cmdClose.SetFocus
End Sub

Private Sub cmdStart_Click()
On Error GoTo error:
Dim fso As FileSystemObject, nYesNo As Integer, temp As Boolean

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

UnloadForms (Me.Name)
DoEvents

Set fso = CreateObject("Scripting.FileSystemObject")

Source = txtSource.Text
Dest = txtDest.Text

If fso.FileExists(Source) Then
    Source = GetShortName(Source)
Else
    MsgBox "Source file: " & vbCrLf & Source & vbCrLf & "was not found."
    Exit Sub
End If

If fso.FileExists(Dest) Then
    Dest = GetShortName(Dest)
Else
    MsgBox "Destination file: " & vbCrLf & Dest & vbCrLf & "was not found."
    Exit Sub
End If

SourceKeyBuffer = Source & Chr(0)
DestKeyBuffer = Dest & Chr(0)

frmMain.Enabled = False

temp = OpenSourceDB
If temp = False Then GoTo close1:

temp = OpenDestDB
If temp = False Then GoTo close1:

temp = CheckSame
If temp = False Then GoTo close1:

If Right(App.Path, 1) = "\" Then
    LogFile = App.Path & "NMR-Log_Merge.txt"
Else
    LogFile = App.Path & "\" & "NMR-Log_Merge.txt"
End If

If fso.FileExists(LogFile) Then fso.DeleteFile (LogFile), True

DoEvents
Set ts = fso.OpenTextFile(LogFile, ForWriting, True)
ts.WriteLine ("Merge job started " & Date & " @ " & Time)
ts.WriteLine ("Merging " & cmbDBType.Text)
ts.WriteLine ("Source: " & Source)
ts.WriteLine ("Dest: " & Dest)
ts.WriteBlankLines 1

Set fso = Nothing

'Select Case cmbDBType.ListIndex
'    Case 0: Call MergeActions
'     Case 1: temp = MergeBanks
'    Case 2: Call MergeClasses
'    Case 3: Call MergeItems
'    Case 4: Call MergeMessages
'    Case 5: Call MergeMonsters
'    Case 6: Call MergeRaces
'    Case 7: Call MergeRooms
'    Case 8: Call MergeShops
'    Case 9: Call MergeSpells
'    Case 10: Call MergeTextblocks
'     Case 11: temp = MergeUsers
'End Select

temp = MergeUsers

If temp = True Then
    ts.WriteBlankLines 1
    ts.WriteLine ("Finished -- " & Date & " @ " & Time)
    ts.Close
    
    DoEvents
    nYesNo = MsgBox("Merge Complete, view log?", vbYesNo, "Merge Complete")
    If nYesNo = vbYes Then Call cmdLog_Click
Else
    ts.WriteBlankLines 1
    ts.WriteLine ("Error encountered -- " & Date & " @ " & Time)
    ts.Close
    
    DoEvents
End If

close1:
frmMain.Enabled = True
Call CloseOut

Exit Sub
error:
Call HandleError
Call CloseOut
frmMain.Enabled = True
Set fso = Nothing
End Sub
Private Function CheckSame() As Boolean
Dim nStatus As Integer, RecLen As Long

CheckSame = False

nStatus = BTRCALL(BSTAT, SourcePosBlock, DBStatDatabuf, Len(DBStatDatabuf), ByVal SourceKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Unable to retrieve source database stats, Error: " & BtrieveErrorCode(nStatus), vbExclamation
    Exit Function
End If
Call DBStatRowToStruct(DBStatDatabuf.buf)
RecLen = DBStat.RecLen

nStatus = BTRCALL(BSTAT, DestPosBlock, DBStatDatabuf, Len(DBStatDatabuf), ByVal DestKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Unable to retrieve Destination database stats, Error: " & BtrieveErrorCode(nStatus), vbExclamation
    Exit Function
End If
Call DBStatRowToStruct(DBStatDatabuf.buf)

If RecLen <> DBStat.RecLen Then
    MsgBox "Source (" & RecLen & ") and destination (" & DBStat.RecLen & ") database record lengths do not match!", vbExclamation
    Exit Function
End If

CheckSame = True

End Function

Private Sub cmdBrowseDest_Click()
Dim fso As FileSystemObject

Set fso = CreateObject("Scripting.FileSystemObject")

CommonDialog1.Filter = "User Dat Files (*user?.dat)|*user?.dat|All Files (*.*)|*.*"
CommonDialog1.DialogTitle = "Select Destination Database..."
CommonDialog1.InitDir = ReadINI("Settings", "WGPath" & IIf(WorksWithN = True, "_n", IIf(WorksWithWG = True, "_wg", "")))

On Error GoTo canceled:
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then GoTo canceled:

If fso.FileExists(CommonDialog1.FileName) Then
    txtDest.Text = GetShortName(CommonDialog1.FileName)
End If

canceled:

Set fso = Nothing
End Sub

Private Sub cmdBrowseSource_Click()
Dim fso As FileSystemObject

Set fso = CreateObject("Scripting.FileSystemObject")

CommonDialog1.Filter = "User Dat Files (*user?.dat)|*user?.dat|All Files (*.*)|*.*"
CommonDialog1.DialogTitle = "Select Source Database..."

On Error GoTo canceled:
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then GoTo canceled:

If fso.FileExists(CommonDialog1.FileName) Then
    txtSource.Text = GetShortName(CommonDialog1.FileName)
End If

canceled:

Set fso = Nothing
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdLog_Click()
Dim fso As FileSystemObject
Dim LogFile As String

Set fso = CreateObject("Scripting.FileSystemObject")

If Right(App.Path, 1) = "\" Then
    LogFile = App.Path & "NMR-Log_Merge.txt"
Else
    LogFile = App.Path & "\" & "NMR-Log_Merge.txt"
End If

If fso.FileExists(LogFile) Then
    Call ShellExecute(0&, "open", LogFile, vbNullString, vbNullString, vbNormalFocus)
Else
    MsgBox LogFile & " does not exist."
End If

Set fso = Nothing
End Sub

Private Sub CloseOut()
On Error Resume Next
Dim nStatus As Integer

nStatus = BTRCALL(BCLOSE, SourcePosBlock, 0, 0, 0, 0, 0)
nStatus = BTRCALL(BCLOSE, DestPosBlock, 0, 0, 0, 0, 0)

End Sub
Private Function OpenSourceDB() As Boolean
Dim nStatus As Integer

OpenSourceDB = False

nStatus = BTRCALL(BOPEN, SourcePosBlock, 0, 0, ByVal SourceKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then MsgBox "Unable to open source database, Error: " & BtrieveErrorCode(nStatus): Exit Function

OpenSourceDB = True

End Function
Private Function OpenDestDB() As Boolean
Dim nStatus As Integer
OpenDestDB = False

nStatus = BTRCALL(BOPEN, DestPosBlock, 0, 0, ByVal DestKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then MsgBox "Unable to open destination database, Error: " & BtrieveErrorCode(nStatus): Exit Function

OpenDestDB = True

End Function
Private Function MergeUsers() As Boolean
On Error GoTo error:
Dim nStatusS As Integer, nStatusD As Integer
Dim recnum As Long, nYesNo As Integer
Dim BBSName As String, FirstName As String, bChg As Boolean

MergeUsers = False

recnum = 1
stsStatusBar.Panels(1).Text = "Merging..."
stsStatusBar.Panels(2).Text = recnum

DoEvents
nStatusS = BTRCALL(BGETFIRST, SourcePosBlock, Userdatabuf, Len(Userdatabuf), ByVal SourceKeyBuffer, KEY_BUF_LEN, 0)
If nStatusS <> 0 Then
    MsgBox "Unable to retrieve first record from source database, Error: " & BtrieveErrorCode(nStatusS)
    ts.WriteLine ("Error getting first record")
    GoTo close1:
End If

ts.WriteLine ("BBS Name (Mud Name)")
ts.WriteBlankLines (1)

If chkSingleUser.Value = 1 Then
    BBSName = txtSingleUser.Text
    BBSName = BBSName & String(30 - Len(BBSName), vbNullChar)
    
    nStatusS = BTRCALL(BGETEQUAL, SourcePosBlock, Userdatabuf, Len(Userdatabuf), ByVal BBSName, KEY_BUF_LEN, 0)
    
    If Not nStatusS = 0 Then
        MsgBox "Error on BGETGEQUAL: " & BtrieveErrorCode(nStatusS)
        GoTo close1:
    End If
End If

Do While nStatusS <> 9
    UserRowToStruct Userdatabuf.buf
    bChg = False
again:
    BBSName = Userrec.BBSName
    If Not InStr(1, Userrec.BBSName, vbNullChar) = 0 Then BBSName = Left(Userrec.BBSName, InStr(1, Userrec.BBSName, vbNullChar) - 1)
    FirstName = Userrec.FirstName
    If Not InStr(1, Userrec.FirstName, vbNullChar) = 0 Then FirstName = Left(Userrec.FirstName, InStr(1, Userrec.FirstName, vbNullChar) - 1)

    nStatusD = BTRCALL(BINSERT, DestPosBlock, Userdatabuf, Len(Userdatabuf), ByVal DestKeyBuffer, KEY_BUF_LEN, 0)
    If nStatusD <> 0 Then
        If nStatusD = 5 Then
            If chkPrompt.Value = 1 Then
                nYesNo = MsgBox("BBS Account '" & BBSName & ",' with mud name '" & FirstName _
                    & "' already exists, change?" & vbCrLf & "(Both BBS Name and First Name must be different than any other.)" _
                    & vbCrLf & vbCrLf & "Click cancel to stop further prompting.", vbYesNoCancel + vbQuestion + vbDefaultButton1)
                If nYesNo = vbYes Then
redobbs:
                    BBSName = InputBox("Enter BBS Name" & vbCrLf & vbCrLf & "It is recommended not to use numbers.", "Enter BBS Name", BBSName)
                    If Len(BBSName) > Len(Userrec.BBSName) Then MsgBox "BBS Name cannot be longer than " & Len(Userrec.BBSName) & " characters.": GoTo redobbs:
redofirst:
                    FirstName = InputBox("Enter First Name" & vbCrLf & vbCrLf & "It is recommended not to use numbers.", "Enter First Name", FirstName)
                    If Len(FirstName) > Len(Userrec.FirstName) Then MsgBox "First Name cannot be longer than " & Len(Userrec.FirstName) & " characters.": GoTo redofirst:
                    
                    If BBSName = "" Or FirstName = "" Then
                        
                        BBSName = Userrec.BBSName
                        If Not InStr(1, Userrec.BBSName, vbNullChar) = 0 Then BBSName = Left(Userrec.BBSName, InStr(1, Userrec.BBSName, vbNullChar) - 1)
                        FirstName = Userrec.FirstName
                        If Not InStr(1, Userrec.FirstName, vbNullChar) = 0 Then FirstName = Left(Userrec.FirstName, InStr(1, Userrec.FirstName, vbNullChar) - 1)
                        
                        ts.WriteLine BBSName & " (" & FirstName & ") -- Record exists"
                    Else
                        
                        BBSName = Trim(BBSName)
                        If Len(BBSName) < Len(Userrec.BBSName) Then BBSName = Trim(BBSName & String(Len(Userrec.BBSName) - Len(BBSName), Chr(0)))
                        FirstName = Trim(RemoveCharacter(FirstName, " "))
                        If Len(FirstName) < Len(Userrec.FirstName) Then FirstName = FirstName & String(Len(Userrec.FirstName) - Len(FirstName), Chr(0))
                        
                        Userrec.BBSName = BBSName
                        Userrec.FirstName = FirstName
                        
                        UserStructToRow Userdatabuf.buf
                        bChg = True
                        GoTo again:
                    End If
                Else
                    If nYesNo = vbCancel Then chkPrompt.Value = 0
                    DoEvents
                    ts.WriteLine BBSName & " (" & FirstName & ") -- Record exists, skipped."
                End If
            Else
                ts.WriteLine BBSName & " (" & FirstName & ") -- Record exists, skipped."
            End If
        Else
            ts.WriteLine BBSName & " (" & FirstName & ") -- Insert Error: " & BtrieveErrorCode(nStatusD)
        End If
    Else
        If bChg = True Then
            ts.WriteLine BBSName & " (" & FirstName & ") -- Insert Successful *changed*"
        Else
            ts.WriteLine BBSName & " (" & FirstName & ") -- Insert Successful"
        End If
    End If
    
    If chkSingleUser.Value = 1 Then GoTo done:
    
    nStatusS = BTRCALL(BGETNEXT, SourcePosBlock, Userdatabuf, Len(Userdatabuf), ByVal SourceKeyBuffer, KEY_BUF_LEN, 0)
    recnum = recnum + 1
    stsStatusBar.Panels(2).Text = recnum
    DoEvents
Loop

done:
MergeUsers = True

close1:
stsStatusBar.Panels(1).Text = ""
stsStatusBar.Panels(2).Text = ""

Exit Function
error:
Call HandleError
stsStatusBar.Panels(1).Text = ""
stsStatusBar.Panels(2).Text = ""

End Function


Private Sub Form_Unload(Cancel As Integer)
Set ts = Nothing
End Sub

Private Sub txtDest_GotFocus()
Call SelectAll(txtDest)

End Sub

Private Sub txtSource_GotFocus()
Call SelectAll(txtSource)

End Sub
