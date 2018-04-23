VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   8745
   Begin VB.Frame Frame2 
      Caption         =   "Previous Saves:"
      Height          =   675
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5595
      Begin VB.ComboBox cmbPreviousSaves 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   5355
      End
   End
   Begin VB.CommandButton cmdRecreateINI 
      Caption         =   "Recreate settings.ini"
      Height          =   375
      Left            =   3300
      TabIndex        =   25
      Top             =   4140
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      Caption         =   "Settings"
      Height          =   3975
      Left            =   5760
      TabIndex        =   6
      Top             =   60
      Width           =   2895
      Begin VB.CheckBox chkOppositeListOrder 
         Caption         =   "Load Lists in Opposite Order (Highest Record # First)"
         Height          =   375
         Left            =   180
         TabIndex        =   27
         Top             =   3480
         Width           =   2430
      End
      Begin VB.CheckBox chkOnlyNames 
         Caption         =   "Only Load # && Name Columns (Speedier on Editors w/Lists)"
         Height          =   435
         Left            =   180
         TabIndex        =   20
         Top             =   2100
         Width           =   2490
      End
      Begin VB.CheckBox chkTaskHide 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   1605
         TabIndex        =   16
         Top             =   1440
         Width           =   210
      End
      Begin VB.TextBox txtTaskDelay 
         Height          =   285
         Left            =   1620
         MaxLength       =   1
         TabIndex        =   18
         Top             =   1740
         Width           =   345
      End
      Begin VB.ComboBox cmbTaskPos 
         Height          =   315
         ItemData        =   "frmSettings.frx":08CA
         Left            =   1620
         List            =   "frmSettings.frx":08D7
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1080
         Width           =   1155
      End
      Begin VB.CheckBox chkAutoMonsterIndex 
         Caption         =   "Create Monster Index on Load"
         Height          =   255
         Left            =   180
         TabIndex        =   21
         Top             =   2580
         Width           =   2490
      End
      Begin VB.CheckBox chkUseCPU 
         Caption         =   "Use all Avail. CPU on Jobs"
         Height          =   255
         Left            =   180
         TabIndex        =   22
         Top             =   2880
         Width           =   2370
      End
      Begin VB.ComboBox cmbVersion 
         Height          =   315
         ItemData        =   "frmSettings.frx":08ED
         Left            =   1380
         List            =   "frmSettings.frx":08EF
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   180
         Width           =   1395
      End
      Begin VB.CheckBox chkAutoCompile 
         Caption         =   "Auto Compile Update on exit"
         Height          =   255
         Left            =   180
         TabIndex        =   23
         Top             =   3180
         Width           =   2370
      End
      Begin VB.TextBox txtDatCallLetters 
         Height          =   285
         Left            =   1620
         MaxLength       =   2
         TabIndex        =   11
         Top             =   660
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "(seconds)"
         Height          =   195
         Index           =   2
         Left            =   2040
         TabIndex        =   19
         Top             =   1800
         Width           =   690
      End
      Begin VB.Label lblAutoHide 
         Caption         =   "Taskbar Autohide:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Autohide Delay:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   1740
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Taskbar Position:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   1140
         Width           =   1275
      End
      Begin VB.Label Label4 
         Caption         =   "Dat File Version:"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Dat Call Letters:"
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "item?.dat"
         Height          =   255
         Left            =   2025
         TabIndex        =   12
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label3 
         Caption         =   "w"
         Height          =   255
         Left            =   1425
         TabIndex        =   10
         Top             =   720
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7380
      TabIndex        =   26
      Top             =   4140
      Width           =   1275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   60
      TabIndex        =   24
      Top             =   4140
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datfile Path"
      Height          =   3255
      Left            =   60
      TabIndex        =   2
      Top             =   780
      Width           =   5595
      Begin VB.FileListBox File1 
         Height          =   2820
         Left            =   3120
         Pattern         =   "w*.dat"
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
      Begin VB.DirListBox Dir1 
         Height          =   2565
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2895
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit


Public bCancelUnload As Boolean
'psinfo format
'1-5 (ps#), 1-4 (settings), 1 - actual value
'1-5 (ps#), 1-4 (settings), 2 - user friendly value
Public bReload As Boolean
Dim bLoaded As Boolean
Dim PSInfo(1 To 5, 1 To 4, 1 To 2) As String


Private Sub cmbVersion_Click()
bReload = True
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim fso As FileSystemObject, oFolder As Folder
Dim temp As String, x As Integer, z1 As Integer, z2 As Integer

cmbVersion.clear '***NEWMUDVER***
If Not WorksWithWG Then
    cmbVersion.AddItem "v1.11h", 0
    cmbVersion.AddItem "v1.11i", 1
    cmbVersion.AddItem "v1.11J", 2
    cmbVersion.AddItem "v1.11k", 3
    cmbVersion.AddItem "v1.11L", 4
    cmbVersion.AddItem "v1.11m", 5
    cmbVersion.AddItem "v1.11n", 6
    cmbVersion.AddItem "v1.11o", 7
    cmbVersion.AddItem "v1.11p-b13", 8
    cmbVersion.AddItem "v1.11p", 9
Else
    cmbVersion.AddItem "v1.11p-WG", 0
End If

bLoaded = False
Set fso = CreateObject("Scripting.FileSystemObject")

temp = ReadINI("Settings", "WGPath" & IIf(WorksWithN = True, "_n", IIf(WorksWithWG = True, "_wg", "")))
If fso.FolderExists(temp) = False Then
    Drive1.Drive = App.Path
    Dir1.Path = App.Path
Else
    Drive1.Drive = temp
    Dir1.Path = GetLongDirName(temp)
End If

txtDatCallLetters.Text = ReadINI("Settings", "DatCallLetters" & IIf(WorksWithN = True, "_n", IIf(WorksWithWG = True, "_wg", "")))
chkAutoCompile.Value = Val(ReadINI("Settings", "AutoCompile"))
chkUseCPU.Value = Val(ReadINI("Settings", "UseCPU"))
cmbVersion.ListIndex = eDatFileVersion
chkAutoMonsterIndex.Value = Val(ReadINI("Settings", "AutoMonsterIndex"))
chkTaskHide.Value = Val(ReadINI("Settings", "TaskBarAutoHide"))
cmbTaskPos.ListIndex = Val(ReadINI("Settings", "TaskBarPos"))
txtTaskDelay.Text = Val(ReadINI("Settings", "TaskBarDelay"))
chkOnlyNames.Value = Val(ReadINI("Settings", "OnlyLoadNames"))
chkOppositeListOrder.Value = Val(ReadINI("Settings", "OppositeListOrder"))

If Val(txtTaskDelay.Text) < 1 Then txtTaskDelay.Text = 1

Call GetPSInfo
cmbPreviousSaves.clear
For z1 = 1 To 5
    If PSInfo(z1, 1, 1) = "0" Or PSInfo(z1, 1, 1) = "" Then
        cmbPreviousSaves.AddItem (z1 & ". none")
    Else
        temp = ""
        For z2 = 1 To 3 'default 4 ... 3 for no autocompile check
            If Not z2 = 1 Then temp = temp & ", "
            temp = temp & PSInfo(z1, z2, 2)
        Next
        cmbPreviousSaves.AddItem (z1 & ". " & temp)
    End If
Next
If Not cmbPreviousSaves.ListCount = 0 Then cmbPreviousSaves.ListIndex = 0

Me.Left = frmMain.Width / 8
Me.Top = frmMain.Height / 8
Me.Show
Me.SetFocus
cmdCancel.SetFocus

bReload = False
bLoaded = True
Set fso = Nothing

End Sub


Private Sub chkUseCPU_Click()

If bLoaded = False Or bUseCPU = True Then Exit Sub
If chkUseCPU.Value = 1 Then MsgBox "NOTE: This option will make most large database operations faster." & vbCrLf _
    & "However, during those operations the program may appear to hang," & vbCrLf _
    & "yet it's really just using too much CPU to update the window.", vbInformation
    
End Sub

Private Sub cmbPreviousSaves_Click()
Dim fso As FileSystemObject
Dim nSel As Integer, sPath As String, nVer As Integer, sCall As String ', nAC As Integer

On Error GoTo error:

If bLoaded = False Then Exit Sub

Set fso = CreateObject("Scripting.FileSystemObject")

nSel = cmbPreviousSaves.ListIndex + 1

sPath = PSInfo(nSel, 1, 1)
nVer = Val(PSInfo(nSel, 2, 1))
sCall = PSInfo(nSel, 3, 1)
'nAC = Val(PSInfo(nSel, 4, 1))

If sPath = "" Or sPath = "0" Then GoTo Skip:

If fso.FolderExists(sPath) Then Dir1.Path = GetLongDirName(sPath)
If Not nVer > cmbVersion.ListCount Then cmbVersion.ListIndex = nVer
txtDatCallLetters.Text = sCall
'If nAC = 0 Or nAC = 1 Then chkAutoCompile.value = nAC

Skip:
Set fso = Nothing

out:
Exit Sub
error:
Call HandleError("cmbPreviousSaves_Click")
Resume out:
End Sub


Private Sub GetPSInfo()
On Error GoTo error:
Dim x As Integer, temp As String, sStr As String
Dim y1 As Integer, y2 As Integer 'positions in string
Dim z1 As Integer, z2 As Integer 'psinfo array variables

For z1 = 1 To 5
    For z2 = 1 To 4
        PSInfo(z1, z2, 1) = "0"
        PSInfo(z1, z2, 2) = "0"
    Next
Next
    
For z1 = 1 To 5
    
    sStr = ReadINI("Settings", "PS" & z1)
    
    'path
    z2 = 1 'field
    y1 = 0 'starting point
    y2 = InStr(y1 + 1, sStr, ";") 'finishing point
    If y2 = 0 Then GoTo Skip:
    temp = Mid(sStr, y1 + 1, y2 - y1 - 1)
    PSInfo(z1, z2, 1) = temp
    If Len(temp) > 25 Then
        PSInfo(z1, z2, 2) = " ..." & Right(temp, 25)
    Else
        PSInfo(z1, z2, 2) = temp
    End If
    
    'datversion
    z2 = z2 + 1 'field
    y1 = y2 'starting point
    y2 = InStr(y1 + 1, sStr, ";") 'finishing point
    If y2 = 0 Then GoTo Skip:
    temp = Mid(sStr, y1 + 1, y2 - y1 - 1)
    PSInfo(z1, z2, 1) = temp
    PSInfo(z1, z2, 2) = FriendlyDatVersion(Val(temp))
    
    'callletters
    z2 = z2 + 1 'field
    y1 = y2 'starting point
    y2 = InStr(y1 + 1, sStr, ";") 'finishing point
    If y2 = 0 Then GoTo Skip:
    temp = Mid(sStr, y1 + 1, y2 - y1 - 1)
    PSInfo(z1, z2, 1) = temp
    PSInfo(z1, z2, 2) = temp

'    'autocompile
'    z2 = z2 + 1 'field
'    y1 = y2 'starting point
'    y2 = InStr(y1 + 1, sStr, ";") 'finishing point
'    If y2 = 0 Then GoTo Skip:
'    temp = Mid(sStr, y1 + 1, y2 - y1 - 1)
'    PSInfo(z1, z2, 1) = temp
'    If Val(temp) = 1 Then
'        PSInfo(z1, z2, 2) = "AC-Yes"
'    Else
'        PSInfo(z1, z2, 2) = "AC-No"
'    End If

Skip:
Next

Exit Sub
error:
Call HandleError
End Sub

Private Sub cmdCancel_Click()
If bKeepSettingsOpen Then
    MsgBox "Please correct your settings and click save.", vbInformation
    Exit Sub
End If
Unload Me
End Sub

Private Sub cmdRecreateINI_Click()
On Error GoTo error:

Call modSettings.CreateSettings
'Call cmdSave_Click
bKeepSettingsOpen = False
Call ReloadApp
Call Form_Load
bReload = True

out:
Exit Sub
error:
Call HandleError("cmdRecreateINI_Click")
Resume out:

End Sub

Private Sub cmdSave_Click()
Dim fso As FileSystemObject, fldr1 As Folder, sStr As String
Dim x As Integer, sPS(1 To 5) As String, sPStmp(1 To 5) As String

On Error GoTo error:

Set fso = CreateObject("Scripting.FileSystemObject")
Set fldr1 = fso.GetFolder(Dir1.Path)

sStr = fldr1.ShortPath
If Not Right(sStr, 1) = "\" Then sStr = sStr & "\"

Call WriteINI("Settings", "FirstRun" & IIf(WorksWithN = True, "_n", IIf(WorksWithWG = True, "_wg", "")), 1)

Call WriteINI("Settings", "WGPath" & IIf(WorksWithN = True, "_n", IIf(WorksWithWG = True, "_wg", "")), sStr)
Call WriteINI("Settings", "DatCallLetters" & IIf(WorksWithN = True, "_n", IIf(WorksWithWG = True, "_wg", "")), txtDatCallLetters.Text)
Call WriteINI("Settings", "AutoCompile", chkAutoCompile.Value)
Call WriteINI("Settings", "eDatFileVersion" & IIf(WorksWithN = True, "_n", IIf(WorksWithWG = True, "_wg", "")), cmbVersion.ListIndex)
Call WriteINI("Settings", "UseCPU", chkUseCPU.Value)
Call WriteINI("Settings", "AutoMonsterIndex", chkAutoMonsterIndex.Value)
Call WriteINI("Settings", "OnlyLoadNames", chkOnlyNames.Value)
Call WriteINI("Settings", "OppositeListOrder", chkOppositeListOrder.Value)

Call WriteINI("Settings", "TaskBarAutoHide", chkTaskHide.Value)
Call WriteINI("Settings", "TaskBarPos", cmbTaskPos.ListIndex)
Call WriteINI("Settings", "TaskBarDelay", _
    IIf(Val(txtTaskDelay.Text) > 1, Val(txtTaskDelay.Text), 1))

For x = 1 To 5
    sPStmp(x) = ReadINI("Settings", "PS" & x)
Next
For x = 1 To 4 '4 default
    sPS(x + 1) = sPStmp(x)
Next

sPS(1) = sStr & ";" & cmbVersion.ListIndex & ";" & txtDatCallLetters.Text & ";" '& chkAutoCompile.value & ";"

If Not sPS(1) = sPS(2) Then
    For x = 1 To 5
        Call WriteINI("Settings", "PS" & x, sPS(x))
    Next
End If

If bReload Then
    Call ReloadApp
Else
    Call InitTaskbar
    If chkUseCPU.Value = 1 Then
        bUseCPU = True
        frmMain.stsStatusBar.Panels(4).Text = "Use CPU: On"
    Else
        bUseCPU = False
        frmMain.stsStatusBar.Panels(4).Text = "Use CPU: Off"
    End If
    If chkOnlyNames.Value = 1 Then
        bOnlyNames = True
    Else
        bOnlyNames = False
    End If
    Unload Me
End If

out:
Set fso = Nothing
Set fldr1 = Nothing
Exit Sub
error:
Call HandleError("cmdSave_Click")
Resume out:

End Sub


Private Sub Dir1_Change()
On Error GoTo error:

bReload = True
File1.Path = Dir1.Path

out:
Exit Sub
error:
Call HandleError("Dir1_Change")
Resume out:
End Sub

Private Sub Drive1_Change()
On Error GoTo error:

Dir1.Path = Drive1.Drive
File1.Path = Drive1.Drive

out:
Exit Sub
error:
Call HandleError("Drive1_Change")
Resume out:
End Sub

Private Sub lblAutoHide_Click()
If chkTaskHide.Value = 1 Then
    chkTaskHide.Value = 0
Else
    chkTaskHide.Value = 1
End If
End Sub

Private Sub txtDatCallLetters_Change()
bReload = True
End Sub

Private Sub txtDatCallLetters_GotFocus()
Call SelectAll(txtDatCallLetters)
End Sub

Private Sub txtTaskDelay_GotFocus()
Call SelectAll(txtTaskDelay)

End Sub


