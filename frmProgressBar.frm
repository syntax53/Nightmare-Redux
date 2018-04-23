VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgressBar 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1785
   ClientLeft      =   5175
   ClientTop       =   5640
   ClientWidth     =   6360
   ClipControls    =   0   'False
   Icon            =   "frmProgressBar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1785
   ScaleMode       =   0  'User
   ScaleWidth      =   2300
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   180
      Width           =   1275
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   1140
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblPanel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   -60
      TabIndex        =   5
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblPanel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2340
      TabIndex        =   4
      Top             =   1560
      Width           =   4035
   End
   Begin VB.Label lblNOTE 
      Caption         =   "NOTE: Minimizing this window can significantly increase the processing speed."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   6195
   End
   Begin VB.Label lblCaption 
      Caption         =   "status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim nPerCount As Long
Dim nPerUp As Long
Dim nScale As Integer
Dim nScaleCount As Long
Public sCaption As String
Public FormOwner As Form

Private Sub Form_Load()
On Error Resume Next
Dim hMenu As Long, nCount As Long

hMenu = GetSystemMenu(Me.hwnd, 0)
nCount = GetMenuItemCount(hMenu)
Call RemoveMenu(hMenu, nCount - 1, MF_DISABLED Or MF_BYPOSITION)
DrawMenuBar Me.hwnd

nScale = 0
nScaleCount = 1
ProgressBar.Value = 0
ProgressBar.Min = 0
ProgressBar.Max = 32767
lblNote.Visible = False

End Sub
Private Sub cmdCancel_Click()
Dim nYesNo As Integer
On Error GoTo error:

If lblCaption.Caption = "Compiling update file..." Then
    Call UpdateFileCancel
    DoEvents
    Exit Sub
End If

Select Case sCaption
    Case "Map Builder":
        Call frmMap.ToggleStopBuild
    
    Case "Building Control Room List"
        bStopControlBuild = True
        
    Case "TextBlock Search":
        If FormOwner Is Nothing Then
            Call frmTextblock.StopSearch
        Else
            Call FormOwner.StopSearch
        End If
        
    Case "Message Search":
        If FormOwner Is Nothing Then
            Call frmMessage.StopSearch
        Else
            Call FormOwner.StopSearch
        End If
        
    Case "Changing room call letters", "Find Room Item", _
        "Padding Room Numbers", "Deleting buffer rooms", "Combining Monster Exp", _
        "Removing Limits on Items", "Removing Level Restrictions on Items", _
        "Dividing Monster EXP", "Multiplying Monster EXP", "Combining like items in rooms", _
        "Setting Monster Last Killed Date/Time", "Reseting Monster Last Killed Date/Time", _
        "Restocking Shops", "Stripping TB Chars", "Fixing Number of Uses on Items", _
        "Change User's Gang Names", "Retrain Users", "Fixing Number of Uses on Monster Item Drops":
        nYesNo = MsgBox("Are you sure you want to cancel?", vbYesNo + vbQuestion + vbDefaultButton2)
        If Not nYesNo = vbYes Then Exit Sub
        Call frmMain.CancelProcess
        
    Case "Building Limited Item List":
        Call frmLimitedList.ToggleStopBuild
        
    Case "Find Item":
        Call frmItemFind.ToggleStopBuild
        
    Case "Building Monster Group/Index List":
        Call frmMonsterIndex.ToggleStopBuild
    
    Case "Building Monster NPC/Room List":
        Call frmMonsterNPC_List.ToggleStopBuild
        
    Case "Search Rooms for String":
        If FormOwner Is Nothing Then
            Call frmRoom.ToggleSearchStop
        Else
            Call FormOwner.ToggleSearchStop
        End If
End Select

DoEvents

out:
Exit Sub
error:
Call HandleError("cmdCancel_Click")
Resume out:

End Sub
Public Sub IncreaseProgress()
On Error Resume Next

If Me.Caption = "" Then Me.Caption = "0% NMR: " & sCaption

If nScale > 0 Then
    If nScaleCount >= nScale Then
        'If ProgressBar.Value + 1 < ProgressBar.Max Then ProgressBar.Value = ProgressBar.Value + 1
        ProgressBar.Value = ProgressBar.Value + 1
        nScaleCount = 1
    Else
        nScaleCount = nScaleCount + 1
    End If
Else
    'If ProgressBar.Value + 1 < ProgressBar.Max Then ProgressBar.Value = ProgressBar.Value + 1
    ProgressBar.Value = ProgressBar.Value + 1
End If

If nPerCount >= nPerUp Then
    Me.Caption = CStr(Round(frmProgressBar.ProgressBar.Value / frmProgressBar.ProgressBar.Max, 2) * 100) & "% NMR: " & sCaption
    nPerCount = 1
Else
    nPerCount = nPerCount + 1
End If

End Sub
Public Sub SetRange(ByVal MaxValue As Double)
Dim nNewMax As Integer

nScale = 0

If MaxValue > MaxInt Then
    If MaxValue / 2 < MaxInt Then
        nScale = 2
        nNewMax = MaxValue / 2
    ElseIf MaxValue / 4 < MaxInt Then
        nScale = 4
        nNewMax = MaxValue / 4
    ElseIf MaxValue / 8 < MaxInt Then
        nScale = 8
        nNewMax = MaxValue / 8
    ElseIf MaxValue / 10 < MaxInt Then
        nScale = 10
        nNewMax = MaxValue / 10
    ElseIf MaxValue / 50 < MaxInt Then
        nScale = 50
        nNewMax = MaxValue / 50
    Else
        MaxValue = MaxInt
    End If
Else
    nNewMax = MaxValue
End If

nNewMax = Fix(nNewMax)

nScaleCount = 1
ProgressBar.Value = 0
ProgressBar.Min = 0
ProgressBar.Max = nNewMax
nPerCount = 1
nPerUp = Fix((nNewMax * IIf(nScale > 0, nScale, 1)) * 0.01)
End Sub
Private Sub UpdateFileCancel()
On Error Resume Next
Dim nYesNo As Integer, WGPath As String, nStatus As Integer
Dim fso As FileSystemObject, fil1 As File

nYesNo = MsgBox("Are you sure you want to cancel?", vbYesNo + vbQuestion)
If nYesNo = vbNo Then Exit Sub

Call modUpdateFile.StopUpdate
DoEvents

Set fso = CreateObject("Scripting.FileSystemObject")

WGPath = ReadINI("Settings", "WGPath" & IIf(WorksWithN = True, "_n", IIf(WorksWithWG = True, "_wg", "")))
UpdateKeyBuffer = WGPath & "w" & strDatCallLetters & strDatSuffix_UPDAT

If fso.FileExists(UpdateKeyBuffer) = True Then
    Set fil1 = fso.GetFile(UpdateKeyBuffer)
    fil1.Delete True
End If

Set fso = Nothing
Set fil1 = Nothing

DoEvents
MsgBox "Update file creation cancled, update file has been deleted.", vbExclamation

End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then frmMain.WindowState = vbMinimized
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FormOwner = Nothing
End Sub

