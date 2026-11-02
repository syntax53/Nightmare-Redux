VERSION 5.00
Begin VB.Form frmRemoteActionReport
   Caption         =   "Remote Action Report"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   11760
   Begin VB.ListBox lstRooms
      Height          =   6105
      IntegralHeight  =   0   'False
      Left            =   8460
      TabIndex        =   1
      Top             =   300
      Width           =   3300
   End
   Begin VB.TextBox txtReport
      BeginProperty Font
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6420
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   8400
   End
   Begin VB.Label lblRooms
      Caption         =   "Rooms (double-click to open):"
      Height          =   255
      Left            =   8460
      TabIndex        =   2
      Top             =   45
      Width           =   3300
   End
End
Attribute VB_Name = "frmRemoteActionReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

'The Room Editor's remote action report (modRemoteActionCheck). It follows the Room Editor that opened it: that
'editor refreshes it as it shows other rooms, and the room list jumps that editor (there can be more than one).

Private Const LIST_WIDTH As Long = 3300
Private mRoomForm As Object

Private Sub Form_Load()
On Error Resume Next
Me.Top = ReadINI("Windows", "RemoteReportTop", , "100")
Me.Left = ReadINI("Windows", "RemoteReportLeft", , "100")
Me.Width = ReadINI("Windows", "RemoteReportWidth", , "13800")
Me.Height = ReadINI("Windows", "RemoteReportHeight", , "7000")
If ReadINI("Windows", "RemoteReportMaxed") = "1" Then Me.WindowState = vbMaximized
End Sub

'shows the report for room sRoom ("map/room") of Room Editor oRoomForm: the one-line-per-exit summary on top,
'then the full report (see RemoteActionReport)
Public Sub ShowReport(ByVal oRoomForm As Object, ByVal sRoom As String, ByVal sRoomName As String, _
    ByVal sSummary As String, ByVal sText As String, ByVal nProblems As Long, ByVal nIdle As Long, _
    ByVal sRooms As String)
Dim v As Variant
On Error GoTo error:

Set mRoomForm = oRoomForm
Me.Caption = "Remote Action Report - " & sRoom & IIf(sRoomName = "", "", " (" & sRoomName & ")")
If sText = "" Then
    txtReport.Text = "Room " & sRoom & " has no remote actions and no hidden exit that needs them."
Else
    Me.Caption = Me.Caption & " - " _
        & IIf(nProblems = 0, "all OK", nProblems & " problem" & IIf(nProblems = 1, "", "s")) _
        & IIf(nIdle > 0, ", " & nIdle & " action" & IIf(nIdle = 1, "", "s") & " that do" & IIf(nIdle = 1, "es", "") & " nothing", "")
    txtReport.Text = sSummary & vbCrLf & String$(REMOTE_REPORT_WIDTH, "=") & vbCrLf & vbCrLf & sText
End If
lstRooms.Clear
For Each v In Split(sRooms, "|")
    If v <> "" Then lstRooms.AddItem v & "  " & RemoteActionRoomName(v)
Next v

Exit Sub
error:
Call HandleError("frmRemoteActionReport.ShowReport")
End Sub

'whether the report follows Room Editor oRoomForm
Public Function IsFor(ByVal oRoomForm As Object) As Boolean
IsFor = (mRoomForm Is oRoomForm)
End Function

Private Sub lstRooms_DblClick()
Dim sRoom() As String
On Error GoTo error:

If lstRooms.ListIndex < 0 Or mRoomForm Is Nothing Then Exit Sub
sRoom = Split(Split(lstRooms.Text, " ")(0), "/")
Call mRoomForm.GotoRoom(Val(sRoom(0)), Val(sRoom(1))) 'saves the room shown first; the editor refreshes this report

Exit Sub
error:
Call HandleError("frmRemoteActionReport.lstRooms_DblClick")
End Sub

Private Sub Form_Resize()
Dim nWidth As Long, nHeight As Long
On Error Resume Next

If Me.WindowState = vbMinimized Then Exit Sub
nWidth = Me.ScaleWidth
nHeight = Me.ScaleHeight
If nWidth < LIST_WIDTH + 2000 Then nWidth = LIST_WIDTH + 2000
If nHeight < 2000 Then nHeight = 2000

txtReport.Move 0, 0, nWidth - LIST_WIDTH - 60, nHeight
lblRooms.Move nWidth - LIST_WIDTH, 45, LIST_WIDTH, 255
lstRooms.Move nWidth - LIST_WIDTH, 300, LIST_WIDTH, nHeight - 300
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mRoomForm = Nothing
If Me.WindowState = vbMinimized Then Exit Sub

If Me.WindowState = vbMaximized Then
    Call WriteINI("Windows", "RemoteReportMaxed", 1)
Else
    Call WriteINI("Windows", "RemoteReportMaxed", 0)
    Call WriteINI("Windows", "RemoteReportTop", Me.Top)
    Call WriteINI("Windows", "RemoteReportLeft", Me.Left)
    Call WriteINI("Windows", "RemoteReportHeight", Me.Height)
    Call WriteINI("Windows", "RemoteReportWidth", Me.Width)
End If
End Sub
