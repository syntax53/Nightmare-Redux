VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmAbilityEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ability List Editor"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   Icon            =   "frmAbilityEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3555
   ScaleWidth      =   8205
   Begin TabDlg.SSTab SSTab1 
      Height          =   3435
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   6059
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "View"
      TabPicture(0)   =   "frmAbilityEdit.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "CommonDialog1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Edit"
      TabPicture(1)   =   "frmAbilityEdit.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdNext"
      Tab(1).Control(1)=   "cmdInfo"
      Tab(1).Control(2)=   "cmdPrev"
      Tab(1).Control(3)=   "cmdDiscard"
      Tab(1).Control(4)=   "cmdSave"
      Tab(1).Control(5)=   "cmdRemove"
      Tab(1).Control(6)=   "cmdNew"
      Tab(1).Control(7)=   "Frame5"
      Tab(1).Control(8)=   "Frame4"
      Tab(1).Control(9)=   "Frame3"
      Tab(1).ControlCount=   10
      Begin VB.CommandButton cmdNext 
         Caption         =   ">>"
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
         Left            =   -71520
         TabIndex        =   29
         Top             =   3000
         Width           =   435
      End
      Begin VB.CommandButton cmdInfo 
         Caption         =   "Info"
         Height          =   255
         Left            =   -70680
         TabIndex        =   30
         Top             =   3000
         Width           =   795
      End
      Begin VB.Frame Frame1 
         Caption         =   "Abilities"
         Height          =   2895
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Width           =   3015
         Begin VB.TextBox txtNumSearch 
            BackColor       =   &H00000000&
            ForeColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Top             =   420
            Width           =   615
         End
         Begin VB.TextBox txtNameSearch 
            BackColor       =   &H00000000&
            ForeColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   840
            TabIndex        =   5
            Top             =   420
            Width           =   2055
         End
         Begin VB.ListBox lstAbilities 
            BackColor       =   &H00000000&
            ForeColor       =   &H00E0E0E0&
            Height          =   2010
            Left            =   120
            TabIndex        =   6
            Top             =   780
            Width           =   2775
         End
         Begin VB.Label Label6 
            Caption         =   "Search Box: press --> for next"
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
            Left            =   840
            TabIndex        =   3
            Top             =   240
            Width           =   1875
         End
         Begin VB.Label lblNumberSearch 
            Caption         =   "#"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "<<"
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
         Left            =   -71940
         TabIndex        =   28
         Top             =   3000
         Width           =   435
      End
      Begin VB.CommandButton cmdDiscard 
         Caption         =   "Discard"
         Height          =   255
         Left            =   -68280
         TabIndex        =   32
         Top             =   3000
         Width           =   1155
      End
      Begin VB.Frame Frame6 
         Caption         =   "Name"
         Height          =   795
         Left            =   3240
         TabIndex        =   7
         Top             =   420
         Width           =   4695
         Begin VB.TextBox txtName 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   435
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   255
         Left            =   -69420
         TabIndex        =   31
         Top             =   3000
         Width           =   1155
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Delete"
         Height          =   255
         Left            =   -73680
         TabIndex        =   27
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "Add"
         Height          =   255
         Left            =   -74820
         TabIndex        =   26
         Top             =   3000
         Width           =   1155
      End
      Begin VB.Frame Frame5 
         Caption         =   "Information"
         Height          =   975
         Left            =   -74820
         TabIndex        =   11
         Top             =   480
         Width           =   3135
         Begin VB.TextBox txtNameEdit 
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            MaxLength       =   25
            TabIndex        =   15
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtNumberEdit 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Number:"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Color"
         Height          =   1455
         Left            =   -74820
         TabIndex        =   18
         Top             =   1440
         Width           =   3135
         Begin VB.CommandButton cmdPickColor 
            Caption         =   "PickColor"
            Height          =   375
            Left            =   2040
            TabIndex        =   25
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtColor 
            Height          =   285
            Index           =   2
            Left            =   1200
            MaxLength       =   3
            TabIndex        =   24
            Top             =   1020
            Width           =   615
         End
         Begin VB.TextBox txtColor 
            Height          =   285
            Index           =   1
            Left            =   1200
            MaxLength       =   3
            TabIndex        =   22
            Top             =   660
            Width           =   615
         End
         Begin VB.TextBox txtColor 
            Height          =   285
            Index           =   0
            Left            =   1200
            MaxLength       =   3
            TabIndex        =   20
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Blue:"
            Height          =   255
            Left            =   540
            TabIndex        =   23
            Top             =   1020
            Width           =   435
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Green:"
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   660
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Red:"
            Height          =   255
            Left            =   360
            TabIndex        =   19
            Top             =   300
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Description"
         Height          =   2415
         Left            =   -71580
         TabIndex        =   16
         Top             =   480
         Width           =   4455
         Begin VB.TextBox txtDescEdit 
            Height          =   2055
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   240
            Width           =   4215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Description"
         Height          =   2055
         Left            =   3240
         TabIndex        =   9
         Top             =   1260
         Width           =   4695
         Begin VB.TextBox txtDesc 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1695
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   240
            Width           =   4455
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3240
         Top             =   60
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmAbilityEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Private Sub Form_Load()
On Error Resume Next
Dim nYesNo As Integer

Me.Top = ReadINI("Windows", "AbEdTop")
Me.Left = ReadINI("Windows", "AbEdLeft")
    
If Not ReadINI("Settings", "ShowAbilityEditWarning") = 0 Then
    nYesNo = MsgBox("NOTE: Changes you make in the ability editor don't actually effect the game." & vbCrLf & "This is only used resolve ability numbers to names when making edits." & vbCrLf & vbCrLf & "Show this message again?", vbYesNo + vbInformation + vbDefaultButton1)
    If nYesNo = vbNo Then
        Call WriteINI("Settings", "ShowAbilityEditWarning", 0)
    Else
        Call WriteINI("Settings", "ShowAbilityEditWarning", 1)
    End If
End If
    
Call LoadAbilities
    
Me.Show
Me.SetFocus
txtNameSearch.SetFocus

End Sub
Private Sub cmdInfo_Click()
Dim nYesNo As Integer

    nYesNo = MsgBox("NOTE: Changes you make in the ability editor don't actually effect the game." & vbCrLf & "This is only used resolve ability numbers to names when making edits." & vbCrLf & vbCrLf & "Show this message again?", vbYesNo + vbInformation + vbDefaultButton1)
    If nYesNo = vbNo Then
        Call WriteINI("Settings", "ShowAbilityEditWarning", 0)
    Else
        Call WriteINI("Settings", "ShowAbilityEditWarning", 1)
    End If

End Sub

Private Sub cmdPickColor_Click()
On Error GoTo error:
CommonDialog1.CancelError = True
'CommonDialog1.Flags = &H4

On Error GoTo canceled:
CommonDialog1.ShowColor

On Error GoTo error:
txtColor(0).Text = CommonDialog1.Color And &HFF
txtColor(1).Text = (CommonDialog1.Color \ &H100) And &HFF
txtColor(2).Text = CommonDialog1.Color \ &H10000

canceled:
Exit Sub
error:
Call HandleError
End Sub


Private Sub cmdDiscard_Click()
    Call RefreshRecord
End Sub

Private Sub cmdNext_Click()

If lstAbilities.ListIndex + 1 < lstAbilities.ListCount Then
    lstAbilities.ListIndex = lstAbilities.ListIndex + 1
End If

End Sub

Private Sub cmdPrev_Click()

If lstAbilities.ListIndex - 1 >= 0 Then
    lstAbilities.ListIndex = lstAbilities.ListIndex - 1
End If

End Sub


Private Sub LoadAbilities()
Dim nCount As Integer

lstAbilities.clear

rsAbilities.MoveFirst
nCount = 0
Do Until rsAbilities.EOF
again:
    If Val(rsAbilities.Fields("Number")) > nCount Then
        lstAbilities.AddItem nCount & ". " & "unknown"
        nCount = nCount + 1
        GoTo again:
    End If
    
    lstAbilities.AddItem rsAbilities.Fields("Number") & ". " & rsAbilities.Fields("Name")
    
    rsAbilities.MoveNext
    nCount = nCount + 1
Loop

If Not lstAbilities.ListCount = 0 Then lstAbilities.ListIndex = 0

End Sub

Private Sub cmdNew_Click()
On Error GoTo error:
Dim temp As String
temp = InputBox("Enter New Ability #", "Add New Ability", lstAbilities.ListCount)
If temp = "" Or Val(temp) = 0 Then Exit Sub

    rsAbilities.AddNew
    rsAbilities.Fields("Number") = Val(temp)
    rsAbilities.Fields("Name") = "New Ability"
    rsAbilities.Fields("RED") = "255"
    rsAbilities.Fields("GREEN") = "255"
    rsAbilities.Fields("BLUE") = "255"
    rsAbilities.Fields("Description") = "New ability description here"
    rsAbilities.Update

    Call LoadAbilities
    
Exit Sub
error:
Call HandleError
End Sub

Private Sub cmdRemove_Click()
On Error GoTo error:
Dim nYesNo As Variant

nYesNo = MsgBox("Are you sure you want to delete ability #" & lstAbilities.ListIndex & "?", vbYesNo, "CONFIRM")
If nYesNo = vbNo Then Exit Sub
    
rsAbilities.Index = "PrimaryKey"
rsAbilities.Seek "=", lstAbilities.ListIndex
If rsAbilities.NoMatch Then
    MsgBox "Unable to retrieve ability #" & lstAbilities.ListIndex, vbOKOnly + vbExclamation, "Delete Failed..."
    lstAbilities.ListIndex = 0
End If

rsAbilities.Delete

Call LoadAbilities

Exit Sub
error:
Call HandleError
End Sub

Private Sub cmdSave_Click()
    Call saverecord
End Sub

Private Sub Form_Unload(Cancel As Integer)
        If Me.WindowState = vbMinimized Then Exit Sub
        Call WriteINI("Windows", "AbEdTop", frmAbilityEdit.Top)
        Call WriteINI("Windows", "AbEdLeft", frmAbilityEdit.Left)
End Sub

Private Sub RefreshRecord()
    
    txtName.Text = rsAbilities.Fields("Name")
    txtNameEdit.Text = rsAbilities.Fields("Name")
    
    txtNumberEdit.Text = rsAbilities.Fields("Number")
    
    txtDesc.Text = rsAbilities.Fields("Description")
    txtDescEdit.Text = rsAbilities.Fields("Description")
    
    txtColor(0).Text = rsAbilities.Fields("RED")
    txtColor(1).Text = rsAbilities.Fields("GREEN")
    txtColor(2).Text = rsAbilities.Fields("BLUE")
    
End Sub

Private Sub lstAbilities_Click()

rsAbilities.Index = "PrimaryKey"
rsAbilities.Seek "=", lstAbilities.ListIndex
If Not rsAbilities.NoMatch Then
    Call RefreshRecord
Else
    Call UnknownRecord
End If

End Sub

Private Sub UnknownRecord()

    txtName.Text = "unknown"
    txtNameEdit.Text = "unknown"
    
    txtNumberEdit.Text = ""
    
    txtDesc.Text = "unknown"
    txtDescEdit.Text = "unknown"
    
    txtColor(0).Text = 200
    txtColor(1).Text = 200
    txtColor(2).Text = 200

End Sub
Private Sub saverecord()
    
rsAbilities.Index = "PrimaryKey"
rsAbilities.Seek "=", lstAbilities.ListIndex
If rsAbilities.NoMatch Then
    MsgBox "Unable to retrieve ability #" & lstAbilities.ListIndex, vbOKOnly + vbExclamation, "Save Failed..."
    
    Call LoadAbilities
End If

    rsAbilities.Edit
    rsAbilities.Fields("Name") = txtNameEdit.Text
    rsAbilities.Fields("Description") = txtDescEdit.Text
    rsAbilities.Fields("RED") = Val(txtColor(0).Text)
    rsAbilities.Fields("GREEN") = Val(txtColor(1).Text)
    rsAbilities.Fields("BLUE") = Val(txtColor(2).Text)
    rsAbilities.Update

End Sub

Private Sub txtColor_Change(Index As Integer)

Call ChangeColors

End Sub

Private Sub txtColor_GotFocus(Index As Integer)
Call SelectAll(txtColor(Index))

End Sub

Private Sub txtColor_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

Call ChangeColors

End Sub
Private Sub ChangeColors()

txtName.ForeColor = RGB(Val(txtColor(0).Text), Val(txtColor(1).Text), Val(txtColor(2).Text))
txtNameEdit.ForeColor = RGB(Val(txtColor(0).Text), Val(txtColor(1).Text), Val(txtColor(2).Text))

End Sub

Private Sub txtNameEdit_GotFocus()
Call SelectAll(txtNameEdit)

End Sub

Private Sub txtNameSearch_GotFocus()
Call SelectAll(txtNameSearch)

End Sub

Private Sub txtNameSearch_KeyUp(KeyCode As Integer, Shift As Integer)
Dim nNum As Integer

nNum = 0
If txtNameSearch.Text = "" Then Exit Sub

If KeyCode = vbKeyUp Then Exit Sub
If KeyCode = vbKeyDown Then lstAbilities.SetFocus
If KeyCode = vbKeyLeft Then Exit Sub
If KeyCode = vbKeyRight Then nNum = lstAbilities.ListIndex + 1
    
For nNum = nNum To lstAbilities.ListCount - 1
    If Not InStr(1, LCase(lstAbilities.List(nNum)), LCase(txtNameSearch.Text)) = 0 Then
        lstAbilities.ListIndex = nNum
        Exit Sub
    End If
Next

End Sub


Private Sub txtNumberEdit_GotFocus()
Call SelectAll(txtNumberEdit)
End Sub

Private Sub txtNumSearch_Change()
Dim temp As Integer

If Val(txtNumSearch.Text) > 32000 Then txtNumSearch.Text = 32000
If Val(txtNumSearch.Text) < 0 Then txtNumSearch.Text = 0

temp = Val(txtNumSearch.Text)

If temp >= lstAbilities.ListCount - 1 Then Exit Sub

lstAbilities.ListIndex = temp

End Sub

Private Sub txtNumSearch_GotFocus()
Call SelectAll(txtNumSearch)

End Sub
