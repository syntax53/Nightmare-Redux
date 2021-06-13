VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQuests 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quest Organizer -- Quests"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8580
   Icon            =   "frmQuests.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   8580
   Begin VB.Frame Frame1 
      Caption         =   "File"
      Height          =   915
      Left            =   7080
      TabIndex        =   11
      Top             =   1620
      Width           =   1395
      Begin VB.OptionButton optFile 
         Caption         =   "Custom"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   13
         Top             =   600
         Width           =   1035
      End
      Begin VB.OptionButton optFile 
         Caption         =   "Default"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdReload 
      Caption         =   "&Reload"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5340
      TabIndex        =   9
      Top             =   1740
      Width           =   1455
   End
   Begin VB.CommandButton cmdEditFile 
      Caption         =   "Edit Config &File"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5340
      TabIndex        =   10
      Top             =   2160
      Width           =   1455
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4440
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   6480
      TabIndex        =   5
      Top             =   480
      Width           =   435
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   6480
      TabIndex        =   8
      Top             =   1260
      Width           =   435
   End
   Begin VB.TextBox txtDescription 
      Height          =   1335
      Left            =   5220
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   2580
      Width           =   3255
   End
   Begin MSComctlLib.TreeView tvwQuests 
      Height          =   3795
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6694
      _Version        =   393217
      Indentation     =   476
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblPartNum 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7260
      TabIndex        =   4
      Top             =   420
      Width           =   1155
   End
   Begin VB.Label lblPartNumn 
      Caption         =   "Part Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7260
      TabIndex        =   2
      Top             =   180
      Width           =   1155
   End
   Begin VB.Label lblTextNum 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   5220
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblLinksTon 
      Caption         =   "Links To:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5220
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblTextNum 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   5220
      TabIndex        =   3
      Top             =   420
      Width           =   1215
   End
   Begin VB.Label lbllkkjhk 
      Caption         =   "Textblock number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5220
      TabIndex        =   1
      Top             =   180
      Width           =   1695
   End
End
Attribute VB_Name = "frmQuests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim fso As New FileSystemObject
Dim ts As TextStream
Dim sFil As String
Private Sub Form_Load()
On Error Resume Next

Me.Top = ReadINI("Windows", "QuestOrgTop")
Me.Left = ReadINI("Windows", "QuestOrgLeft")

Call CreateTree

Me.Show
Me.SetFocus
cmdReload.SetFocus
If ReadINI("Windows", "QuestOrgMaxed") = "1" Then Me.WindowState = vbMaximized
End Sub
Private Sub CreateTree()
On Error GoTo error:
Dim nStatus As Integer, Line As String
Dim nodX As Node, i As Integer, CurrentSubTree As Integer, imgX As ListImage
Dim WorkingTree As Integer, CurrentTree As Integer

If optFile(0).Value = True Then
    sFil = "NMR-Quests.txt"
Else
    sFil = "NMR-QuestsCustom.txt"
End If

If Right(App.Path, 1) = "\" Then
    sFil = App.Path & sFil
Else
    sFil = App.Path & "\" & sFil
End If

tryagain:
If fso.FileExists(sFil) = False Then
    If optFile(0).Value = True Then
        MsgBox sFil & " was not found.", vbInformation
        Exit Sub
    Else
        i = MsgBox("Custom quest file (" & sFil & ")" & vbCrLf _
            & "was not found.  Create it?", vbYesNo + vbQuestion + vbDefaultButton1)
        If i = vbNo Then Exit Sub
        Set ts = fso.CreateTextFile(sFil, True)
        
        ts.WriteLine "####"
        ts.WriteLine "#"
        ts.WriteLine "# any line beginning with a ..."
        ts.WriteLine "# '#' is ignored"
        ts.WriteLine "# '-' is a section from the root"
        ts.WriteLine "# '--' is a sub-section of the last section"
        ts.WriteLine "# '---' is a sub-section of the last sub-section"
        ts.WriteLine "# ' is a comment within the current section/sub-section"
        ts.WriteLine "#"
        ts.WriteLine "#"
        ts.WriteLine "# Textblock lines must be in the following format:"
        ts.WriteLine "#"
        ts.WriteLine "# textblock number|part number|link to number|description"
        ts.WriteLine "#"
        ts.WriteLine "# example:"
        ts.WriteLine "# -Mystic Form Quest"
        ts.WriteLine "# 2903|0|0|room 16/2667 cmd text - give <totem name> to kuel"
        ts.WriteLine "# --Form of the Crane"
        ts.WriteLine "# 2911|0|0|form of the crane spell race check"
        ts.WriteLine "# --Form of the Dragon"
        ts.WriteLine "# 2910|0|0|form of the dragon spell race check"
        ts.WriteLine "#"
        ts.WriteLine "#"
        ts.WriteLine "####"
        ts.WriteBlankLines (2)
        ts.Close
        GoTo tryagain:
    End If
Else
    Set ts = fso.OpenTextFile(sFil, ForReading)
End If

tvwQuests.Nodes.clear
ImageList1.ListImages.clear
Set imgX = ImageList1.ListImages.add(, , LoadResPicture("STAR", vbResBitmap))
Set imgX = ImageList1.ListImages.add(, , LoadResPicture("ARROW", vbResBitmap))
Set imgX = ImageList1.ListImages.add(, , LoadResPicture("PAPER", vbResBitmap))
tvwQuests.ImageList = ImageList1

i = 1

Set nodX = tvwQuests.Nodes.add(, , "Tree Quests", "Quests", 1)
nodX.Expanded = True
CurrentTree = i
WorkingTree = i
CurrentSubTree = i

Do While ts.AtEndOfStream = False

    i = i + 1

SkipLine:
    If ts.AtEndOfStream = True Then Exit Do
    Line = ts.ReadLine
    If Left(Line, 1) = "#" Then GoTo SkipLine
    If Line = "" Then GoTo SkipLine
    
    If Left(Line, 1) = "-" Then 'line could be a section/sub-section/sub-sub-section
        If Left(Line, 2) = "--" Then 'line could be a sub-section/sub-sub-section
            If Left(Line, 3) = "---" Then 'line is a sub-sub-section
                Set nodX = tvwQuests.Nodes.add(CurrentSubTree, tvwChild, "Tree " & i, Right(Line, Len(Line) - 3), 2)
                tvwQuests.Nodes(CurrentSubTree).Image = 1
                WorkingTree = i
            Else 'line is a sub-section
                Set nodX = tvwQuests.Nodes.add(CurrentTree, tvwChild, "Tree " & i, Right(Line, Len(Line) - 2), 2)
                CurrentSubTree = i
                WorkingTree = i
            End If
        Else 'line is a new section
            Set nodX = tvwQuests.Nodes.add(1, tvwChild, "Tree " & i, Right(Line, Len(Line) - 1), 1)
            CurrentTree = i
            WorkingTree = i
        End If
    Else
        If Left(Line, 1) = "'" Then 'line is a comment
            Set nodX = tvwQuests.Nodes.add(WorkingTree, tvwChild, "Cmnt " & ts.Line - 1, Left(Right(ExtractField(Line, 1), Len(ExtractField(Line, 1)) - 1), 40) & "...", 3)
        Else 'line is a textblock line
            Set nodX = tvwQuests.Nodes.add(WorkingTree, tvwChild, "Line " & ts.Line - 1, ExtractField(Line, 1) & " - p" & ExtractField(Line, 2) & " --> " & ExtractField(Line, 3), 3)
        End If
    End If
    
Loop

Set nodX = Nothing
Set imgX = Nothing

Exit Sub
error:
Call HandleError
Set nodX = Nothing
Set imgX = Nothing

End Sub
Private Sub cmdEdit_Click(Index As Integer)
    
On Error Resume Next

If FormIsLoaded("frmTextblock") Then
    If frmTextblock.fraEdit.Visible = True Then
        Call frmTextblock.GotoTB_InLine(Val(lblTextNum(Index).Caption))
    Else
        Call frmTextblock.GotoTB(Val(lblTextNum(Index).Caption), Val(lblPartNum.Caption))
    End If
Else
    Call frmTextblock.GotoTB(Val(lblTextNum(Index).Caption), Val(lblPartNum.Caption))
End If

If frmTextblock.Visible = False Then frmTextblock.Show
frmTextblock.SetFocus

End Sub

Private Sub cmdEditFile_Click()

If fso.FileExists(sFil) = True Then
    Call ShellExecute(0&, "open", sFil, vbNullString, vbNullString, vbNormalFocus)
Else
    MsgBox sFil & " was not found.", vbInformation
End If

End Sub


Private Sub cmdReload_Click()
Call CreateTree
End Sub

Private Function ExtractField(Line As String, field As Integer)
Dim x As Integer, sChar As String, Data As String, CrntFld As Integer

CrntFld = 0
x = 0

NextField:
CrntFld = CrntFld + 1
Data = ""

For x = x + 1 To Len(Line)
    sChar = Mid(Line, x, 1)
    If sChar = "|" Then Exit For
    Data = Data & sChar
Next x

If CrntFld = field Then
    ExtractField = Data
Else
    GoTo NextField
End If

End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set ts = Nothing
Set fso = Nothing
If Me.WindowState = vbMinimized Then Exit Sub

If Me.WindowState = vbMaximized Then
    Call WriteINI("Windows", "QuestOrgMaxed", 1)
Else
    Call WriteINI("Windows", "QuestOrgMaxed", 0)
    Call WriteINI("Windows", "QuestOrgTop", frmQuests.Top)
    Call WriteINI("Windows", "QuestOrgLeft", frmQuests.Left)
End If
End Sub

Private Sub optFile_Click(Index As Integer)
Call CreateTree
End Sub

Private Sub tvwQuests_NodeClick(ByVal Node As MSComctlLib.Node)
Dim LineNum As Integer, Data As String
Dim oNode As Node

For Each oNode In tvwQuests.Nodes
    oNode.BackColor = &HFFFFFF
Next

Set oNode = Nothing

Node.BackColor = &HBBBBBB

If Left(Node.Key, 4) = "Tree" Then
    lblTextNum(0).Caption = 0
    lblPartNum.Caption = 0
    lblTextNum(1).Caption = 0
    'Me.Caption = "Quest Organizer -- " & Node.Text
    txtDescription.Text = ""
    Exit Sub
End If

ts.Close
Set ts = fso.OpenTextFile(sFil, ForReading)
LineNum = Right(Node.Key, Len(Node.Key) - 5)    'the 5 is the common length of "tree " and "line "

NextLine:
If ts.Line = LineNum Then
    If Left(Node.Key, 4) = "Cmnt" Then
        Data = ts.ReadLine
        lblTextNum(0).Caption = 0
        lblPartNum.Caption = 0
        lblTextNum(1).Caption = 0
        txtDescription.Text = Right(ExtractField(Data, 1), Len(ExtractField(Data, 1)) - 1)
        Exit Sub
    Else
        Data = ts.ReadLine
        lblTextNum(0).Caption = Val(ExtractField(Data, 1))
        lblPartNum.Caption = Val(ExtractField(Data, 2))
        lblTextNum(1).Caption = Val(ExtractField(Data, 3))
        txtDescription.Text = ExtractField(Data, 4)
    End If
Else
    If ts.AtEndOfStream = True Then Exit Sub
    ts.SkipLine
    GoTo NextLine
End If

End Sub
