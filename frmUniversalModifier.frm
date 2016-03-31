VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUniversalModifier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Universal Modifier"
   ClientHeight    =   4380
   ClientLeft      =   450
   ClientTop       =   975
   ClientWidth     =   7860
   Icon            =   "frmUniversalModifier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4380
   ScaleWidth      =   7860
   Begin VB.Frame Frame4 
      Caption         =   "Only Perform Action If:"
      Height          =   1155
      Left            =   3660
      TabIndex        =   37
      Top             =   2040
      Width           =   4095
      Begin VB.CheckBox chkOnlyIfOn 
         Caption         =   "Must be"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   43
         Top             =   780
         Width           =   915
      End
      Begin VB.TextBox txtOnlyIfValue 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   2400
         TabIndex        =   42
         Text            =   "0"
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cmbOnlyIfModifier 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         ItemData        =   "frmUniversalModifier.frx":08CA
         Left            =   1020
         List            =   "frmUniversalModifier.frx":08DA
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   720
         Width           =   1275
      End
      Begin VB.ComboBox cmbOnlyIf 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         ItemData        =   "frmUniversalModifier.frx":08F8
         Left            =   120
         List            =   "frmUniversalModifier.frx":08FA
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox cmbOnlyIfValue 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         ItemData        =   "frmUniversalModifier.frx":08FC
         Left            =   2400
         List            =   "frmUniversalModifier.frx":0921
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cmbOnlyIfAuxValue 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         ItemData        =   "frmUniversalModifier.frx":097C
         Left            =   2400
         List            =   "frmUniversalModifier.frx":09A1
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblOnlyIfAuxValue 
         Caption         =   "lbl"
         Height          =   195
         Index           =   1
         Left            =   2400
         TabIndex        =   44
         Top             =   180
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdNotes 
      Caption         =   "*&READ!*"
      Height          =   375
      Left            =   1860
      TabIndex        =   33
      Top             =   2820
      Width           =   1575
   End
   Begin VB.CheckBox chkOnlyChanges 
      Caption         =   "Only Log Changes"
      Height          =   195
      Left            =   1020
      TabIndex        =   36
      Top             =   2280
      Width           =   1635
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "&Log"
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   2820
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Only Perform Action If:"
      Height          =   1155
      Left            =   3660
      TabIndex        =   19
      Top             =   780
      Width           =   4095
      Begin VB.ComboBox cmbOnlyIfAuxValue 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         ItemData        =   "frmUniversalModifier.frx":09FC
         Left            =   2400
         List            =   "frmUniversalModifier.frx":0A21
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cmbOnlyIfValue 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         ItemData        =   "frmUniversalModifier.frx":0A7C
         Left            =   2400
         List            =   "frmUniversalModifier.frx":0AA1
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cmbOnlyIf 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         ItemData        =   "frmUniversalModifier.frx":0AFC
         Left            =   120
         List            =   "frmUniversalModifier.frx":0AFE
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox cmbOnlyIfModifier 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         ItemData        =   "frmUniversalModifier.frx":0B00
         Left            =   1020
         List            =   "frmUniversalModifier.frx":0B10
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   720
         Width           =   1275
      End
      Begin VB.TextBox txtOnlyIfValue 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   2400
         TabIndex        =   25
         Text            =   "0"
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox chkOnlyIfOn 
         Caption         =   "Must be"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   780
         Width           =   915
      End
      Begin VB.Label lblOnlyIfAuxValue 
         Caption         =   "lbl"
         Height          =   195
         Index           =   0
         Left            =   2400
         TabIndex        =   21
         Top             =   180
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Range:"
      Height          =   1155
      Left            =   300
      TabIndex        =   11
      Top             =   780
      Width           =   2895
      Begin VB.TextBox txtMap 
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   18
         Top             =   720
         Width           =   795
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "All"
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
         Left            =   180
         TabIndex        =   14
         Top             =   360
         Value           =   1  'Checked
         Width           =   675
      End
      Begin VB.TextBox txtR1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   15
         Top             =   360
         Width           =   795
      End
      Begin VB.TextBox txtR2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1860
         TabIndex        =   16
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblMap 
         Caption         =   "Map:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblTo 
         Caption         =   "To"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1860
         TabIndex        =   13
         Top             =   180
         Width           =   795
      End
      Begin VB.Label lblFrom 
         Caption         =   "From"
         Enabled         =   0   'False
         Height          =   195
         Left            =   900
         TabIndex        =   12
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Limit of Result:"
      Height          =   675
      Left            =   3660
      TabIndex        =   27
      Top             =   3360
      Width           =   4095
      Begin VB.CheckBox chkLimit 
         Caption         =   "Limit to"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   300
         Width           =   855
      End
      Begin VB.TextBox txtLimit 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   30
         Text            =   "0"
         Top             =   240
         Width           =   1395
      End
      Begin VB.ComboBox cmbLimitModifier 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmUniversalModifier.frx":0B2E
         Left            =   1080
         List            =   "frmUniversalModifier.frx":0B38
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   495
      Left            =   1860
      TabIndex        =   32
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   495
      Left            =   120
      TabIndex        =   31
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox txtValue 
      Height          =   315
      Left            =   6540
      TabIndex        =   10
      Top             =   300
      Width           =   1215
   End
   Begin VB.ComboBox cmbModifier 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmUniversalModifier.frx":0B4C
      Left            =   5580
      List            =   "frmUniversalModifier.frx":0B62
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   300
      Width           =   855
   End
   Begin VB.ComboBox cmbField 
      Height          =   315
      ItemData        =   "frmUniversalModifier.frx":0B78
      Left            =   1740
      List            =   "frmUniversalModifier.frx":0B7A
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   300
      Width           =   1815
   End
   Begin VB.ComboBox cmbEditor 
      Height          =   315
      ItemData        =   "frmUniversalModifier.frx":0B7C
      Left            =   60
      List            =   "frmUniversalModifier.frx":0B95
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   300
      Width           =   1575
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   34
      Top             =   4125
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3414
            MinWidth        =   3414
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10362
            MinWidth        =   2927
         EndProperty
      EndProperty
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
   Begin VB.ComboBox cmbAbilities 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmUniversalModifier.frx":0BD0
      Left            =   3660
      List            =   "frmUniversalModifier.frx":0BD2
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   300
      Width           =   1815
   End
   Begin VB.ComboBox cmbValue 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6540
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   300
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblModifier 
      Caption         =   "Modifier"
      Height          =   255
      Left            =   5580
      TabIndex        =   3
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Field"
      Height          =   255
      Left            =   1740
      TabIndex        =   1
      Top             =   60
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Editor"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1575
   End
   Begin VB.Label lblValue 
      Caption         =   "Value"
      Height          =   195
      Left            =   6600
      TabIndex        =   4
      Top             =   60
      Width           =   675
   End
   Begin VB.Label lblAbility 
      Caption         =   "Ability"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3660
      TabIndex        =   2
      Top             =   60
      Width           =   1635
   End
End
Attribute VB_Name = "frmUniversalModifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Private Enum eDataType
    dtInteger = 1
    dtLong = 2
    dtByte = 3
End Enum

Private Enum eSU
    suSigned = 1
    suUnsigned = 2
End Enum

Dim sLogFile As String
Dim ts As TextStream
Dim fso As FileSystemObject
Dim bStopProcess As Boolean



Private Sub cmbField_Click()
Dim bEnable As Boolean, bNoValue As Boolean, bNoMod As Boolean

cmbAbilities.Enabled = False
lblAbility.Enabled = False
cmbModifier.Enabled = True
lblModifier.Enabled = True
txtValue.Enabled = True
lblValue.Enabled = True

Select Case cmbEditor.ListIndex
    Case 0: '0 - class
        Select Case cmbField.ListIndex
            Case 4:
                bEnable = True
                bNoMod = True
            Case 5:
                bEnable = True
                bNoValue = True
                bNoMod = True
            Case 6:
                bEnable = True
        End Select
    Case 1: '1 - item
        Select Case cmbField.ListIndex
            Case 11:
                bEnable = True
                bNoMod = True
            Case 12:
                bEnable = True
                bNoValue = True
                bNoMod = True
            Case 13:
                bEnable = True
        End Select
        
    Case 2: '2 - mons
        Select Case cmbField.ListIndex
            Case 22:
                bEnable = True
                bNoMod = True
            Case 23:
                bEnable = True
                bNoValue = True
                bNoMod = True
            Case 24:
                bEnable = True
        End Select
        
    Case 3: '3 - race
        Select Case cmbField.ListIndex
            Case 17:
                bEnable = True
                bNoMod = True
            Case 18:
                bEnable = True
                bNoValue = True
                bNoMod = True
            Case 19:
                bEnable = True
        End Select
        
    Case 4: '4 - room
        
    Case 5: '5 - shop
        
    Case 6: '6 - spell
        Select Case cmbField.ListIndex
            Case 14:
                bEnable = True
                bNoMod = True
            Case 15:
                bEnable = True
                bNoValue = True
                bNoMod = True
            Case 16:
                bEnable = True
        End Select
        
End Select

If bEnable = True Then
    cmbModifier.ListIndex = 0
    cmbAbilities.Enabled = True
    lblAbility.Enabled = True
    
    If bNoValue = True Then
        lblValue.Enabled = False
        txtValue.Enabled = False
    End If
    
    If bNoMod = True Then
        lblModifier.Enabled = False
        cmbModifier.Enabled = False
    End If
End If

End Sub


Private Sub cmdLog_Click()
On Error GoTo error:

If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")

If Right(App.Path, 1) = "\" Then
    sLogFile = App.Path & "NMR-Log_Universal.txt"
Else
    sLogFile = App.Path & "\NMR-Log_Universal.txt"
End If

If fso.FileExists(sLogFile) = True Then
    Call ShellExecute(0&, "open", sLogFile, vbNullString, vbNullString, vbNormalFocus)
Else
    MsgBox sLogFile & " was not found.", vbInformation
End If

out:
Exit Sub
error:
Call HandleError("cmdLog_Click")
Resume out:
End Sub

Private Sub Form_Load()
On Error Resume Next
cmbEditor.ListIndex = 0
cmbModifier.ListIndex = 0
cmbLimitModifier.ListIndex = 1
cmbValue.ListIndex = 0
cmbOnlyIfModifier(0).ListIndex = 0
cmbOnlyIfModifier(1).ListIndex = 0
cmbAbilities.clear

Dim i As Integer
For i = 1 To 100
    cmbValue.AddItem i
Next i
For i = 1 To 9
    cmbValue.AddItem "1" & i * 10
Next i
For i = 2 To 10
    cmbValue.AddItem i * 100
Next i

Call AddAbilities

Me.Top = ReadINI("Windows", "UniTop")
Me.Left = ReadINI("Windows", "UniLeft")
chkOnlyChanges.Value = ReadINI("Settings", "UniLogOnlyChanges")

Me.Show
Me.SetFocus
'cmdCancel.SetFocus

End Sub

Private Sub AddAbilities()
On Error GoTo error:
Dim x As Integer

If bAbilityDBOpen = False Then
    For x = 0 To 190
        cmbAbilities.AddItem "Ability #" & x
        cmbAbilities.ItemData(cmbAbilities.NewIndex) = x
    Next
    cmbAbilities.ListIndex = 0
    Exit Sub
End If

rsAbilities.MoveFirst

x = 0
Do While Not rsAbilities.EOF
again:
    If Val(rsAbilities.Fields("Number")) > x Then
        cmbAbilities.AddItem x & "-unknown"
        cmbAbilities.ItemData(cmbAbilities.ListCount - 1) = x
        x = x + 1
        GoTo again:
    End If
    
    cmbAbilities.AddItem rsAbilities.Fields("Number") & "-" & rsAbilities.Fields("Name")
    cmbAbilities.ItemData(cmbAbilities.ListCount - 1) = Val(rsAbilities.Fields("Number"))
    rsAbilities.MoveNext
    x = x + 1
Loop
cmbAbilities.ListIndex = 0

Exit Sub
error:
Call HandleError
End Sub
Private Sub chkOnlyIfOn_Click(Index As Integer)
If chkOnlyIfOn(Index).Value = 1 Then
    cmbOnlyIf(Index).Enabled = True
    txtOnlyIfValue(Index).Enabled = True
    cmbOnlyIfModifier(Index).Enabled = True
    cmbOnlyIfAuxValue(Index).Enabled = True
    cmbOnlyIfValue(Index).Enabled = True
    Call cmbOnlyIf_Click(Index)
Else
    cmbOnlyIf(Index).Enabled = False
    
    txtOnlyIfValue(Index).Enabled = False
'    txtOnlyIfValue(Index).Visible = True
'    cmbOnlyIfValue(Index0).Visible = False
    cmbOnlyIfValue(Index).Enabled = False
    
    cmbOnlyIfAuxValue(Index).Visible = False
    cmbOnlyIfAuxValue(Index).Enabled = False
    
    cmbOnlyIfModifier(Index).Enabled = False
End If
End Sub

Private Sub cmbOnlyIf_Click(Index As Integer)
Dim x As Integer

txtOnlyIfValue(Index).Visible = True
cmbOnlyIfValue(Index).Visible = False
cmbOnlyIfModifier(Index).Locked = False
cmbOnlyIfAuxValue(Index).Visible = False
cmbOnlyIfAuxValue(Index).Locked = False
lblOnlyIfAuxValue(Index).Visible = False

Select Case cmbEditor.ListIndex

    Case 0: '0 - class
        Select Case cmbOnlyIf(Index).ListIndex
            Case 0: 'combat
            
            Case 1: 'magery
                cmbOnlyIfAuxValue(Index).clear
                cmbOnlyIfAuxValue(Index).AddItem "None"
                cmbOnlyIfAuxValue(Index).AddItem "Mage"
                cmbOnlyIfAuxValue(Index).AddItem "Priest"
                cmbOnlyIfAuxValue(Index).AddItem "Druid"
                cmbOnlyIfAuxValue(Index).AddItem "Bard"
                cmbOnlyIfAuxValue(Index).AddItem "Kai"
                
                cmbOnlyIfAuxValue(Index).ListIndex = 0
                cmbOnlyIfAuxValue(Index).Visible = True
            Case 2: 'exp%
            Case 3: 'hp Min
            Case 4: 'hp max
            Case 5, 6: 'abilities
                cmbOnlyIfAuxValue(Index).clear
                For x = 0 To cmbAbilities.ListCount - 1
                    cmbOnlyIfAuxValue(Index).AddItem cmbAbilities.List(x), x
                    cmbOnlyIfAuxValue(Index).ItemData(x) = cmbAbilities.ItemData(x)
                Next
                
                cmbOnlyIfAuxValue(Index).ListIndex = 0
                cmbOnlyIfAuxValue(Index).Visible = True
        End Select
        
    Case 1: '1 - item
        Select Case cmbOnlyIf(Index).ListIndex
            Case 1: 'item type
                cmbOnlyIfValue(Index).clear
                cmbOnlyIfValue(Index).AddItem "Armour"
                cmbOnlyIfValue(Index).AddItem "Weapon"
                cmbOnlyIfValue(Index).AddItem "Projectile"
                cmbOnlyIfValue(Index).AddItem "Sign"
                cmbOnlyIfValue(Index).AddItem "Food"
                cmbOnlyIfValue(Index).AddItem "Drink"
                cmbOnlyIfValue(Index).AddItem "Light"
                cmbOnlyIfValue(Index).AddItem "Key"
                cmbOnlyIfValue(Index).AddItem "Container"
                cmbOnlyIfValue(Index).AddItem "Scroll"
                cmbOnlyIfValue(Index).AddItem "Special"
                
                Call OFUseComboValue(Index)
            
            Case 2: 'armour type
                cmbOnlyIfAuxValue(Index).clear
                cmbOnlyIfAuxValue(Index).AddItem "Armour"
                cmbOnlyIfAuxValue(Index).AddItem "Weapon"
                cmbOnlyIfAuxValue(Index).AddItem "Projectile"
                cmbOnlyIfAuxValue(Index).AddItem "Sign"
                cmbOnlyIfAuxValue(Index).AddItem "Food"
                cmbOnlyIfAuxValue(Index).AddItem "Drink"
                cmbOnlyIfAuxValue(Index).AddItem "Light"
                cmbOnlyIfAuxValue(Index).AddItem "Key"
                cmbOnlyIfAuxValue(Index).AddItem "Container"
                cmbOnlyIfAuxValue(Index).AddItem "Scroll"
                cmbOnlyIfAuxValue(Index).AddItem "Special"
                
                cmbOnlyIfAuxValue(Index).ListIndex = 0
                cmbOnlyIfAuxValue(Index).Locked = True
                cmbOnlyIfAuxValue(Index).Visible = True
                
                lblOnlyIfAuxValue(Index).Caption = "Item Type"
                lblOnlyIfAuxValue(Index).Visible = True
                
                cmbOnlyIfValue(Index).clear
                cmbOnlyIfValue(Index).AddItem "Natural"
                cmbOnlyIfValue(Index).AddItem "Robes"
                cmbOnlyIfValue(Index).AddItem "Padded/Ninja"
                cmbOnlyIfValue(Index).AddItem "Soft Leather"
                cmbOnlyIfValue(Index).AddItem "Soft Stud Leather"
                cmbOnlyIfValue(Index).AddItem "Rigid Leather"
                cmbOnlyIfValue(Index).AddItem "Stud Rigid Leather"
                cmbOnlyIfValue(Index).AddItem "Chainmail"
                cmbOnlyIfValue(Index).AddItem "Scalemail"
                cmbOnlyIfValue(Index).AddItem "Platemail"
                
                Call OFUseComboValue(Index)
                
            Case 3: 'weapon type
                cmbOnlyIfAuxValue(Index).clear
                cmbOnlyIfAuxValue(Index).AddItem "Armour"
                cmbOnlyIfAuxValue(Index).AddItem "Weapon"
                cmbOnlyIfAuxValue(Index).AddItem "Projectile"
                cmbOnlyIfAuxValue(Index).AddItem "Sign"
                cmbOnlyIfAuxValue(Index).AddItem "Food"
                cmbOnlyIfAuxValue(Index).AddItem "Drink"
                cmbOnlyIfAuxValue(Index).AddItem "Light"
                cmbOnlyIfAuxValue(Index).AddItem "Key"
                cmbOnlyIfAuxValue(Index).AddItem "Container"
                cmbOnlyIfAuxValue(Index).AddItem "Scroll"
                cmbOnlyIfAuxValue(Index).AddItem "Special"
                
                cmbOnlyIfAuxValue(Index).ListIndex = 1
                cmbOnlyIfAuxValue(Index).Locked = True
                cmbOnlyIfAuxValue(Index).Visible = True
                
                lblOnlyIfAuxValue(Index).Caption = "Item Type"
                lblOnlyIfAuxValue(Index).Visible = True
                
                cmbOnlyIfValue(Index).clear
                cmbOnlyIfValue(Index).AddItem "1 H Blunt"
                cmbOnlyIfValue(Index).AddItem "2 H Blunt"
                cmbOnlyIfValue(Index).AddItem "1 H Sharp"
                cmbOnlyIfValue(Index).AddItem "2 H Sharp"
                
                Call OFUseComboValue(Index)
            
            Case 4: 'worn on
                cmbOnlyIfValue(Index).clear
                cmbOnlyIfValue(Index).AddItem "Nowhere"
                cmbOnlyIfValue(Index).AddItem "Everywhere"
                cmbOnlyIfValue(Index).AddItem "Head"
                cmbOnlyIfValue(Index).AddItem "Hands"
                cmbOnlyIfValue(Index).AddItem "Finger (1)"
                cmbOnlyIfValue(Index).AddItem "Feet"
                cmbOnlyIfValue(Index).AddItem "Arms"
                cmbOnlyIfValue(Index).AddItem "Back"
                cmbOnlyIfValue(Index).AddItem "Neck"
                cmbOnlyIfValue(Index).AddItem "Legs"
                cmbOnlyIfValue(Index).AddItem "Waist"
                cmbOnlyIfValue(Index).AddItem "Torso"
                cmbOnlyIfValue(Index).AddItem "Off-Hand"
                cmbOnlyIfValue(Index).AddItem "Finger (2)"
                cmbOnlyIfValue(Index).AddItem "Wrist (1)"
                cmbOnlyIfValue(Index).AddItem "Ears"
                cmbOnlyIfValue(Index).AddItem "Worn"
                cmbOnlyIfValue(Index).AddItem "Wrist (2)"
                cmbOnlyIfValue(Index).AddItem "Eyes"
                cmbOnlyIfValue(Index).AddItem "Face"
                
'                Wrist (1)
'                Ears
'                Worn
'                Wrist (2)
'                Eyes
'                Face
'
                Call OFUseComboValue(Index)
            
            Case 9: 'cost
                
                cmbOnlyIfAuxValue(Index).clear
                cmbOnlyIfAuxValue(Index).AddItem "Copper"
                cmbOnlyIfAuxValue(Index).AddItem "Silver"
                cmbOnlyIfAuxValue(Index).AddItem "Gold"
                cmbOnlyIfAuxValue(Index).AddItem "Platinum"
                cmbOnlyIfAuxValue(Index).AddItem "Runic"
                cmbOnlyIfAuxValue(Index).AddItem "Any"
                
                cmbOnlyIfAuxValue(Index).ListIndex = 0
                cmbOnlyIfAuxValue(Index).Visible = True
            Case 14, 15: 'abilities
                cmbOnlyIfAuxValue(Index).clear
                For x = 0 To cmbAbilities.ListCount - 1
                    cmbOnlyIfAuxValue(Index).AddItem cmbAbilities.List(x), x
                    cmbOnlyIfAuxValue(Index).ItemData(x) = cmbAbilities.ItemData(x)
                Next
                
                cmbOnlyIfAuxValue(Index).ListIndex = 0
                cmbOnlyIfAuxValue(Index).Visible = True
        End Select
        
    Case 2: '2 - mons
        Select Case cmbOnlyIf(Index).ListIndex
            Case 3: 'mon index w/group
                cmbOnlyIfAuxValue(Index).clear
                cmbOnlyIfAuxValue(Index).AddItem "Lair"
                cmbOnlyIfAuxValue(Index).AddItem "Wanderer"
                cmbOnlyIfAuxValue(Index).AddItem "NPC"
                cmbOnlyIfAuxValue(Index).AddItem "Living"
                cmbOnlyIfAuxValue(Index).AddItem "Random"
                cmbOnlyIfAuxValue(Index).AddItem "Guard"
                cmbOnlyIfAuxValue(Index).AddItem "Group 1"
                cmbOnlyIfAuxValue(Index).AddItem "Group 2"
                cmbOnlyIfAuxValue(Index).AddItem "Group 3"
                cmbOnlyIfAuxValue(Index).AddItem "Group 4"
                cmbOnlyIfAuxValue(Index).AddItem "Group 5"
                cmbOnlyIfAuxValue(Index).AddItem "Group 6"
                cmbOnlyIfAuxValue(Index).AddItem "Group 7"
                cmbOnlyIfAuxValue(Index).AddItem "Group 8"
                cmbOnlyIfAuxValue(Index).AddItem "Group 9"
                cmbOnlyIfAuxValue(Index).AddItem "Group 10"
                cmbOnlyIfAuxValue(Index).AddItem "Group 11"
                cmbOnlyIfAuxValue(Index).AddItem "Group 12"
                cmbOnlyIfAuxValue(Index).AddItem "Group 13"
                cmbOnlyIfAuxValue(Index).AddItem "Group 14"
                cmbOnlyIfAuxValue(Index).AddItem "Group 15"
                cmbOnlyIfAuxValue(Index).AddItem "Group 16"
                cmbOnlyIfAuxValue(Index).AddItem "Group 17"
                cmbOnlyIfAuxValue(Index).AddItem "Group 18"
                cmbOnlyIfAuxValue(Index).AddItem "Group 19"
                cmbOnlyIfAuxValue(Index).AddItem "Group 20"
                cmbOnlyIfAuxValue(Index).AddItem "Group 21"
                cmbOnlyIfAuxValue(Index).AddItem "Group 22"
                cmbOnlyIfAuxValue(Index).AddItem "Group 23"
                cmbOnlyIfAuxValue(Index).AddItem "Group 24"
                cmbOnlyIfAuxValue(Index).AddItem "Group 25"
                cmbOnlyIfAuxValue(Index).AddItem "Group 26"
                cmbOnlyIfAuxValue(Index).AddItem "Group 27"
                cmbOnlyIfAuxValue(Index).AddItem "Group 28"
                cmbOnlyIfAuxValue(Index).AddItem "Group 29"
                cmbOnlyIfAuxValue(Index).AddItem "Group 30"
                cmbOnlyIfAuxValue(Index).AddItem "Arena"
                cmbOnlyIfAuxValue(Index).AddItem "Angel"
                cmbOnlyIfAuxValue(Index).AddItem "Quest"
                cmbOnlyIfAuxValue(Index).AddItem "Other"
                
                cmbOnlyIfAuxValue(Index).ListIndex = 0
                cmbOnlyIfAuxValue(Index).Visible = True
                
                lblOnlyIfAuxValue(Index).Caption = "Group"
                lblOnlyIfAuxValue(Index).Visible = True
            
            Case 4: 'group
                cmbOnlyIfValue(Index).clear
                cmbOnlyIfValue(Index).AddItem "Lair"
                cmbOnlyIfValue(Index).AddItem "Wanderer"
                cmbOnlyIfValue(Index).AddItem "NPC"
                cmbOnlyIfValue(Index).AddItem "Living"
                cmbOnlyIfValue(Index).AddItem "Random"
                cmbOnlyIfValue(Index).AddItem "Guard"
                cmbOnlyIfValue(Index).AddItem "Group 1"
                cmbOnlyIfValue(Index).AddItem "Group 2"
                cmbOnlyIfValue(Index).AddItem "Group 3"
                cmbOnlyIfValue(Index).AddItem "Group 4"
                cmbOnlyIfValue(Index).AddItem "Group 5"
                cmbOnlyIfValue(Index).AddItem "Group 6"
                cmbOnlyIfValue(Index).AddItem "Group 7"
                cmbOnlyIfValue(Index).AddItem "Group 8"
                cmbOnlyIfValue(Index).AddItem "Group 9"
                cmbOnlyIfValue(Index).AddItem "Group 10"
                cmbOnlyIfValue(Index).AddItem "Group 11"
                cmbOnlyIfValue(Index).AddItem "Group 12"
                cmbOnlyIfValue(Index).AddItem "Group 13"
                cmbOnlyIfValue(Index).AddItem "Group 14"
                cmbOnlyIfValue(Index).AddItem "Group 15"
                cmbOnlyIfValue(Index).AddItem "Group 16"
                cmbOnlyIfValue(Index).AddItem "Group 17"
                cmbOnlyIfValue(Index).AddItem "Group 18"
                cmbOnlyIfValue(Index).AddItem "Group 19"
                cmbOnlyIfValue(Index).AddItem "Group 20"
                cmbOnlyIfValue(Index).AddItem "Group 21"
                cmbOnlyIfValue(Index).AddItem "Group 22"
                cmbOnlyIfValue(Index).AddItem "Group 23"
                cmbOnlyIfValue(Index).AddItem "Group 24"
                cmbOnlyIfValue(Index).AddItem "Group 25"
                cmbOnlyIfValue(Index).AddItem "Group 26"
                cmbOnlyIfValue(Index).AddItem "Group 27"
                cmbOnlyIfValue(Index).AddItem "Group 28"
                cmbOnlyIfValue(Index).AddItem "Group 29"
                cmbOnlyIfValue(Index).AddItem "Group 30"
                cmbOnlyIfValue(Index).AddItem "Arena"
                cmbOnlyIfValue(Index).AddItem "Angel"
                cmbOnlyIfValue(Index).AddItem "Quest"
                cmbOnlyIfValue(Index).AddItem "Other"
                
                Call OFUseComboValue(Index)
            Case 19, 20: 'abilities
                cmbOnlyIfAuxValue(Index).clear
                For x = 0 To cmbAbilities.ListCount - 1
                    cmbOnlyIfAuxValue(Index).AddItem cmbAbilities.List(x), x
                    cmbOnlyIfAuxValue(Index).ItemData(x) = cmbAbilities.ItemData(x)
                Next
                
                cmbOnlyIfAuxValue(Index).ListIndex = 0
                cmbOnlyIfAuxValue(Index).Visible = True
        End Select

    Case 3: '3 - race
        Select Case cmbOnlyIf(Index).ListIndex
            Case 0, 1: 'abilities
                cmbOnlyIfAuxValue(Index).clear
                For x = 0 To cmbAbilities.ListCount - 1
                    cmbOnlyIfAuxValue(Index).AddItem cmbAbilities.List(x), x
                    cmbOnlyIfAuxValue(Index).ItemData(x) = cmbAbilities.ItemData(x)
                Next
                
                cmbOnlyIfAuxValue(Index).ListIndex = 0
                cmbOnlyIfAuxValue(Index).Visible = True
        End Select
    Case 4: '4 - room
        Select Case cmbOnlyIf(Index).ListIndex
            Case 0: 'room type
                cmbOnlyIfValue(Index).clear
                cmbOnlyIfValue(Index).AddItem "Normal"
                cmbOnlyIfValue(Index).AddItem "Shop"
                cmbOnlyIfValue(Index).AddItem "Arena"
                cmbOnlyIfValue(Index).AddItem "Lair"
                cmbOnlyIfValue(Index).AddItem "Hotel"
                cmbOnlyIfValue(Index).AddItem "Colliseum"
                cmbOnlyIfValue(Index).AddItem "Jail"
                cmbOnlyIfValue(Index).AddItem "Library"
                
                Call OFUseComboValue(Index)
            Case 9: 'monster type
                cmbOnlyIfValue(Index).clear
                cmbOnlyIfValue(Index).AddItem "Lair"
                cmbOnlyIfValue(Index).AddItem "Wanderer"
                cmbOnlyIfValue(Index).AddItem "NPC"
                cmbOnlyIfValue(Index).AddItem "Living"
                cmbOnlyIfValue(Index).AddItem "Random"
                cmbOnlyIfValue(Index).AddItem "Guard"
                cmbOnlyIfValue(Index).AddItem "Group 1"
                cmbOnlyIfValue(Index).AddItem "Group 2"
                cmbOnlyIfValue(Index).AddItem "Group 3"
                cmbOnlyIfValue(Index).AddItem "Group 4"
                cmbOnlyIfValue(Index).AddItem "Group 5"
                cmbOnlyIfValue(Index).AddItem "Group 6"
                cmbOnlyIfValue(Index).AddItem "Group 7"
                cmbOnlyIfValue(Index).AddItem "Group 8"
                cmbOnlyIfValue(Index).AddItem "Group 9"
                cmbOnlyIfValue(Index).AddItem "Group 10"
                cmbOnlyIfValue(Index).AddItem "Group 11"
                cmbOnlyIfValue(Index).AddItem "Group 12"
                cmbOnlyIfValue(Index).AddItem "Group 13"
                cmbOnlyIfValue(Index).AddItem "Group 14"
                cmbOnlyIfValue(Index).AddItem "Group 15"
                cmbOnlyIfValue(Index).AddItem "Group 16"
                cmbOnlyIfValue(Index).AddItem "Group 17"
                cmbOnlyIfValue(Index).AddItem "Group 18"
                cmbOnlyIfValue(Index).AddItem "Group 19"
                cmbOnlyIfValue(Index).AddItem "Group 20"
                cmbOnlyIfValue(Index).AddItem "Group 21"
                cmbOnlyIfValue(Index).AddItem "Group 22"
                cmbOnlyIfValue(Index).AddItem "Group 23"
                cmbOnlyIfValue(Index).AddItem "Group 24"
                cmbOnlyIfValue(Index).AddItem "Group 25"
                cmbOnlyIfValue(Index).AddItem "Group 26"
                cmbOnlyIfValue(Index).AddItem "Group 27"
                cmbOnlyIfValue(Index).AddItem "Group 28"
                cmbOnlyIfValue(Index).AddItem "Group 29"
                cmbOnlyIfValue(Index).AddItem "Group 30"
                cmbOnlyIfValue(Index).AddItem "Arena"
                cmbOnlyIfValue(Index).AddItem "Angel"
                cmbOnlyIfValue(Index).AddItem "Quest"
                cmbOnlyIfValue(Index).AddItem "Other"
                
                Call OFUseComboValue(Index)
        End Select
        
    Case 5: '5 - shop
        Select Case cmbOnlyIf(Index).ListIndex
            Case 0: 'shop type
                cmbOnlyIfValue(Index).clear
                cmbOnlyIfValue(Index).AddItem "General"
                cmbOnlyIfValue(Index).AddItem "Weapons"
                cmbOnlyIfValue(Index).AddItem "Armour"
                cmbOnlyIfValue(Index).AddItem "Items"
                cmbOnlyIfValue(Index).AddItem "Spells"
                cmbOnlyIfValue(Index).AddItem "Hospital"
                cmbOnlyIfValue(Index).AddItem "Tavern"
                cmbOnlyIfValue(Index).AddItem "Bank"
                cmbOnlyIfValue(Index).AddItem "Training"
                cmbOnlyIfValue(Index).AddItem "Inn"
                cmbOnlyIfValue(Index).AddItem "Specific"
                cmbOnlyIfValue(Index).AddItem "Gang Shop"
                cmbOnlyIfValue(Index).AddItem "Deed Shop"
                
                Call OFUseComboValue(Index)
                
            Case 1: 'regen time
            Case 2: 'regen %
            
        End Select

    Case 6: '6 - spell
        Select Case cmbOnlyIf(Index).ListIndex
            Case 0: 'magery
                cmbOnlyIfAuxValue(Index).clear
                cmbOnlyIfAuxValue(Index).AddItem "None"
                cmbOnlyIfAuxValue(Index).AddItem "Mage"
                cmbOnlyIfAuxValue(Index).AddItem "Priest"
                cmbOnlyIfAuxValue(Index).AddItem "Druid"
                cmbOnlyIfAuxValue(Index).AddItem "Bard"
                cmbOnlyIfAuxValue(Index).AddItem "Kai"
                
                cmbOnlyIfAuxValue(Index).ListIndex = 0
                cmbOnlyIfAuxValue(Index).Visible = True
                
            Case 1: 'req level
            Case 15, 16: 'abilities
                cmbOnlyIfAuxValue(Index).clear
                For x = 0 To cmbAbilities.ListCount - 1
                    cmbOnlyIfAuxValue(Index).AddItem cmbAbilities.List(x), x
                    cmbOnlyIfAuxValue(Index).ItemData(x) = cmbAbilities.ItemData(x)
                Next
                
                cmbOnlyIfAuxValue(Index).ListIndex = 0
                cmbOnlyIfAuxValue(Index).Visible = True
        End Select
End Select


End Sub
Private Sub OFUseComboValue(Index As Integer)
    cmbOnlyIfValue(Index).ListIndex = 0
    cmbOnlyIfValue(Index).Visible = True
    txtOnlyIfValue(Index).Visible = False
    cmbOnlyIfModifier(Index).ListIndex = 0
'    cmbOnlyIfModifier(Index).Locked = True
End Sub


Private Sub cmdNotes_Click()
MsgBox "Notes on the Universal Modifier:" & vbCrLf _
    & "------------------------------------------" & vbCrLf _
    & "Use the 'Limit' and 'Only if' options to your advantage.  You'll need to think about the outcome of *ALL*" _
    & "the values that will be calculated.  For example, say you wanted to subtract 5 from every monster's" _
    & "regen time. Well make sure you set the limit to ' > or = 1 ' so that you don't end up with any monsters" _
    & "that instant-regen.  You'll also want to say, 'Only if Regen Time > 0' so that you don't make a monster" _
    & "that had 0 regen time now have 1 regen time.", vbInformation
End Sub


Private Sub chkAll_Click()
If chkAll.Value = 0 Then
    txtR1.Enabled = True
    txtR2.Enabled = True
    lblFrom.Enabled = True
    lblTo.Enabled = True
    If cmbEditor.ListIndex = 4 Then
        lblMap.Enabled = True
        txtMap.Enabled = True
    End If
Else
    txtR1.Enabled = False
    txtR2.Enabled = False
    lblFrom.Enabled = False
    lblTo.Enabled = False
    lblMap.Enabled = False
    txtMap.Enabled = False
End If
End Sub

Private Sub chkLimit_Click()
If chkLimit.Value = 1 Then
    cmbLimitModifier.Enabled = True
    txtLimit.Enabled = True
Else
    cmbLimitModifier.Enabled = False
    txtLimit.Enabled = False
End If
End Sub

Private Sub cmbEditor_Click()
Dim x As Integer

lblMap.Enabled = False
txtMap.Enabled = False
chkOnlyIfOn(0).Value = 0
chkOnlyIfOn(1).Value = 0
chkOnlyIfOn(0).Enabled = False
chkOnlyIfOn(1).Enabled = False
cmbOnlyIfAuxValue(0).Visible = False
cmbOnlyIfAuxValue(1).Visible = False
cmbOnlyIf(0).clear
cmbOnlyIf(1).clear
        
Select Case cmbEditor.ListIndex

    Case 0: '0 - class
        cmbField.clear
        cmbField.AddItem "Exp %", 0
        cmbField.AddItem "HP Min", 1
        cmbField.AddItem "HP Max", 2
        cmbField.AddItem "Combat", 3
        cmbField.AddItem "Give Ability", 4
        cmbField.AddItem "Take Ability", 5
        cmbField.AddItem "Change Ability", 6
        
        For x = 0 To 1
            chkOnlyIfOn(x).Enabled = True
            cmbOnlyIf(x).AddItem "Combat", 0
            cmbOnlyIf(x).AddItem "Magery", 1
            cmbOnlyIf(x).AddItem "Exp %", 2
            cmbOnlyIf(x).AddItem "HP Min", 3
            cmbOnlyIf(x).AddItem "HP Max", 4
            cmbOnlyIf(x).AddItem "Has Ability", 5
            cmbOnlyIf(x).AddItem "Doesn't have Ability", 6
            cmbOnlyIf(x).ListIndex = 0
        Next x
    Case 1: '1 - item
        cmbField.clear
        cmbField.AddItem "Game Limit", 0
        cmbField.AddItem "Weight", 1
        cmbField.AddItem "Min Hit", 2
        cmbField.AddItem "Max Hit", 3
        cmbField.AddItem "Speed", 4
        cmbField.AddItem "Req. Strength", 5
        cmbField.AddItem "AC", 6
        cmbField.AddItem "DR", 7
        cmbField.AddItem "Accuracy", 8
        cmbField.AddItem "Uses", 9
        cmbField.AddItem "Cost", 10
        cmbField.AddItem "Give Ability", 11
        cmbField.AddItem "Take Ability", 12
        cmbField.AddItem "Change Ability", 13
        
        For x = 0 To 1
            chkOnlyIfOn(x).Enabled = True
            cmbOnlyIf(x).AddItem "Game Limit", 0
            cmbOnlyIf(x).AddItem "Item Type", 1
            cmbOnlyIf(x).AddItem "Armour Type", 2
            cmbOnlyIf(x).AddItem "Weapon Type", 3
            cmbOnlyIf(x).AddItem "Worn On", 4
            cmbOnlyIf(x).AddItem "Weight", 5
            cmbOnlyIf(x).AddItem "Speed", 6
            cmbOnlyIf(x).AddItem "Req. Strength", 7
            cmbOnlyIf(x).AddItem "Accuracy", 8
            cmbOnlyIf(x).AddItem "Cost", 9
            cmbOnlyIf(x).AddItem "AC", 10
            cmbOnlyIf(x).AddItem "DR", 11
            cmbOnlyIf(x).AddItem "Min Hit", 12
            cmbOnlyIf(x).AddItem "Max Hit", 13
            cmbOnlyIf(x).AddItem "Has Ability", 14
            cmbOnlyIf(x).AddItem "Doesn't have Ability", 15
            cmbOnlyIf(x).ListIndex = 0
        Next x
        
    Case 2: '2 - mons
        cmbField.clear
        cmbField.AddItem "Exp", 0
        cmbField.AddItem "Exp Multiplier", 1
        cmbField.AddItem "Total Exp", 2
        cmbField.AddItem "MR", 3
        cmbField.AddItem "Charm LVL", 4
        cmbField.AddItem "AC", 5
        cmbField.AddItem "DR", 6
        cmbField.AddItem "Follow %", 7
        cmbField.AddItem "Regen Time", 8
        cmbField.AddItem "Game Limit", 9
        cmbField.AddItem "HPs", 10
        cmbField.AddItem "HP Regen", 11
        cmbField.AddItem "Energy", 12
        cmbField.AddItem "Runic", 13
        cmbField.AddItem "Platinum", 14
        cmbField.AddItem "Gold", 15
        cmbField.AddItem "Silver", 16
        cmbField.AddItem "Copper", 17
        cmbField.AddItem "ALL Money", 18
        cmbField.AddItem "ALL Item Drop %", 19
        cmbField.AddItem "ALL Item Uses", 20
        cmbField.AddItem "Active", 21
        cmbField.AddItem "Give Ability", 22
        cmbField.AddItem "Take Ability", 23
        cmbField.AddItem "Change Ability", 24
        
        For x = 0 To 1
            chkOnlyIfOn(x).Enabled = True
            cmbOnlyIf(x).AddItem "Game Limit", 0
            cmbOnlyIf(x).AddItem "Experience", 1
            cmbOnlyIf(x).AddItem "Regen Time", 2
            cmbOnlyIf(x).AddItem "Index (w/Group)", 3
            cmbOnlyIf(x).AddItem "Group", 4
            cmbOnlyIf(x).AddItem "Runic", 5
            cmbOnlyIf(x).AddItem "Platinum", 6
            cmbOnlyIf(x).AddItem "Gold", 7
            cmbOnlyIf(x).AddItem "Silver", 8
            cmbOnlyIf(x).AddItem "Copper", 9
            cmbOnlyIf(x).AddItem "Charm LVL", 10
            cmbOnlyIf(x).AddItem "Follow %", 11
            cmbOnlyIf(x).AddItem "MR", 12
            cmbOnlyIf(x).AddItem "HP Regen", 13
            cmbOnlyIf(x).AddItem "Hit Points", 14
            cmbOnlyIf(x).AddItem "AC", 15
            cmbOnlyIf(x).AddItem "DR", 16
            cmbOnlyIf(x).AddItem "Drop % (per item)", 17
            cmbOnlyIf(x).AddItem "Drop Uses (per item)", 18
            cmbOnlyIf(x).AddItem "Has Ability", 19
            cmbOnlyIf(x).AddItem "Doesn't have Ability", 20
            cmbOnlyIf(x).ListIndex = 0
        Next x
        
    Case 3: '3 - race
        cmbField.clear
        cmbField.AddItem "Exp %", 0
        cmbField.AddItem "Start CP", 1
        cmbField.AddItem "HP Bonus", 2
        cmbField.AddItem "Str Min", 3
        cmbField.AddItem "Str Max", 4
        cmbField.AddItem "Agil Min", 5
        cmbField.AddItem "Agil Max", 6
        cmbField.AddItem "Int Min", 7
        cmbField.AddItem "Int Max", 8
        cmbField.AddItem "Hea Min", 9
        cmbField.AddItem "Hea Max", 10
        cmbField.AddItem "Wis Min", 11
        cmbField.AddItem "Wis Max", 12
        cmbField.AddItem "Chm Min", 13
        cmbField.AddItem "Chm Max", 14
        cmbField.AddItem "ALL Min Stats", 15
        cmbField.AddItem "ALL Max Stats", 16
        cmbField.AddItem "Give Ability", 17
        cmbField.AddItem "Take Ability", 18
        cmbField.AddItem "Change Ability", 19
        
        For x = 0 To 1
            chkOnlyIfOn(x).Enabled = True
            cmbOnlyIf(x).AddItem "Has Ability", 0
            cmbOnlyIf(x).AddItem "Doesn't have Ability", 1
            cmbOnlyIf(x).ListIndex = 0
        Next x
        
    Case 4: '4 - room
        cmbField.clear
        cmbField.AddItem "Delay", 0
        cmbField.AddItem "Max Regen", 1
        cmbField.AddItem "Max Area", 2
        cmbField.AddItem "Light", 3
        cmbField.AddItem "Min Index", 4
        cmbField.AddItem "Max Index", 5
        cmbField.AddItem "GangHouse #", 6
        cmbField.AddItem "Control Room", 7
        cmbField.AddItem "Room Spell", 8
        If chkAll.Value = 0 Then
            lblMap.Enabled = True
            txtMap.Enabled = True
        End If
        
        For x = 0 To 1
            chkOnlyIfOn(x).Enabled = True
            cmbOnlyIf(x).AddItem "Room Type", 0
            cmbOnlyIf(x).AddItem "Min Index", 1
            cmbOnlyIf(x).AddItem "Max Index", 2
            cmbOnlyIf(x).AddItem "Max Regen", 3
            cmbOnlyIf(x).AddItem "Delay", 4
            cmbOnlyIf(x).AddItem "Light", 5
            cmbOnlyIf(x).AddItem "GangHouse #", 6
            cmbOnlyIf(x).AddItem "Max Area", 7
            cmbOnlyIf(x).AddItem "Control Room", 8
            cmbOnlyIf(x).AddItem "Monster Type", 9
            cmbOnlyIf(x).AddItem "Room Spell", 10
            cmbOnlyIf(x).ListIndex = 0
        Next x
        
    Case 5: '5 - shop
        cmbField.clear
        cmbField.AddItem "Min LVL", 0
        cmbField.AddItem "Max LVL", 1
        cmbField.AddItem "Markup", 2
        cmbField.AddItem "ALL Normal Stock", 3
        cmbField.AddItem "ALL Max Stock", 4
        cmbField.AddItem "ALL Regen Times", 5
        cmbField.AddItem "ALL Regen %", 6
        cmbField.AddItem "ALL Regen Amounts", 7
        
        For x = 0 To 1
            chkOnlyIfOn(x).Enabled = True
            cmbOnlyIf(x).AddItem "Shop Type", 0
            cmbOnlyIf(x).AddItem "Stock Now (per item)", 1
            cmbOnlyIf(x).AddItem "Max Stock (per item)", 2
            cmbOnlyIf(x).AddItem "Regen Time (per item)", 3
            cmbOnlyIf(x).AddItem "Regen % (per item)", 4
            cmbOnlyIf(x).AddItem "Regen # (per item)", 5
            cmbOnlyIf(x).AddItem "Markup", 6
            cmbOnlyIf(x).ListIndex = 0
        Next x

    Case 6: '6 - spell
        cmbField.clear
        cmbField.AddItem "Req LVL", 0
        cmbField.AddItem "Energy", 1
        cmbField.AddItem "Mana", 2
        cmbField.AddItem "Difficulty", 3
        cmbField.AddItem "Min", 4
        cmbField.AddItem "Max", 5
        cmbField.AddItem "Duration", 6
        cmbField.AddItem "LVL Increase Cap", 7
        cmbField.AddItem "LVLs Min Increase", 8
        cmbField.AddItem "# Min Increase", 9
        cmbField.AddItem "LVLs Max Increase", 10
        cmbField.AddItem "# Max Increase", 11
        cmbField.AddItem "LVLs Dur Increase", 12
        cmbField.AddItem "# Dur Increase", 13
        cmbField.AddItem "Give Ability", 14
        cmbField.AddItem "Take Ability", 15
        cmbField.AddItem "Change Ability", 16
        
        For x = 0 To 1
            chkOnlyIfOn(x).Enabled = True
            cmbOnlyIf(x).AddItem "Magery", 0
            cmbOnlyIf(x).AddItem "Req. Level", 1
            cmbOnlyIf(x).AddItem "Energy", 2
            cmbOnlyIf(x).AddItem "Mana", 3
            cmbOnlyIf(x).AddItem "Difficulty", 4
            cmbOnlyIf(x).AddItem "Min", 5
            cmbOnlyIf(x).AddItem "Max", 6
            cmbOnlyIf(x).AddItem "Duration", 7
            cmbOnlyIf(x).AddItem "LVL Increase Cap", 8
            cmbOnlyIf(x).AddItem "LVLs Min Increase", 9
            cmbOnlyIf(x).AddItem "# Min Increase", 10
            cmbOnlyIf(x).AddItem "LVLs Max Increase", 11
            cmbOnlyIf(x).AddItem "# Max Increase", 12
            cmbOnlyIf(x).AddItem "LVLs Dur Increase", 13
            cmbOnlyIf(x).AddItem "# Dur Increase", 14
            cmbOnlyIf(x).AddItem "Has Ability", 15
            cmbOnlyIf(x).AddItem "Doesn't have Ability", 16
            cmbOnlyIf(x).ListIndex = 0
        Next x
        
End Select

cmbField.refresh
cmbField.ListIndex = 0
End Sub
Private Sub cmbModifier_Click()

If cmbModifier.ListIndex = 4 Then '%
    txtValue.Visible = False
    cmbValue.Visible = True
Else
    txtValue.Visible = True
    cmbValue.Visible = False
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdStart_Click()
On Error GoTo error
Dim nYesNo As Integer
Dim q1 As String, q2 As String, q3 As String, q4 As String, q5 As String

bStopProcess = False
If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

If chkAll.Value = 1 Then
    q2 = "each record in the " & cmbEditor.Text & " database"
Else
    If cmbEditor.ListIndex = 4 Then 'rooms
        q2 = "Map " & Val(txtMap.Text) & ", rooms " & Val(txtR1.Text) & " to " & Val(txtR2.Text) & " of the " & cmbEditor.Text & " database"
    Else
        q2 = "records " & Val(txtR1.Text) & " to " & Val(txtR2.Text) & " of the " & cmbEditor.Text & " database"
    End If
End If
     
If cmbAbilities.Enabled = True Then
    q1 = cmbField.Text & " """ & cmbAbilities.Text & """"
    
    If cmbModifier.Enabled = True Then
        Select Case cmbModifier.ListIndex
            Case 0: q1 = q1 & " by adding " & Val(txtValue.Text) & " to the current ability value"
            Case 1: q1 = q1 & " by subtracting " & Val(txtValue.Text) & " from the current ability value"
            Case 2: q1 = q1 & " by multiplying " & Val(txtValue.Text) & " to the current ability value"
            Case 3: q1 = q1 & " by dividing " & Val(txtValue.Text) & " from the current ability value"
            Case 4: q1 = q1 & " by setting the current ability value to " & Val(cmbValue.Text) & "% of it's original value"
            Case 5: q1 = q1 & " by setting the current ability value equal to " & Val(cmbValue.Text)
        End Select
    Else
        If txtValue.Enabled = True Then
            q1 = q1 & " with a value of " & Val(txtValue.Text)
        Else
            q1 = q1 & " (no matter what the value is)"
        End If
    End If
    
    q1 = q1 & " for "
Else
    Select Case cmbModifier.ListIndex
        Case 0: q1 = "add " & Val(txtValue.Text) & " to the '" & cmbField.Text & "' for "
        Case 1: q1 = "subtract " & Val(txtValue.Text) & " from the '" & cmbField.Text & "' for "
        Case 2: q1 = "multiply the '" & cmbField.Text & "' by " & Val(txtValue.Text) & " for "
        Case 3: q1 = "divide the '" & cmbField.Text & "' by " & Val(txtValue.Text) & " for "
        Case 4: q1 = "set the '" & cmbField.Text & "' to " & Val(cmbValue.Text) & "% of it's original value for "
        Case 5:
            q1 = "set the '" & cmbField.Text & "' for "
            q2 = q2 & " to " & Val(txtValue.Text)
    End Select
End If

q3 = ""
If chkLimit.Value = 1 Then
    q3 = vbCrLf & vbCrLf & "The end result will be limited to having a value"
    Select Case cmbLimitModifier.ListIndex
        Case 0: '<=
            q3 = q3 & " less than or equal to "
        Case 1: '>=
            q3 = q3 & " greater than or equal to "
    End Select
    q3 = q3 & Val(txtLimit.Text) & "."
End If

q4 = ""
If chkOnlyIfOn(0).Value = 1 Then
        q4 = vbCrLf & vbCrLf & "This action will only be performed if "
        
        If Not InStr(1, cmbOnlyIf(0).Text, "bility") = 0 Then
            q4 = q4 & "it " & cmbOnlyIf(0).Text & " """ & cmbOnlyIfAuxValue(0).Text & """ with a value"
        Else
            q4 = q4 & "the "
            
            If cmbOnlyIfAuxValue(0).Visible Then
                If lblOnlyIfAuxValue(0).Visible Then
                    q4 = q4 & lblOnlyIfAuxValue(0).Caption & " is " & cmbOnlyIfAuxValue(0).Text & ", and " & cmbOnlyIf(0).Text & " is"
                Else
                    q4 = q4 & cmbOnlyIf(0).Text & " is " & cmbOnlyIfAuxValue(0).Text & " and"
                End If
            Else
                q4 = q4 & cmbOnlyIf(0).Text & " is"
            End If
        End If
        
        Select Case cmbOnlyIfModifier(0).ListIndex
            Case 0: '=
                q4 = q4 & " equal to "
            Case 1: '<=
                q4 = q4 & " less than or equal to "
            Case 2: '>=
                q4 = q4 & " greater than or equal to "
            Case 3: 'NOT =
                q4 = q4 & " NOT equal to "
        End Select
        
        If cmbOnlyIfValue(0).Visible = True Then
            q4 = q4 & cmbOnlyIfValue(0).Text
        Else
            q4 = q4 & Val(txtOnlyIfValue(0).Text)
        End If
        
        q4 = q4 & "."
End If

q5 = ""
If chkOnlyIfOn(1).Value = 1 Then
        q5 = vbCrLf & vbCrLf & "This action will only be performed if "
        
        If Not InStr(1, cmbOnlyIf(1).Text, "bility") = 0 Then
            q5 = q5 & "it " & cmbOnlyIf(1).Text & " """ & cmbOnlyIfAuxValue(1).Text & """ with a value"
        Else
            q5 = q5 & "the "
            
            If cmbOnlyIfAuxValue(1).Visible Then
                If lblOnlyIfAuxValue(1).Visible Then
                    q5 = q5 & lblOnlyIfAuxValue(1).Caption & " is " & cmbOnlyIfAuxValue(1).Text & ", and " & cmbOnlyIf(1).Text & " is"
                Else
                    q5 = q5 & cmbOnlyIf(1).Text & " is " & cmbOnlyIfAuxValue(1).Text & " and"
                End If
            Else
                q5 = q5 & cmbOnlyIf(1).Text & " is"
            End If
        End If
        
        Select Case cmbOnlyIfModifier(1).ListIndex
            Case 0: '=
                q5 = q5 & " equal to "
            Case 1: '<=
                q5 = q5 & " less than or equal to "
            Case 2: '>=
                q5 = q5 & " greater than or equal to "
            Case 3: 'NOT =
                q5 = q5 & " NOT equal to "
        End Select
        
        If cmbOnlyIfValue(1).Visible = True Then
            q5 = q5 & cmbOnlyIfValue(1).Text
        Else
            q5 = q5 & Val(txtOnlyIfValue(1).Text)
        End If
        
        q5 = q5 & "."
End If

nYesNo = MsgBox("Are you sure you want to " & q1 & q2 & "?" & q3 & q4 & q5, vbYesNo + vbQuestion, "Are you sure?")
If nYesNo <> vbYes Then Exit Sub

Call UnloadForms(Me.Name)

frmMain.Enabled = False
cmdStart.Enabled = False
cmdCancel.Enabled = False

If Right(App.Path, 1) = "\" Then
    sLogFile = App.Path & "NMR-Log_Universal.txt"
Else
    sLogFile = App.Path & "\NMR-Log_Universal.txt"
End If

If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(sLogFile) Then Call fso.DeleteFile(sLogFile, True)
Set ts = fso.OpenTextFile(sLogFile, ForWriting, True)

ts.WriteLine ("Universal Modifier job started " & Date & " @ " & Time)
ts.WriteLine q1 & q2
If Not q3 = "" Then ts.WriteLine RemoveCharacter(RemoveCharacter(q3, vbCr), vbLf)
If Not q4 = "" Then ts.WriteLine RemoveCharacter(RemoveCharacter(q4, vbCr), vbLf)
ts.WriteBlankLines (1)

DoEvents
Select Case cmbEditor.ListIndex
    Case 0 'class
        Call ModifyClass(cmbField.ListIndex, Val(txtR1.Text), Val(txtR2.Text))
    Case 1 'item
        Call ModifyItem(cmbField.ListIndex, Val(txtR1.Text), Val(txtR2.Text))
    Case 2 'monster
        If eDatFileVersion < v111j And cmbField.ListIndex = 1 Then
            MsgBox "The 'Exp Multiplier' field is only available when using v1.11J dats or newer"
            GoTo quit:
        ElseIf eDatFileVersion < v111j And cmbField.ListIndex = 2 Then
            MsgBox "The 'Total Exp' option is only available when using v1.11J dats or newer"
            GoTo quit:
        End If
        Call ModifyMonster(cmbField.ListIndex, Val(txtR1.Text), Val(txtR2.Text))
    Case 3 'race
        Call ModifyRace(cmbField.ListIndex, Val(txtR1.Text), Val(txtR2.Text))
    Case 4 'room
        Call ModifyRoom(cmbField.ListIndex, Val(txtR1.Text), Val(txtR2.Text))
    Case 5 'shop
        Call ModifyShop(cmbField.ListIndex, Val(txtR1.Text), Val(txtR2.Text))
    Case 6 'spell
        Call ModifySpell(cmbField.ListIndex, Val(txtR1.Text), Val(txtR2.Text))
End Select

ts.WriteBlankLines (1)
ts.WriteLine ("Complete - " & Date & " @ " & Time)
ts.Close

Set ts = Nothing

nYesNo = MsgBox("Complete, View Log?", vbInformation + vbYesNo + vbDefaultButton1)
StatusBar.Panels(1).Text = ""
StatusBar.Panels(2).Text = ""
If nYesNo = vbYes Then Call cmdLog_Click

quit:
frmMain.Enabled = True
cmdStart.Enabled = True
cmdCancel.Enabled = True

Exit Sub
error:
frmMain.Enabled = True
cmdStart.Enabled = True
cmdCancel.Enabled = True

Call HandleError
End Sub
Private Function TestOnlyIf(ByVal nValue As Double, nOnlyIfIndex As Integer) As Boolean
Dim nTest As Variant

'true = pass
'false = fail

TestOnlyIf = False

If cmbOnlyIfValue(nOnlyIfIndex).Visible Then 'combo value
    nTest = cmbOnlyIfValue(nOnlyIfIndex).ListIndex
Else
    nTest = Val(txtOnlyIfValue(nOnlyIfIndex).Text)
End If

Select Case cmbOnlyIfModifier(nOnlyIfIndex).ListIndex
    Case 0: '=
        If nValue = nTest Then TestOnlyIf = True
    Case 1: '<=
        If nValue <= nTest Then TestOnlyIf = True
    Case 2: '>=
        If nValue >= nTest Then TestOnlyIf = True
    Case 3: 'NOT =
        If Not nValue = nTest Then TestOnlyIf = True
End Select

End Function
Private Function TestOnlyIfAux(ByVal nValue As Double, nOnlyIfIndex As Integer) As Boolean

TestOnlyIfAux = False

If nValue = cmbOnlyIfAuxValue(nOnlyIfIndex).ListIndex Then TestOnlyIfAux = True

End Function

Private Function DoMath(ByVal nValue As Variant, ByVal dataType As eDataType, SU As eSU, _
    ByVal sRecName As String, Optional ByVal nIndivNum As Integer = -1) As Variant
On Error GoTo error:

'check divide by 0
If cmbModifier.ListIndex = 3 And Val(txtValue.Text) = 0 Then
    MsgBox "Can't divide by zero buddy! :)", vbExclamation + vbOKOnly
    bStopProcess = True
    'Unload Me
    Exit Function
End If

'do the math
Select Case cmbModifier.ListIndex
    Case 0: '+
        DoMath = nValue + Val(txtValue.Text)
    Case 1: '-
        DoMath = nValue - Val(txtValue.Text)
    Case 2: '*
        DoMath = nValue * Val(txtValue.Text)
    Case 3: '/
        DoMath = nValue / Val(txtValue.Text)
    Case 4: '%
        DoMath = nValue * (Val(cmbValue.Text) / 100)
    Case 5: '=
        DoMath = Val(txtValue.Text)
End Select

'check if it's within the limit
If chkLimit.Value = 1 Then
    Select Case cmbLimitModifier.ListIndex
        Case 0: '<=
            If DoMath > Val(txtLimit.Text) Then DoMath = Val(txtLimit.Text)
        Case 1: '>=
            If DoMath < Val(txtLimit.Text) Then DoMath = Val(txtLimit.Text)
    End Select
End If

DoMath = Round(DoMath)

'make sure it meets the datatype specification
Select Case dataType
    Case 1: 'int
        Select Case SU
            Case 1: 'signed
                If DoMath > MaxInt Then DoMath = MaxInt
                If DoMath < (0 - MaxInt) - 1 Then DoMath = (0 - MaxInt) - 1
            Case 2: 'unsigned
                If DoMath > IntOffset - 1 Then DoMath = IntOffset - 1
                If DoMath < 0 Then DoMath = 0
                DoMath = UInt2SInt(DoMath)
        End Select
    Case 2: 'long
        Select Case SU
            Case 1: 'signed
                If DoMath > MaxLong Then DoMath = MaxLong
                If DoMath < (0 - MaxLong) - 1 Then DoMath = (0 - MaxLong) - 1
            Case 2: 'unsigned
                If DoMath > LongOffset - 1 Then DoMath = LongOffset - 1
                If DoMath < 0 Then DoMath = 0
                DoMath = ULong2SLong(DoMath)
        End Select
    Case 3: 'byte
        If DoMath > 255 Then DoMath = 255
        If DoMath < 0 Then DoMath = 0
End Select

If nValue = DoMath Then
    If chkOnlyChanges.Value = 0 Then ts.WriteLine sRecName & IIf(nIndivNum >= 0, " (" & nIndivNum & ")", "") & ": No Change"
Else
    ts.WriteLine sRecName & IIf(nIndivNum >= 0, " (" & nIndivNum & ")", "") _
        & ": " & nValue & " ==> " & DoMath
End If

Exit Function
error:
Call HandleError("DoMath")
End Function
Private Sub ModifySpell(ByVal nIndex As Integer, ByVal nRange1 As Long, ByVal nRange2 As Long)
Dim nStatus As Integer
Dim x As Integer, bSkip As Boolean, nOnlyIfIndex As Integer
Dim sName As String

StatusBar.Panels(1).Text = cmbEditor.Text & " - " & cmbField.Text

If chkAll.Value = 1 Then
    nStatus = BTRCALL(BGETFIRST, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Error Getting First Spell Record, exiting. - " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
Else
    nStatus = BTRCALL(BGETEQUAL, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), nRange1, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Error Getting Spell Record # " & nRange1 & ", exiting. - " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
End If

Do While nStatus = 0 And bStopProcess = False
    SpellRowToStruct Spelldatabuf.buf
    
    sName = Spellrec.Number & "-" & ClipNull(Spellrec.Name)
    
    If chkAll.Value = 0 Then
        If Spellrec.Number > nRange2 Then Exit Do
    End If
    
    For nOnlyIfIndex = 0 To 1
        If chkOnlyIfOn(nOnlyIfIndex).Value = 1 Then
            Select Case cmbOnlyIf(nOnlyIfIndex).ListIndex
                Case 0: 'magery
                    bSkip = TestOnlyIfAux(Spellrec.MageryA, nOnlyIfIndex)
                    If bSkip = False Then GoTo Skip:
                    bSkip = TestOnlyIf(Spellrec.MageryB, nOnlyIfIndex)
                
                Case 1: 'req level
                    bSkip = TestOnlyIf(Spellrec.Level, nOnlyIfIndex)
                Case 2: '"Energy", 2
                    bSkip = TestOnlyIf(Spellrec.Energy, nOnlyIfIndex)
                Case 3: '"Mana", 3
                    bSkip = TestOnlyIf(Spellrec.Mana, nOnlyIfIndex)
                Case 4: '"Difficulty", 4
                    bSkip = TestOnlyIf(Spellrec.Difficulty, nOnlyIfIndex)
                Case 5: '"Min", 5
                    bSkip = TestOnlyIf(Spellrec.Min, nOnlyIfIndex)
                Case 6: '"Max", 6
                    bSkip = TestOnlyIf(Spellrec.Max, nOnlyIfIndex)
                Case 7: '"Duration", 7
                    bSkip = TestOnlyIf(Spellrec.duration, nOnlyIfIndex)
                Case 8: '"LVL Increase Cap", 8
                    bSkip = TestOnlyIf(Spellrec.LevelCap, nOnlyIfIndex)
                Case 9: '"LVLs Min Increase", 9
                    bSkip = TestOnlyIf(Spellrec.LVLSMinIncr, nOnlyIfIndex)
                Case 10: '"# Min Increase", 10
                    bSkip = TestOnlyIf(Spellrec.MinIncrease, nOnlyIfIndex)
                Case 11: '"LVLs Max Increase", 11
                    bSkip = TestOnlyIf(Spellrec.LVLSMaxIncr, nOnlyIfIndex)
                Case 12: '"# Max Increase", 12
                    bSkip = TestOnlyIf(Spellrec.MaxIncrease, nOnlyIfIndex)
                Case 13: '"LVLs Dur Increase", 13
                    bSkip = TestOnlyIf(Spellrec.LVLSDurIncr, nOnlyIfIndex)
                Case 14: '"# Dur Increase", 14
                    bSkip = TestOnlyIf(Spellrec.DurIncrease, nOnlyIfIndex)
                Case 15: 'has abil
                    For x = 0 To 9
                        bSkip = TestOnlyIfAux(Spellrec.AbilityA(x), nOnlyIfIndex)
                        If bSkip = True Then
                            bSkip = TestOnlyIf(Spellrec.AbilityB(x), nOnlyIfIndex)
                            If bSkip = True Then Exit For
                        End If
                    Next x
                Case 16: 'doesn't have abil
                    For x = 0 To 9
                        bSkip = TestOnlyIfAux(Spellrec.AbilityA(x), nOnlyIfIndex)
                        If bSkip = True Then
                            bSkip = TestOnlyIf(Spellrec.AbilityB(x), nOnlyIfIndex)
                            If bSkip = True Then 'true because if it matches then it 'HAS' it, so skip
                                bSkip = False
                                Exit For
                            End If
                        End If
                        bSkip = True
                    Next x
            End Select
            If bSkip = False Then
                If chkOnlyChanges.Value = 0 Then ts.WriteLine sName & " ==> Skipped due to 'Only If' Setting"
                GoTo Skip:
            End If
        End If
    Next nOnlyIfIndex
    
    StatusBar.Panels(2).Text = Spellrec.Number
    Select Case nIndex
        Case 0: Spellrec.Level = DoMath(Spellrec.Level, 1, 1, sName)
        Case 1: Spellrec.Energy = DoMath(Spellrec.Energy, 1, 1, sName)
        Case 2: Spellrec.Mana = DoMath(Spellrec.Mana, 1, 1, sName)
        Case 3: Spellrec.Difficulty = DoMath(Spellrec.Difficulty, 1, 1, sName)
        Case 4: Spellrec.Min = DoMath(Spellrec.Min, 1, 1, sName)
        Case 5: Spellrec.Max = DoMath(Spellrec.Max, 1, 1, sName)
        Case 6: Spellrec.duration = DoMath(Spellrec.duration, 1, 1, sName)
        Case 7: Spellrec.LevelCap = DoMath(Spellrec.LevelCap, 3, 1, sName)
        Case 8: Spellrec.LVLSMinIncr = DoMath(Spellrec.LVLSMinIncr, 3, 1, sName)
        Case 9: Spellrec.MinIncrease = DoMath(Spellrec.MinIncrease, 3, 1, sName)
        Case 10: Spellrec.LVLSMaxIncr = DoMath(Spellrec.LVLSMaxIncr, 3, 1, sName)
        Case 11: Spellrec.MaxIncrease = DoMath(Spellrec.MaxIncrease, 3, 1, sName)
        Case 12: Spellrec.LVLSDurIncr = DoMath(Spellrec.LVLSDurIncr, 3, 1, sName)
        Case 13: Spellrec.DurIncrease = DoMath(Spellrec.DurIncrease, 3, 1, sName)
        Case 14: 'give
            For x = 0 To 9
                If Spellrec.AbilityA(x) = cmbAbilities.ItemData(cmbAbilities.ListIndex) Then
                    If Spellrec.AbilityB(x) = Val(txtValue.Text) Then
                        If chkOnlyChanges.Value = 0 Then ts.WriteLine sName & ": Already has ability"
                        GoTo Skip:
                    End If
                End If
            Next
            For x = 0 To 9
                If Spellrec.AbilityA(x) = 0 Then
                    Spellrec.AbilityA(x) = cmbAbilities.ItemData(cmbAbilities.ListIndex)
                    Spellrec.AbilityB(x) = Val(txtValue.Text)
                    ts.WriteLine sName & ": Ability added to slot " & x
                    Exit For
                End If
            Next
            If chkOnlyChanges.Value = 0 Then If x = 10 Then ts.WriteLine sName & ": No Empty Slot to Add Ability"
        Case 15: 'take
            bSkip = False
            For x = 0 To 9
                If Spellrec.AbilityA(x) = cmbAbilities.ItemData(cmbAbilities.ListIndex) Then
                    Spellrec.AbilityA(x) = 0
                    ts.WriteLine sName & ": Ability taken from slot " & x
                    bSkip = True
                End If
            Next
            If chkOnlyChanges.Value = 0 And Not bSkip Then ts.WriteLine sName & ": Ability not found"
        Case 16: 'change
            For x = 0 To 9
                If Spellrec.AbilityA(x) = cmbAbilities.ItemData(cmbAbilities.ListIndex) Then
                    Spellrec.AbilityB(x) = DoMath(Spellrec.AbilityB(x), dtInteger, suSigned, sName)
                End If
            Next
            If chkOnlyChanges.Value = 0 Then If x = 10 Then ts.WriteLine sName & ": Ability not found"
    End Select
    
    nStatus = UpdateSpell
    If Not nStatus = 0 Then
        MsgBox "Error Updating Spell Record # " & Spellrec.Number & ", exiting. - " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
Skip:
    nStatus = BTRCALL(BGETNEXT, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

Exit Sub

End Sub
Private Sub ModifyShop(ByVal nIndex As Integer, ByVal nRange1 As Long, ByVal nRange2 As Long)
Dim nStatus As Integer, bSkip As Boolean
Dim x As Integer, nOnlyIfIndex As Integer
Dim sName As String

On Error GoTo error:

StatusBar.Panels(1).Text = cmbEditor.Text & " - " & cmbField.Text

If chkAll.Value = 1 Then
    nStatus = BTRCALL(BGETFIRST, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Error Getting First Shop Record, exiting. - " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
Else
    nStatus = BTRCALL(BGETEQUAL, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), nRange1, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Error Getting Shop Record # " & nRange1 & ", exiting. - " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
End If

Do While nStatus = 0 And bStopProcess = False
    ShopRowToStruct Shopdatabuf.buf
    
    sName = Shoprec.Number & "-" & ClipNull(Shoprec.Name)
    
    If chkAll.Value = 0 Then
        If Shoprec.Number > nRange2 Then Exit Do
    End If
    
    For nOnlyIfIndex = 0 To 1
        If chkOnlyIfOn(nOnlyIfIndex).Value = 1 Then
            bSkip = True
            Select Case cmbOnlyIf(nOnlyIfIndex).ListIndex
                Case 0: 'shop type
                    bSkip = TestOnlyIf(Shoprec.ShopType, nOnlyIfIndex)
                Case 6: 'markup
                    bSkip = TestOnlyIf(Shoprec.ShopMarkUp, nOnlyIfIndex)
            End Select
            If bSkip = False Then
                If chkOnlyChanges.Value = 0 Then ts.WriteLine sName & " ==> Skipped due to 'Only If' Setting"
                GoTo Skip:
            End If
        End If
    Next nOnlyIfIndex
    
    StatusBar.Panels(2).Text = Shoprec.Number
    Select Case nIndex
        Case 0: Shoprec.ShopMinLvL = DoMath(Shoprec.ShopMinLvL, 1, 1, sName)
        Case 1: Shoprec.ShopMaxLvl = DoMath(Shoprec.ShopMaxLvl, 1, 1, sName)
        Case 2: Shoprec.ShopMarkUp = DoMath(Shoprec.ShopMarkUp, 1, 1, sName)
        
        Case 3, 4, 5, 6, 7: 'now/max/time/%/#
        
            For x = 0 To 19
                
                For nOnlyIfIndex = 0 To 1
                    If chkOnlyIfOn(nOnlyIfIndex).Value = 1 Then 'per item checks
                        bSkip = True
                        Select Case cmbOnlyIf(nOnlyIfIndex).ListIndex
                            Case 1: 'now
                                bSkip = TestOnlyIf(Shoprec.ShopNow(x), nOnlyIfIndex)
                            Case 2: 'max
                                bSkip = TestOnlyIf(Shoprec.ShopMax(x), nOnlyIfIndex)
                            Case 3: 'time
                                bSkip = TestOnlyIf(Shoprec.ShopRgnTime(x), nOnlyIfIndex)
                            Case 4: '%
                                bSkip = TestOnlyIf(Shoprec.ShopRgnPercentage(x), nOnlyIfIndex)
                            Case 5: '#
                                bSkip = TestOnlyIf(Shoprec.ShopRgnNumber(x), nOnlyIfIndex)
                                
                        End Select
                        If bSkip = False Then
                            If chkOnlyChanges.Value = 0 Then ts.WriteLine sName & " (" & x & ")" _
                                & " ==> Skipped due to 'Only If' Setting"
                            GoTo SkipIndiv:
                        End If
                    End If
                Next nOnlyIfIndex
                
                Select Case nIndex
                    Case 3: Shoprec.ShopNow(x) = DoMath(Shoprec.ShopNow(x), 1, 1, sName, x)
                    Case 4: Shoprec.ShopMax(x) = DoMath(Shoprec.ShopMax(x), 1, 1, sName, x)
                    Case 5: Shoprec.ShopRgnTime(x) = DoMath(Shoprec.ShopRgnTime(x), 1, 1, sName, x)
                    Case 6: Shoprec.ShopRgnPercentage(x) = DoMath(Shoprec.ShopRgnPercentage(x), 1, 1, sName, x)
                    Case 7: Shoprec.ShopRgnNumber(x) = DoMath(Shoprec.ShopRgnNumber(x), 1, 1, sName, x)
                End Select
SkipIndiv:
            Next x
                    
    End Select
    nStatus = UpdateShop
    If Not nStatus = 0 Then MsgBox "Error Updating Shop Record # " & Shoprec.Number & ", exiting. - " & BtrieveErrorCode(nStatus): Exit Sub
Skip:
    nStatus = BTRCALL(BGETNEXT, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

Exit Sub
error:
Call HandleError("ModifyShop")

End Sub

Private Sub ModifyRoom(ByVal nIndex As Integer, ByVal nRange1 As Long, ByVal nRange2 As Long)
Dim nStatus As Integer, bSkip As Boolean
Dim i As Long, nOnlyIfIndex As Integer
Dim sName As String

StatusBar.Panels(1).Text = cmbEditor.Text & " - " & cmbField.Text

If chkAll.Value = 1 Then
    nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Error Getting First Room Record, exiting. - " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
Else
    i = nRange1
    RoomKeyStruct.MapNum = Val(txtMap.Text)
    RoomKeyStruct.RoomNum = i
    
    nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), RoomKeyStruct, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Error Getting Room # " & nRange1 & ", exiting. - " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
End If

Do While nStatus = 0 And bStopProcess = False
    RoomRowToStruct Roomdatabuf.buf
    
    sName = Roomrec.MapNumber & "/" & Roomrec.RoomNumber & " - " & ClipNull(Roomrec.Name)
    
    For nOnlyIfIndex = 0 To 1
        If chkOnlyIfOn(nOnlyIfIndex).Value = 1 Then
            Select Case cmbOnlyIf(nOnlyIfIndex).ListIndex
                Case 0: 'room type
                    bSkip = TestOnlyIf(Roomrec.Type, nOnlyIfIndex)
                Case 1:
                    bSkip = TestOnlyIf(Roomrec.MinIndex, nOnlyIfIndex)
                Case 2:
                    bSkip = TestOnlyIf(Roomrec.MaxIndex, nOnlyIfIndex)
                Case 3:
                    bSkip = TestOnlyIf(Roomrec.MaxRegen, nOnlyIfIndex)
                Case 4:
                    bSkip = TestOnlyIf(Roomrec.Delay, nOnlyIfIndex)
                Case 5:
                    bSkip = TestOnlyIf(Roomrec.Light, nOnlyIfIndex)
                Case 6:
                    bSkip = TestOnlyIf(Roomrec.GangHouseNumber, nOnlyIfIndex)
                Case 7:
                    bSkip = TestOnlyIf(Roomrec.MaxArea, nOnlyIfIndex)
                Case 8:
                    bSkip = TestOnlyIf(Roomrec.ControlRoom, nOnlyIfIndex)
                Case 9:
                    bSkip = TestOnlyIf(Roomrec.MonsterType, nOnlyIfIndex)
                Case 10:
                    bSkip = TestOnlyIf(Roomrec.Spell, nOnlyIfIndex)
            End Select
            If bSkip = False Then
                If chkOnlyChanges.Value = 0 Then ts.WriteLine sName & " ==> Skipped due to 'Only If' Setting"
                GoTo Skip:
            End If
        End If
    Next nOnlyIfIndex
    
    StatusBar.Panels(2).Text = Roomrec.RoomNumber
    Select Case nIndex
        Case 0: Roomrec.Delay = DoMath(Roomrec.Delay, 1, 1, sName)
        Case 1: Roomrec.MaxRegen = DoMath(Roomrec.MaxRegen, 2, 1, sName)
        Case 2: Roomrec.MaxArea = DoMath(Roomrec.MaxArea, 1, 1, sName)
        Case 3: Roomrec.Light = DoMath(Roomrec.Light, 1, 1, sName)
        Case 4: Roomrec.MinIndex = DoMath(Roomrec.MinIndex, dtInteger, 1, sName)
        Case 5: Roomrec.MaxIndex = DoMath(Roomrec.MaxIndex, dtInteger, 1, sName)
        Case 6: Roomrec.GangHouseNumber = DoMath(Roomrec.GangHouseNumber, dtInteger, suSigned, sName)
        Case 7: Roomrec.ControlRoom = DoMath(Roomrec.ControlRoom, dtLong, suUnsigned, sName)
        Case 8: Roomrec.Spell = DoMath(Roomrec.Spell, dtLong, suUnsigned, sName)
    End Select
    
    
    nStatus = UpdateRoom
    If Not nStatus = 0 Then
        MsgBox "Error Updating Room # " & Roomrec.RoomNumber & ", exiting. - " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
    
Skip:
    
    If chkAll.Value = 1 Then
        nStatus = BTRCALL(BGETNEXT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    Else
GetNext:
        i = i + 1
        If i > nRange2 Then Exit Do
        RoomKeyStruct.RoomNum = i
        
        nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), RoomKeyStruct, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            If nStatus = 4 Then GoTo GetNext:
            MsgBox "Error getting Room # " & i & ", exiting. - " & BtrieveErrorCode(nStatus)
            Exit Sub
        End If
    End If
    If Not bUseCPU Then DoEvents
Loop

Exit Sub

End Sub
Private Sub ModifyRace(ByVal nIndex As Integer, ByVal nRange1 As Long, ByVal nRange2 As Long)
Dim nStatus As Integer, x As Integer, bSkip As Boolean
Dim sName As String, nOnlyIfIndex As Integer

StatusBar.Panels(1).Text = cmbEditor.Text & " - " & cmbField.Text

If chkAll.Value = 1 Then
    nStatus = BTRCALL(BGETFIRST, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Error Getting First Race Record, exiting. - " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
Else
    nStatus = BTRCALL(BGETEQUAL, RacePosBlock, Racedatabuf, Len(Racedatabuf), nRange1, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Error Getting Race Record # " & nRange1 & ", exiting. - " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
End If

Do While nStatus = 0 And bStopProcess = False
    RaceRowToStruct Racedatabuf.buf
    
    sName = Racerec.Number & "-" & ClipNull(Racerec.Name)
    
    If chkAll.Value = 0 Then
        If Racerec.Number > nRange2 Then Exit Do
    End If
    
    For nOnlyIfIndex = 0 To 1
        If chkOnlyIfOn(nOnlyIfIndex).Value = 1 Then
            Select Case cmbOnlyIf(nOnlyIfIndex).ListIndex
                Case 0: 'has abil
                    For x = 0 To 9
                        bSkip = TestOnlyIfAux(Racerec.AbilityA(x), nOnlyIfIndex)
                        If bSkip = True Then
                            bSkip = TestOnlyIf(Racerec.AbilityB(x), nOnlyIfIndex)
                            If bSkip = True Then Exit For
                        End If
                    Next x
                Case 1: 'doesn't have abil
                    For x = 0 To 9
                        bSkip = TestOnlyIfAux(Racerec.AbilityA(x), nOnlyIfIndex)
                        If bSkip = True Then
                            bSkip = TestOnlyIf(Racerec.AbilityB(x), nOnlyIfIndex)
                            If bSkip = True Then 'true because if it matches then it 'HAS' it, so skip
                                bSkip = False
                                Exit For
                            End If
                        End If
                        bSkip = True
                    Next x
            End Select
            If bSkip = False Then
                If chkOnlyChanges.Value = 0 Then ts.WriteLine sName & " ==> Skipped due to 'Only If' Setting"
                GoTo Skip:
            End If
        End If
    Next nOnlyIfIndex
    
    StatusBar.Panels(2).Text = Racerec.Number
    
    Select Case nIndex
        Case 0: Racerec.ExpChart = DoMath(Racerec.ExpChart, 1, 1, sName)
        Case 1: Racerec.CP = DoMath(Racerec.CP, 1, 1, sName)
        Case 2: Racerec.HPBonus = DoMath(Racerec.HPBonus, 1, 1, sName)
        Case 3: Racerec.MinStr = DoMath(Racerec.MinStr, 1, 1, sName)
        Case 4: Racerec.MaxStr = DoMath(Racerec.MaxStr, 1, 1, sName)
        Case 5: Racerec.MinAgl = DoMath(Racerec.MinAgl, 1, 1, sName)
        Case 6: Racerec.MaxAgl = DoMath(Racerec.MaxAgl, 1, 1, sName)
        Case 7: Racerec.MinInt = DoMath(Racerec.MinInt, 1, 1, sName)
        Case 8: Racerec.MaxInt = DoMath(Racerec.MaxInt, 1, 1, sName)
        Case 9: Racerec.MinHea = DoMath(Racerec.MinHea, 1, 1, sName)
        Case 10: Racerec.MaxHea = DoMath(Racerec.MaxHea, 1, 1, sName)
        Case 11: Racerec.MinWil = DoMath(Racerec.MinWil, 1, 1, sName)
        Case 12: Racerec.MaxWil = DoMath(Racerec.MaxWil, 1, 1, sName)
        Case 13: Racerec.MinChm = DoMath(Racerec.MinChm, 1, 1, sName)
        Case 14: Racerec.MaxChm = DoMath(Racerec.MaxChm, 1, 1, sName)
        Case 15:
            Racerec.MinStr = DoMath(Racerec.MinStr, 1, 1, sName)
            Racerec.MinAgl = DoMath(Racerec.MinAgl, 1, 1, sName)
            Racerec.MinInt = DoMath(Racerec.MinInt, 1, 1, sName)
            Racerec.MinHea = DoMath(Racerec.MinHea, 1, 1, sName)
            Racerec.MinWil = DoMath(Racerec.MinWil, 1, 1, sName)
            Racerec.MinChm = DoMath(Racerec.MinChm, 1, 1, sName)
        Case 16:
            Racerec.MaxStr = DoMath(Racerec.MaxStr, 1, 1, sName)
            Racerec.MaxAgl = DoMath(Racerec.MaxAgl, 1, 1, sName)
            Racerec.MaxInt = DoMath(Racerec.MaxInt, 1, 1, sName)
            Racerec.MaxHea = DoMath(Racerec.MaxHea, 1, 1, sName)
            Racerec.MaxWil = DoMath(Racerec.MaxWil, 1, 1, sName)
            Racerec.MaxChm = DoMath(Racerec.MaxChm, 1, 1, sName)
        Case 17: 'give
            For x = 0 To 9
                If Racerec.AbilityA(x) = cmbAbilities.ItemData(cmbAbilities.ListIndex) Then
                    If Racerec.AbilityB(x) = Val(txtValue.Text) Then
                        If chkOnlyChanges.Value = 0 Then ts.WriteLine sName & ": Already has ability"
                        GoTo Skip:
                    End If
                End If
            Next
            For x = 0 To 9
                If Racerec.AbilityA(x) = 0 Then
                    Racerec.AbilityA(x) = cmbAbilities.ItemData(cmbAbilities.ListIndex)
                    Racerec.AbilityB(x) = Val(txtValue.Text)
                    ts.WriteLine sName & ": Ability added to slot " & x
                    Exit For
                End If
            Next
            If chkOnlyChanges.Value = 0 Then If x = 10 Then ts.WriteLine sName & ": No Empty Slot to Add Ability"
        Case 18: 'take
            bSkip = False
            For x = 0 To 9
                If Racerec.AbilityA(x) = cmbAbilities.ItemData(cmbAbilities.ListIndex) Then
                    Racerec.AbilityA(x) = 0
                    ts.WriteLine sName & ": Ability taken from slot " & x
                    bSkip = True
                End If
            Next
            If chkOnlyChanges.Value = 0 And Not bSkip Then ts.WriteLine sName & ": Ability not found"
        Case 19: 'change
            For x = 0 To 9
                If Racerec.AbilityA(x) = cmbAbilities.ItemData(cmbAbilities.ListIndex) Then
                    Racerec.AbilityB(x) = DoMath(Racerec.AbilityB(x), dtInteger, suSigned, sName)
                End If
            Next
            If chkOnlyChanges.Value = 0 Then If x = 10 Then ts.WriteLine sName & ": Ability not found"
    End Select
    nStatus = UpdateRace
    If Not nStatus = 0 Then MsgBox "Error Updating Race Record # " & Racerec.Number & ", exiting. - " & BtrieveErrorCode(nStatus): Exit Sub
Skip:
    nStatus = BTRCALL(BGETNEXT, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

Exit Sub

End Sub

Private Sub ModifyMonster(ByVal nIndex As Integer, ByVal nRange1 As Long, ByVal nRange2 As Long)
Dim nStatus As Integer, x As Integer
Dim nBase As Double, nMulti As Double, nMultiMax As Double, y As Double, temp As Double, bSkip As Boolean
Dim sName As String, nOnlyIfIndex As Integer

StatusBar.Panels(1).Text = cmbEditor.Text & " - " & cmbField.Text

If chkAll.Value = 1 Then
    nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Error Getting First Monster Record, exiting. - " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
Else
    nStatus = BTRCALL(BGETEQUAL, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), nRange1, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Error Getting Monster Record # " & nRange1 & ", exiting. - " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
End If

Do While nStatus = 0 And bStopProcess = False
    MonsterRowToStruct Monsterdatabuf.buf
    
    sName = Monsterrec.Number & "-" & ClipNull(Monsterrec.Name)
    
    If chkAll.Value = 0 Then
        If Monsterrec.Number > nRange2 Then Exit Do
    End If
    
    For nOnlyIfIndex = 0 To 1
        If chkOnlyIfOn(nOnlyIfIndex).Value = 1 Then
            Select Case cmbOnlyIf(nOnlyIfIndex).ListIndex
                Case 0: 'game limit
                    bSkip = TestOnlyIf(Monsterrec.GameLimit, nOnlyIfIndex)
                Case 1: 'experience
                    If eDatFileVersion >= v111j Then
                        bSkip = TestOnlyIf(CDbl(Monsterrec.Experience) * CDbl(Monsterrec.ExpMulti), nOnlyIfIndex)
                    Else
                        bSkip = TestOnlyIf(Monsterrec.Experience, nOnlyIfIndex)
                    End If
                Case 2: 'regen time
                    bSkip = TestOnlyIf(Monsterrec.RegenTime, nOnlyIfIndex)
                Case 3: 'group w/index
                    bSkip = TestOnlyIfAux(Monsterrec.Group, nOnlyIfIndex)
                    If bSkip = False Then GoTo Skip:
                    bSkip = TestOnlyIf(Monsterrec.Index, nOnlyIfIndex)
                Case 4: 'group
                    bSkip = TestOnlyIf(Monsterrec.Group, nOnlyIfIndex)
                Case 5: 'runic
                    bSkip = TestOnlyIf(Monsterrec.Runic, nOnlyIfIndex)
                Case 6: 'platinum
                    bSkip = TestOnlyIf(Monsterrec.Platinum, nOnlyIfIndex)
                Case 7: 'gold
                    bSkip = TestOnlyIf(Monsterrec.Gold, nOnlyIfIndex)
                Case 8: 'silver
                    bSkip = TestOnlyIf(Monsterrec.Silver, nOnlyIfIndex)
                Case 9: 'copper
                    bSkip = TestOnlyIf(Monsterrec.Copper, nOnlyIfIndex)
                Case 10: 'Charm LVL
                    bSkip = TestOnlyIf(Monsterrec.CharmLvL, nOnlyIfIndex)
                Case 11: 'Follow %
                    bSkip = TestOnlyIf(Monsterrec.Follow, nOnlyIfIndex)
                Case 12: 'MR
                    bSkip = TestOnlyIf(Monsterrec.MR, nOnlyIfIndex)
                Case 13: 'HP Regen
                    bSkip = TestOnlyIf(Monsterrec.HPRegen, nOnlyIfIndex)
                Case 14: 'Hit Points
                    bSkip = TestOnlyIf(Monsterrec.Hitpoints, nOnlyIfIndex)
                Case 15: 'AC"
                    bSkip = TestOnlyIf(Monsterrec.AC, nOnlyIfIndex)
                Case 16: 'DR
                    bSkip = TestOnlyIf(Monsterrec.DR, nOnlyIfIndex)
                Case 17, 18: 'Drop%, Item Uses
                    bSkip = True
                Case 19: 'has abil
                    For x = 0 To 9
                        bSkip = TestOnlyIfAux(Monsterrec.AbilityA(x), nOnlyIfIndex)
                        If bSkip = True Then
                            bSkip = TestOnlyIf(Monsterrec.AbilityB(x), nOnlyIfIndex)
                            If bSkip = True Then Exit For
                        End If
                    Next x
                Case 20: 'doesn't have abil
                    For x = 0 To 9
                        bSkip = TestOnlyIfAux(Monsterrec.AbilityA(x), nOnlyIfIndex)
                        If bSkip = True Then
                            bSkip = TestOnlyIf(Monsterrec.AbilityB(x), nOnlyIfIndex)
                            If bSkip = True Then 'true because if it matches then it 'HAS' it, so skip
                                bSkip = False
                                Exit For
                            End If
                        End If
                        bSkip = True
                    Next x
            End Select
            If bSkip = False Then
                If chkOnlyChanges.Value = 0 Then ts.WriteLine sName & " ==> Skipped due to 'Only If' Setting"
                GoTo Skip:
            End If
        End If
    Next nOnlyIfIndex
    
    StatusBar.Panels(2).Text = Monsterrec.Number
    Select Case nIndex
        Case 0:
            Monsterrec.Experience = DoMath(Monsterrec.Experience, 2, 2, sName)
            If eDatFileVersion >= v111j Then
                If CDbl(SLong2ULong(Monsterrec.Experience)) * CDbl(SLong2ULong(Monsterrec.ExpMulti)) > 2147483646 Then
                    Monsterrec.Experience = 65538
                    Monsterrec.ExpMulti = 32767
                    ts.WriteLine "^--> Changed to 2,147,483,646 (overflow)"
                End If
            End If
        Case 1:
            Monsterrec.ExpMulti = DoMath(Monsterrec.ExpMulti, 2, 2, sName)
            If CDbl(SLong2ULong(Monsterrec.Experience)) * CDbl(SLong2ULong(Monsterrec.ExpMulti)) > 2147483646 Then
                Monsterrec.Experience = 65538
                Monsterrec.ExpMulti = 32767
                ts.WriteLine "^--> Changed to 2,147,483,646 (overflow)"
            End If
        Case 2:
            nBase = CDbl(SLong2ULong(Monsterrec.Experience)) * CDbl(SLong2ULong(Monsterrec.ExpMulti))
            nBase = DoMath(nBase, 2, 2, sName)
            
            If nBase > 2147483646 Then
                nBase = 2147483646
                ts.WriteLine "^--> Changed to 2,147,483,646 (overflow)"
            End If
tryagain:
            If nBase > 100000 Then
                nMultiMax = 20
                For y = 20 To 32767
                    If y * 65538 >= nBase Then
                        nMultiMax = y
                        Exit For
                    End If
                Next y
                
                nMulti = 1
                For y = 3 To nMultiMax
                    temp = nBase Mod y
                    If temp = 0 Then nMulti = y
                Next y
                
                If nMulti = 1 Then
                    nBase = nBase - 1
                    GoTo tryagain:
                End If
                
                nBase = nBase / nMulti
            Else
                nMulti = 1
            End If
            
            If nBase <= 0 Then nBase = 1
                  
            Monsterrec.Experience = ULong2SLong(nBase)
            Monsterrec.ExpMulti = ULong2SLong(nMulti)

        Case 3: Monsterrec.MR = DoMath(Monsterrec.MR, dtInteger, suSigned, sName)
        Case 4: Monsterrec.CharmLvL = DoMath(Monsterrec.CharmLvL, dtInteger, suSigned, sName)
        Case 5: Monsterrec.AC = DoMath(Monsterrec.AC, dtInteger, suSigned, sName)
        Case 6: Monsterrec.DR = DoMath(Monsterrec.DR, dtInteger, suSigned, sName)
        Case 7: Monsterrec.Follow = DoMath(Monsterrec.Follow, dtInteger, suSigned, sName)
        Case 8: Monsterrec.RegenTime = DoMath(Monsterrec.RegenTime, dtInteger, suSigned, sName)
        Case 9: Monsterrec.GameLimit = DoMath(Monsterrec.GameLimit, dtInteger, suSigned, sName)
        Case 10: Monsterrec.Hitpoints = DoMath(Monsterrec.Hitpoints, dtInteger, suSigned, sName)
        Case 11: Monsterrec.HPRegen = DoMath(Monsterrec.HPRegen, dtInteger, suSigned, sName)
        Case 12: Monsterrec.Energy = DoMath(Monsterrec.Energy, dtInteger, suSigned, sName)
        Case 13: Monsterrec.Runic = DoMath(Monsterrec.Runic, dtLong, suUnsigned, sName)
        Case 14: Monsterrec.Platinum = DoMath(Monsterrec.Platinum, dtLong, suUnsigned, sName)
        Case 15: Monsterrec.Gold = DoMath(Monsterrec.Gold, dtLong, suUnsigned, sName)
        Case 16: Monsterrec.Silver = DoMath(Monsterrec.Silver, dtLong, suUnsigned, sName)
        Case 17: Monsterrec.Copper = DoMath(Monsterrec.Copper, dtLong, suUnsigned, sName)
        Case 18:
            Monsterrec.Runic = DoMath(Monsterrec.Runic, dtLong, suUnsigned, sName & "(runic)")
            Monsterrec.Platinum = DoMath(Monsterrec.Platinum, dtLong, suUnsigned, sName & "(platinum)")
            Monsterrec.Gold = DoMath(Monsterrec.Gold, dtLong, suUnsigned, sName & "(gold)")
            Monsterrec.Silver = DoMath(Monsterrec.Silver, dtLong, suUnsigned, sName & "(silver)")
            Monsterrec.Copper = DoMath(Monsterrec.Copper, dtLong, suUnsigned, sName & "(copper)")
        Case 19, 20: 'drop %, drop uses
        
            For x = 0 To 9
                For nOnlyIfIndex = 0 To 1
                    If chkOnlyIfOn(nOnlyIfIndex).Value = 1 Then 'per item checks
                        bSkip = True
                        Select Case cmbOnlyIf(nOnlyIfIndex).ListIndex
                            Case 17: 'percent
                                bSkip = TestOnlyIf(Monsterrec.ItemDropPer(x), nOnlyIfIndex)
                            Case 18: 'uses
                                bSkip = TestOnlyIf(Monsterrec.ItemUses(x), nOnlyIfIndex)
                        End Select
                        If bSkip = False Then
                            If chkOnlyChanges.Value = 0 Then ts.WriteLine sName & " (" & x & ")" _
                                & " ==> Skipped due to 'Only If' Setting"
                            GoTo SkipIndiv:
                        End If
                    End If
                Next nOnlyIfIndex
                
                Select Case nIndex
                    Case 19: Monsterrec.ItemDropPer(x) = DoMath(Monsterrec.ItemDropPer(x), dtByte, suSigned, sName, x)
                    Case 20: Monsterrec.ItemUses(x) = DoMath(Monsterrec.ItemUses(x), dtInteger, suSigned, sName, x)
                End Select
SkipIndiv:
            Next x
            
        Case 21:
            Monsterrec.Active = DoMath(Monsterrec.Active, dtInteger, suSigned, sName)
            
        Case 22: 'give
            For x = 0 To 9
                If Monsterrec.AbilityA(x) = cmbAbilities.ItemData(cmbAbilities.ListIndex) Then
                    If Monsterrec.AbilityB(x) = Val(txtValue.Text) Then
                        If chkOnlyChanges.Value = 0 Then ts.WriteLine sName & ": Already has ability"
                        GoTo Skip:
                    End If
                End If
            Next
            For x = 0 To 9
                If Monsterrec.AbilityA(x) = 0 Then
                    Monsterrec.AbilityA(x) = cmbAbilities.ItemData(cmbAbilities.ListIndex)
                    Monsterrec.AbilityB(x) = Val(txtValue.Text)
                    ts.WriteLine sName & ": Ability added to slot " & x
                    Exit For
                End If
            Next
            If chkOnlyChanges.Value = 0 Then If x = 10 Then ts.WriteLine sName & ": No Empty Slot to Add Ability"
        Case 23: 'take
            bSkip = False
            For x = 0 To 9
                If Monsterrec.AbilityA(x) = cmbAbilities.ItemData(cmbAbilities.ListIndex) Then
                    Monsterrec.AbilityA(x) = 0
                    ts.WriteLine sName & ": Ability taken from slot " & x
                    bSkip = True
                End If
            Next
            If chkOnlyChanges.Value = 0 And Not bSkip Then ts.WriteLine sName & ": Ability not found"
        Case 24: 'change
            For x = 0 To 9
                If Monsterrec.AbilityA(x) = cmbAbilities.ItemData(cmbAbilities.ListIndex) Then
                    Monsterrec.AbilityB(x) = DoMath(Monsterrec.AbilityB(x), dtInteger, suSigned, sName)
                End If
            Next
            If chkOnlyChanges.Value = 0 Then If x = 10 Then ts.WriteLine sName & ": Ability not found"
    End Select
    nStatus = UpdateMonster
    If Not nStatus = 0 Then MsgBox "Error Updating Monster Record # " & Monsterrec.Number & ", exiting. - " & BtrieveErrorCode(nStatus): Exit Sub
Skip:
    nStatus = BTRCALL(BGETNEXT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

Exit Sub

End Sub
Private Sub ModifyClass(ByVal nIndex As Integer, ByVal nRange1 As Long, ByVal nRange2 As Long)
Dim nStatus As Integer, x As Integer
Dim bSkip As Boolean, nOnlyIfIndex As Integer
Dim sName As String

StatusBar.Panels(1).Text = cmbEditor.Text & " - " & cmbField.Text

If chkAll.Value = 1 Then
    nStatus = BTRCALL(BGETFIRST, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Error Getting First Class Record, exiting. - " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
Else
    nStatus = BTRCALL(BGETEQUAL, ClassPosBlock, Classdatabuf, Len(Classdatabuf), nRange1, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Error Getting Class Record # " & nRange1 & ", exiting. - " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
End If

Do While nStatus = 0 And bStopProcess = False
    ClassRowToStruct Classdatabuf.buf
    
    sName = Classrec.Number & "-" & ClipNull(Classrec.Name)
    
    If chkAll.Value = 0 Then
        If Classrec.Number > nRange2 Then Exit Do
    End If
    
    For nOnlyIfIndex = 0 To 1
        If chkOnlyIfOn(nOnlyIfIndex).Value = 1 Then
            Select Case cmbOnlyIf(nOnlyIfIndex).ListIndex
                Case 0: 'combat
                    bSkip = TestOnlyIf(Classrec.Combat - 2, nOnlyIfIndex)
                Case 1: 'magery
                    bSkip = TestOnlyIfAux(Classrec.MagicType, nOnlyIfIndex)
                    If bSkip = False Then GoTo Skip:
                    bSkip = TestOnlyIf(Classrec.MagicLvL, nOnlyIfIndex)
                Case 2: 'exp %
                    bSkip = TestOnlyIf(Classrec.Exp + 100, nOnlyIfIndex)
                Case 3: 'hp min
                    bSkip = TestOnlyIf(Classrec.MinHp, nOnlyIfIndex)
                Case 4: 'hp max
                    bSkip = TestOnlyIf(Classrec.MinHp + Classrec.MaxHP, nOnlyIfIndex)
                Case 5: 'has abil
                    For x = 0 To 9
                        bSkip = TestOnlyIfAux(Classrec.AbilityA(x), nOnlyIfIndex)
                        If bSkip = True Then
                            bSkip = TestOnlyIf(Classrec.AbilityB(x), nOnlyIfIndex)
                            If bSkip = True Then Exit For
                        End If
                    Next x
                Case 6: 'doesn't have abil
                    For x = 0 To 9
                        bSkip = TestOnlyIfAux(Classrec.AbilityA(x), nOnlyIfIndex)
                        If bSkip = True Then
                            bSkip = TestOnlyIf(Classrec.AbilityB(x), nOnlyIfIndex)
                            If bSkip = True Then 'true because if it matches then it 'HAS' it, so skip
                                bSkip = False
                                Exit For
                            End If
                        End If
                        bSkip = True
                    Next x
            End Select
            If bSkip = False Then
                If chkOnlyChanges.Value = 0 Then ts.WriteLine sName & " ==> Skipped due to 'Only If' Setting"
                GoTo Skip:
            End If
        End If
    Next nOnlyIfIndex
    
    StatusBar.Panels(2).Text = Classrec.Number
    
    Select Case nIndex
        Case 0: Classrec.Exp = DoMath(Classrec.Exp + 100, 1, 1, sName) - 100
        Case 1: Classrec.MinHp = DoMath(Classrec.MinHp, 1, 1, sName)
        Case 2: Classrec.MaxHP = DoMath(Classrec.MinHp + Classrec.MaxHP, 1, 1, sName) - Classrec.MinHp
        Case 3: Classrec.Combat = DoMath(Classrec.Combat, dtInteger, suUnsigned, sName)
        Case 4: 'give
            For x = 0 To 9
                If Classrec.AbilityA(x) = cmbAbilities.ItemData(cmbAbilities.ListIndex) Then
                    If Classrec.AbilityB(x) = Val(txtValue.Text) Then
                        If chkOnlyChanges.Value = 0 Then ts.WriteLine sName & ": Already has ability"
                        GoTo Skip:
                    End If
                End If
            Next
            For x = 0 To 9
                If Classrec.AbilityA(x) = 0 Then
                    Classrec.AbilityA(x) = cmbAbilities.ItemData(cmbAbilities.ListIndex)
                    Classrec.AbilityB(x) = Val(txtValue.Text)
                    ts.WriteLine sName & ": Ability added to slot " & x
                    Exit For
                End If
            Next
            If chkOnlyChanges.Value = 0 Then If x = 10 Then ts.WriteLine sName & ": No Empty Slot to Add Ability"
        Case 5: 'take
            bSkip = False
            For x = 0 To 9
                If Classrec.AbilityA(x) = cmbAbilities.ItemData(cmbAbilities.ListIndex) Then
                    Classrec.AbilityA(x) = 0
                    ts.WriteLine sName & ": Ability taken from slot " & x
                    bSkip = True
                End If
            Next
            If chkOnlyChanges.Value = 0 And Not bSkip Then ts.WriteLine sName & ": Ability not found"
        Case 6: 'change
            For x = 0 To 9
                If Classrec.AbilityA(x) = cmbAbilities.ItemData(cmbAbilities.ListIndex) Then
                    Classrec.AbilityB(x) = DoMath(Classrec.AbilityB(x), dtInteger, suSigned, sName)
                    Exit For
                End If
            Next
            If chkOnlyChanges.Value = 0 Then If x = 10 Then ts.WriteLine sName & ": Ability not found"
    End Select
    nStatus = UpdateClass
    If Not nStatus = 0 Then MsgBox "Error Updating Class Record # " & Classrec.Number & ", exiting. - " & BtrieveErrorCode(nStatus): Exit Sub
Skip:
    nStatus = BTRCALL(BGETNEXT, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

Exit Sub

End Sub
Private Sub ModifyItem(ByVal nIndex As Integer, ByVal nRange1 As Long, ByVal nRange2 As Long)
Dim nStatus As Integer, x As Integer
Dim bSkip As Boolean, nOnlyIfIndex As Integer
Dim sName As String

StatusBar.Panels(1).Text = cmbEditor.Text & " - " & cmbField.Text

If chkAll.Value = 1 Then
    nStatus = BTRCALL(BGETFIRST, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Error Getting First Item Record, exiting. - " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
Else
    nStatus = BTRCALL(BGETEQUAL, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), nRange1, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Error Getting Item Record # " & nRange1 & ", exiting. - " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
End If

Do While nStatus = 0 And bStopProcess = False
    ItemRowToStruct Itemdatabuf.buf
    
    sName = Itemrec.Number & "-" & ClipNull(Itemrec.Name)
    
    If chkAll.Value = 0 Then
        If Itemrec.Number > nRange2 Then Exit Do
    End If
    
    For nOnlyIfIndex = 0 To 1
        If chkOnlyIfOn(nOnlyIfIndex).Value = 1 Then
            Select Case cmbOnlyIf(nOnlyIfIndex).ListIndex
                Case 0: 'game limit
                    bSkip = TestOnlyIf(Itemrec.GameLimit, nOnlyIfIndex)
                
                Case 1: 'item type
                    bSkip = TestOnlyIf(Itemrec.Type, nOnlyIfIndex)
                
                Case 2: 'armour type
                    bSkip = TestOnlyIfAux(Itemrec.Type, nOnlyIfIndex)
                    If bSkip = False Then GoTo Skip:
                    bSkip = TestOnlyIf(Itemrec.Armour, nOnlyIfIndex)
                    
                Case 3: 'weapon type
                    bSkip = TestOnlyIfAux(Itemrec.Type, nOnlyIfIndex)
                    If bSkip = False Then GoTo Skip:
                    bSkip = TestOnlyIf(Itemrec.Weapon, nOnlyIfIndex)
                    
                Case 4: 'worn on
                    bSkip = TestOnlyIf(Itemrec.WornOn, nOnlyIfIndex)
                
                Case 5: 'Weight
                    bSkip = TestOnlyIf(Itemrec.Weight, nOnlyIfIndex)
                    
                Case 6: 'Speed
                    bSkip = TestOnlyIf(Itemrec.Speed, nOnlyIfIndex)
                
                Case 7: 'Req. Strength
                    bSkip = TestOnlyIf(Itemrec.ReqStr, nOnlyIfIndex)
                    
                Case 8: 'Accuracy
                    bSkip = TestOnlyIf(Itemrec.Accuracy, nOnlyIfIndex)
                    
                Case 9: 'Cost
                    If cmbOnlyIfAuxValue(0).ListIndex < 5 Then '5 is any
                        bSkip = TestOnlyIfAux(Itemrec.CostType, nOnlyIfIndex)
                        If bSkip = False Then GoTo Skip:
                    End If
                    bSkip = TestOnlyIf(Itemrec.Cost, nOnlyIfIndex)
                
                Case 10: 'AC
                    bSkip = TestOnlyIf(Itemrec.AC, nOnlyIfIndex)
                    
                Case 11: 'DR
                    bSkip = TestOnlyIf(Itemrec.DR, nOnlyIfIndex)
                    
                Case 12: 'Min Hit
                    bSkip = TestOnlyIf(Itemrec.Minhit, nOnlyIfIndex)
                    
                Case 13: 'Max Hit
                    bSkip = TestOnlyIf(Itemrec.Maxhit, nOnlyIfIndex)
                Case 14: 'has abil
                    For x = 0 To 19
                        bSkip = TestOnlyIfAux(Itemrec.AbilityA(x), nOnlyIfIndex)
                        If bSkip = True Then
                            bSkip = TestOnlyIf(Itemrec.AbilityB(x), nOnlyIfIndex)
                            If bSkip = True Then Exit For
                        End If
                    Next x
                Case 15: 'doesn't have abil
                    For x = 0 To 19
                        bSkip = TestOnlyIfAux(Itemrec.AbilityA(x), nOnlyIfIndex)
                        If bSkip = True Then
                            bSkip = TestOnlyIf(Itemrec.AbilityB(x), nOnlyIfIndex)
                            If bSkip = True Then 'true because if it matches then it 'HAS' it, so skip
                                bSkip = False
                                Exit For
                            End If
                        End If
                        bSkip = True
                    Next x
            End Select
            If bSkip = False Then
                If chkOnlyChanges.Value = 0 Then ts.WriteLine sName & " ==> Skipped due to 'Only If' Setting"
                GoTo Skip:
            End If
        End If
    Next nOnlyIfIndex
    
    StatusBar.Panels(2).Text = Itemrec.Number
    Select Case nIndex
        Case 0: Itemrec.GameLimit = DoMath(Itemrec.GameLimit, 1, 1, sName)
        Case 1: Itemrec.Weight = DoMath(Itemrec.Weight, 1, 1, sName)
        Case 2: Itemrec.Minhit = DoMath(Itemrec.Minhit, 1, 1, sName)
        Case 3: Itemrec.Maxhit = DoMath(Itemrec.Maxhit, 1, 1, sName)
        Case 4: Itemrec.Speed = DoMath(Itemrec.Speed, 1, 1, sName)
        Case 5: Itemrec.ReqStr = DoMath(Itemrec.ReqStr, 1, 1, sName)
        Case 6: Itemrec.AC = DoMath(Itemrec.AC, 1, 1, sName)
        Case 7: Itemrec.DR = DoMath(Itemrec.DR, 1, 1, sName)
        Case 8: Itemrec.Accuracy = DoMath(Itemrec.Accuracy, 1, 1, sName)
        Case 9: Itemrec.Uses = DoMath(Itemrec.Uses, 1, 1, sName)
        Case 10: Itemrec.Cost = DoMath(Itemrec.Cost, dtInteger, suUnsigned, sName)
        Case 11: 'give
            For x = 0 To 19
                If Itemrec.AbilityA(x) = cmbAbilities.ItemData(cmbAbilities.ListIndex) Then
                    If Itemrec.AbilityB(x) = Val(txtValue.Text) Then
                        If chkOnlyChanges.Value = 0 Then ts.WriteLine sName & ": Already has ability"
                        GoTo Skip:
                    End If
                End If
            Next
            For x = 0 To 19
                If Itemrec.AbilityA(x) = 0 Then
                    Itemrec.AbilityA(x) = cmbAbilities.ItemData(cmbAbilities.ListIndex)
                    Itemrec.AbilityB(x) = Val(txtValue.Text)
                    ts.WriteLine sName & ": Ability added to slot " & x
                    Exit For
                End If
            Next
            If chkOnlyChanges.Value = 0 Then If x = 20 Then ts.WriteLine sName & ": No Empty Slot to Add Ability"
        Case 12: 'take
            bSkip = False
            For x = 0 To 19
                If Itemrec.AbilityA(x) = cmbAbilities.ItemData(cmbAbilities.ListIndex) Then
                    Itemrec.AbilityA(x) = 0
                    ts.WriteLine sName & ": Ability taken from slot " & x
                    bSkip = True
                End If
            Next
            If chkOnlyChanges.Value = 0 And Not bSkip Then ts.WriteLine sName & ": Ability not found"
        Case 13: 'change
            For x = 0 To 19
                If Itemrec.AbilityA(x) = cmbAbilities.ItemData(cmbAbilities.ListIndex) Then
                    Itemrec.AbilityB(x) = DoMath(Itemrec.AbilityB(x), dtInteger, suSigned, sName)
                End If
            Next
            If chkOnlyChanges.Value = 0 Then If x = 20 Then ts.WriteLine sName & ": Ability not found"
    End Select
    nStatus = UpdateItem
    If Not nStatus = 0 Then MsgBox "Error Updating Item Record # " & Itemrec.Number & ", exiting. - " & BtrieveErrorCode(nStatus): Exit Sub
Skip:
    nStatus = BTRCALL(BGETNEXT, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)

Call WriteINI("Settings", "UniLogOnlyChanges", chkOnlyChanges.Value)

If Me.WindowState = vbNormal Then
    Call WriteINI("Windows", "UniTop", Me.Top)
    Call WriteINI("Windows", "UniLeft", Me.Left)
End If

End Sub


Private Sub txtLimit_GotFocus()
Call SelectAll(txtLimit)

End Sub

Private Sub txtLimit_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtMap_GotFocus()
Call SelectAll(txtMap)

End Sub

Private Sub txtMap_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtOnlyIfValue_GotFocus(Index As Integer)
Call SelectAll(txtOnlyIfValue(Index))

End Sub

Private Sub txtR1_GotFocus()
Call SelectAll(txtR1)

End Sub

Private Sub txtR1_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtR2_GotFocus()
Call SelectAll(txtR2)

End Sub

Private Sub txtR2_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtValue_GotFocus()
Call SelectAll(txtValue)

End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub
Private Sub cmbValue_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub







