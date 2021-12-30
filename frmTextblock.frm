VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20D5284F-7B23-4F0A-B8B1-6C9D18B64F1C}#1.0#0"; "exlimiter.ocx"
Begin VB.Form frmTextblock 
   Caption         =   "Textblock Editor"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9780
   Icon            =   "frmTextblock.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   9780
   Begin VB.Frame fraEdit 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Visible         =   0   'False
      Width           =   9795
      Begin VB.CommandButton cmdLineLinkToBack 
         BackColor       =   &H00C0C0C0&
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   73
         ToolTipText     =   "Back to last Goto or LinkFrom"
         Top             =   60
         Width           =   435
      End
      Begin VB.CommandButton cmdLineMove 
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
         Height          =   375
         Index           =   1
         Left            =   7620
         TabIndex        =   60
         ToolTipText     =   "Next"
         Top             =   60
         Width           =   435
      End
      Begin VB.CommandButton cmdLineMove 
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
         Height          =   375
         Index           =   0
         Left            =   7200
         TabIndex        =   59
         ToolTipText     =   "Previous"
         Top             =   60
         Width           =   435
      End
      Begin VB.CommandButton cmdLineLinkTo 
         Caption         =   "LinksTo"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3180
         TabIndex        =   56
         Top             =   60
         Width           =   1575
      End
      Begin NightmareRedux.cntSplitter splSplitter 
         Height          =   5355
         Left            =   60
         TabIndex        =   71
         Top             =   480
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   9446
         Begin VB.Frame fraRight 
            BorderStyle     =   0  'None
            Height          =   5235
            Left            =   4440
            TabIndex        =   72
            Top             =   0
            Width           =   5115
            Begin VB.CommandButton cmdMoveLine 
               Height          =   555
               Index           =   3
               Left            =   120
               MaskColor       =   &H00FF00FF&
               Picture         =   "frmTextblock.frx":08CA
               Style           =   1  'Graphical
               TabIndex        =   65
               Top             =   840
               UseMaskColor    =   -1  'True
               Width           =   555
            End
            Begin VB.CommandButton cmdMoveLine 
               Height          =   555
               Index           =   2
               Left            =   120
               MaskColor       =   &H00FF00FF&
               Picture         =   "frmTextblock.frx":150C
               Style           =   1  'Graphical
               TabIndex        =   66
               Top             =   1500
               UseMaskColor    =   -1  'True
               Width           =   555
            End
            Begin VB.CommandButton cmdMoveLine 
               Height          =   555
               Index           =   1
               Left            =   120
               MaskColor       =   &H00FF00FF&
               Picture         =   "frmTextblock.frx":214E
               Style           =   1  'Graphical
               TabIndex        =   67
               Top             =   2280
               UseMaskColor    =   -1  'True
               Width           =   555
            End
            Begin VB.CommandButton cmdMoveLine 
               Height          =   555
               Index           =   0
               Left            =   120
               MaskColor       =   &H00FF00FF&
               Picture         =   "frmTextblock.frx":2D90
               Style           =   1  'Graphical
               TabIndex        =   64
               Top             =   60
               UseMaskColor    =   -1  'True
               Width           =   555
            End
            Begin VB.TextBox txtAnsiEdit 
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   9
                  Charset         =   255
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5115
               Left            =   780
               MultiLine       =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   68
               Top             =   0
               Width           =   4215
            End
         End
         Begin MSComctlLib.ListView lvLines 
            Height          =   5115
            Left            =   0
            TabIndex        =   63
            Top             =   0
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   9022
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   61
         Top             =   60
         Width           =   375
      End
      Begin VB.TextBox txtGotoInLine 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6360
         MaxLength       =   6
         TabIndex        =   58
         Top             =   120
         Width           =   675
      End
      Begin VB.CommandButton cmdGotoInLine 
         Caption         =   "&Goto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   57
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdEditCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8760
         TabIndex        =   62
         Top             =   60
         Width           =   975
      End
      Begin VB.CommandButton cmdEditSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   55
         Top             =   60
         Width           =   1035
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1260
         TabIndex        =   69
         Top             =   60
         Width           =   1755
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   5835
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9735
      Begin VB.CommandButton cmdDiscard 
         Caption         =   "Dis&card"
         Height          =   375
         Left            =   8940
         TabIndex        =   13
         Top             =   120
         Width           =   795
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7260
         TabIndex        =   11
         Top             =   120
         Width           =   795
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "&Last"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5700
         TabIndex        =   9
         Top             =   120
         Width           =   675
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   7
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Previous"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   6
         Top             =   120
         Width           =   855
      End
      Begin VB.Frame Frame2 
         Caption         =   "Insert Special Characters:"
         Height          =   915
         Left            =   7080
         TabIndex        =   48
         Top             =   540
         Width           =   2655
         Begin VB.CommandButton cmdAnsiResetCode 
            Caption         =   "&Ansi Reset"
            Height          =   435
            Left            =   1380
            TabIndex        =   50
            Top             =   300
            Width           =   1155
         End
         Begin VB.CommandButton cmdSeperate 
            Caption         =   "Line&break ( | )"
            Height          =   435
            Left            =   120
            TabIndex        =   49
            Top             =   300
            Width           =   1155
         End
      End
      Begin VB.TextBox txtAnsi 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   0
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   47
         Top             =   2460
         Width           =   9735
      End
      Begin VB.Frame Frame1 
         Caption         =   "Insert ANSI Color Code:"
         Height          =   915
         Left            =   60
         TabIndex        =   29
         Top             =   1500
         Width           =   9615
         Begin VB.CommandButton cmdColorQ 
            Caption         =   "?"
            Height          =   255
            Left            =   4320
            TabIndex        =   70
            Top             =   540
            Width           =   315
         End
         Begin VB.CommandButton AnsiBGreen 
            Caption         =   "Bright Green"
            Height          =   255
            Left            =   7260
            MaskColor       =   &H8000000F&
            TabIndex        =   41
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton AnsiMagenta 
            Caption         =   "Magenta"
            Height          =   255
            Left            =   7260
            MaskColor       =   &H8000000F&
            TabIndex        =   45
            Top             =   540
            Width           =   1095
         End
         Begin VB.CommandButton AnsiDGray 
            Caption         =   "Dark Gray"
            Height          =   255
            Left            =   8400
            MaskColor       =   &H8000000F&
            TabIndex        =   42
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton AnsiBCyan 
            Caption         =   "Bright Cyan"
            Height          =   255
            Left            =   6120
            MaskColor       =   &H8000000F&
            TabIndex        =   44
            Top             =   540
            Width           =   1095
         End
         Begin VB.CommandButton AnsiBBlue 
            Caption         =   "Bright Blue"
            Height          =   255
            Left            =   6120
            MaskColor       =   &H8000000F&
            TabIndex        =   40
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton AnsiBRed 
            Caption         =   "Bright Red"
            Height          =   255
            Left            =   4980
            MaskColor       =   &H8000000F&
            TabIndex        =   39
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton AnsiYellow 
            Caption         =   "Yellow"
            Height          =   255
            Left            =   8400
            MaskColor       =   &H8000000F&
            TabIndex        =   46
            Top             =   540
            Width           =   1095
         End
         Begin VB.CommandButton AnsiWhite 
            Caption         =   "White"
            Height          =   255
            Left            =   4980
            MaskColor       =   &H8000000F&
            TabIndex        =   43
            Top             =   540
            Width           =   1095
         End
         Begin VB.CommandButton AnsiBlack 
            Caption         =   "Black"
            Height          =   255
            Left            =   120
            MaskColor       =   &H8000000F&
            TabIndex        =   35
            Top             =   540
            Width           =   915
         End
         Begin VB.CommandButton AnsiPurple 
            Caption         =   "Purple"
            Height          =   255
            Left            =   2040
            MaskColor       =   &H8000000F&
            TabIndex        =   37
            Top             =   540
            Width           =   915
         End
         Begin VB.CommandButton AnsiCyan 
            Caption         =   "Cyan"
            Height          =   255
            Left            =   1080
            MaskColor       =   &H8000000F&
            TabIndex        =   36
            Top             =   540
            Width           =   915
         End
         Begin VB.CommandButton AnsiBlue 
            Caption         =   "Blue"
            Height          =   255
            Left            =   1080
            MaskColor       =   &H8000000F&
            TabIndex        =   32
            Top             =   240
            Width           =   915
         End
         Begin VB.CommandButton AnsiRed 
            Caption         =   "Red"
            Height          =   255
            Left            =   120
            MaskColor       =   &H8000000F&
            TabIndex        =   31
            Top             =   240
            Width           =   915
         End
         Begin VB.CommandButton AnsiBrown 
            Caption         =   "Brown"
            Height          =   255
            Left            =   3000
            MaskColor       =   &H8000000F&
            TabIndex        =   38
            Top             =   540
            Width           =   915
         End
         Begin VB.CommandButton AnsiGrey 
            Caption         =   "Gray"
            Height          =   255
            Left            =   3000
            TabIndex        =   34
            Top             =   240
            Width           =   915
         End
         Begin VB.CommandButton AnsiGreen 
            Caption         =   "Green"
            Height          =   255
            Left            =   2040
            MaskColor       =   &H8000000F&
            TabIndex        =   33
            Top             =   240
            Width           =   915
         End
         Begin VB.CheckBox chkBGColor 
            Caption         =   "Set BG"
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
            Left            =   4080
            TabIndex        =   30
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdGoto 
         Caption         =   "&Goto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   1
         Top             =   120
         Width           =   675
      End
      Begin VB.TextBox txtGotoBlock 
         Height          =   315
         Left            =   780
         TabIndex        =   0
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "Fi&rst"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   8
         Top             =   120
         Width           =   675
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "&Insert"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   10
         Top             =   120
         Width           =   795
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   8160
         TabIndex        =   12
         Top             =   120
         Width           =   795
      End
      Begin VB.TextBox txtGotoPart 
         Height          =   315
         Left            =   1620
         TabIndex        =   2
         Top             =   180
         Width           =   555
      End
      Begin VB.Frame Frame3 
         Caption         =   "Textblock Info:"
         Height          =   915
         Left            =   60
         TabIndex        =   14
         Top             =   540
         Width           =   3075
         Begin VB.TextBox txtPart 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   420
            Width           =   735
         End
         Begin VB.CommandButton cmdGotoLinko 
            Caption         =   "Go2"
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
            Left            =   2520
            TabIndex        =   21
            Top             =   435
            Width           =   435
         End
         Begin VB.TextBox txtNumber 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   420
            Width           =   735
         End
         Begin VB.TextBox txtLinkTo 
            Height          =   315
            Left            =   1740
            TabIndex        =   20
            Top             =   435
            Width           =   735
         End
         Begin VB.Label Label 
            Caption         =   "Part #:"
            Height          =   195
            Index           =   1
            Left            =   900
            TabIndex        =   17
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label 
            Caption         =   "Links to:"
            Height          =   195
            Index           =   2
            Left            =   1740
            TabIndex        =   19
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label 
            Caption         =   "Block #:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Misc:"
         Height          =   915
         Left            =   3180
         TabIndex        =   22
         Top             =   540
         Width           =   3795
         Begin VB.CommandButton cmdFind 
            Caption         =   "Find Ne&xt"
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
            Index           =   1
            Left            =   2760
            TabIndex        =   26
            Top             =   540
            Width           =   915
         End
         Begin VB.CommandButton cmdLineEdit 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Line Editor"
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
            Left            =   1380
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   240
            Width           =   1275
         End
         Begin VB.CommandButton cmdShowPreview 
            Caption         =   "S&how Preview"
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
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   1155
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
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
            Index           =   0
            Left            =   2760
            TabIndex        =   25
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lblCharCount 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1380
            TabIndex        =   28
            Top             =   600
            Width           =   1275
         End
         Begin VB.Label Label4 
            Caption         =   "Chars Left:"
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
            Index           =   0
            Left            =   180
            TabIndex        =   27
            Top             =   630
            Width           =   975
         End
      End
      Begin VB.ComboBox cmbRecent 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdClearPrevious 
         Height          =   135
         Left            =   3060
         TabIndex        =   5
         ToolTipText     =   "Clear Previous List"
         Top             =   0
         Width           =   135
      End
      Begin exlimiter.EL EL1 
         Left            =   8580
         Top             =   0
         _ExtentX        =   1270
         _ExtentY        =   1270
      End
      Begin VB.Label Label3 
         Caption         =   "Part#"
         Height          =   195
         Left            =   1620
         TabIndex        =   53
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Block#"
         Height          =   195
         Left            =   780
         TabIndex        =   52
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Previous"
         Height          =   195
         Left            =   2280
         TabIndex        =   51
         Top             =   0
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmTextblock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim nLastLineBlock As Long
Dim bCheckSave As Boolean
Dim bStopSearch As Boolean
Dim sFindText As String
Dim bLoaded As Boolean

Private Sub cmdLineLinkTo_Click()

nLastLineBlock = Val(txtNumber.Text)
Call GotoTB_InLine(Val(cmdLineLinkTo.Tag))

End Sub

Private Sub cmdLineLinkToBack_Click()
If nLastLineBlock < 0 Then Exit Sub
Call GotoTB_InLine(nLastLineBlock)
End Sub

Private Sub cmdLineMove_Click(Index As Integer)
Dim nStatus As Integer
On Error GoTo error:

If bCheckSave Then
    nStatus = MsgBox("Save current block first?", vbYesNoCancel + vbQuestion + vbDefaultButton1)
    If nStatus = vbYes Then
        Call cmdEditSave_Click
    ElseIf nStatus = vbCancel Then
        Exit Sub
    End If
End If

If Index = 0 Then '<<
    nStatus = BTRCALL(BGETPREVIOUS, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Next_Click, BGETNEXT, Textblock, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    Else
        DispTextblockInfo TextblockDataBuf.buf
    End If
Else '>>
again:
    nStatus = BTRCALL(BGETNEXT, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Next_Click, BGETNEXT, Textblock, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    Else
        Call TextblockRowToStruct(TextblockDataBuf.buf)
        If TextblockRec.PartNum > 0 Then GoTo again:
        DispTextblockInfo TextblockDataBuf.buf
    End If
End If

Call cmdLineEdit_Click

out:
Exit Sub
error:
Call HandleError("cmdLineMove_Click")
Resume out:

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim nStatus As Integer

With EL1
    .FormInQuestion = Me
    .EnableLimiter = True
    .MINHEIGHT = 255 + (TITLEBAR_OFFSET / 10)
    .MINWIDTH = 660
    .CenterOnLoad = False
End With

splSplitter.Orientation = cSPLTOrientationVertical
splSplitter.FullDrag = True
splSplitter.MinimumSize(cSPLTLeftOrTopPanel) = 40 'picSplitMain.ScaleX(picSplitLeft.Width, picSplitMain.ScaleMode, vbPixels)
splSplitter.MinimumSize(cSPLTRightOrBottomPanel) = 100
splSplitter.KeepProportion = True
splSplitter.SplitterSize = 6
splSplitter.Bind lvLines, fraRight
splSplitter.Position = 300

Me.Top = ReadINI("Windows", "TextTop")
Me.Left = ReadINI("Windows", "TextLeft")
Me.Width = ReadINI("Windows", "TextWidth")
Me.Height = ReadINI("Windows", "TextHeight")
'chkAutoLine.Value = ReadINI("Windows", "TextblockAutoUpdateLine")

lvLines.ColumnHeaders.clear
lvLines.ColumnHeaders.add , , "line"
lvLines.ColumnHeaders(1).Width = 30000

nStatus = BTRCALL(BGETFIRST, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "GETFIRST Textblock, Error: " & BtrieveErrorCode(nStatus)
    bLoaded = False
Else
    bLoaded = True
    DispTextblockInfo TextblockDataBuf.buf
End If

'Load frmPreview
Me.Show
Me.SetFocus
txtGotoBlock.SetFocus
If ReadINI("Windows", "TextMaxed") = "1" Then Me.WindowState = vbMaximized

End Sub

Private Sub cmbRecent_Click()
On Error GoTo error:
Dim nStatus As Integer, nNumber As Long, x As Integer, temp As String

If bLoaded Then SaveTextBlock

If Me.Visible = False Then Exit Sub
'Me.Show
'Me.SetFocus

If cmbRecent.ItemData(cmbRecent.ListIndex) = Val(txtNumber.Text) Then Exit Sub

TextblockKey.PartNum = 0
TextblockKey.Number = cmbRecent.ItemData(cmbRecent.ListIndex)

nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKey, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    If nStatus = 4 Then
        MsgBox "Record not found."
    Else
        MsgBox "Goto Error: " & BtrieveErrorCode(nStatus)
    End If
Else
    DispTextblockInfo TextblockDataBuf.buf
    nNumber = cmbRecent.ItemData(cmbRecent.ListIndex)
    cmbRecent.RemoveItem (cmbRecent.ListIndex)
    cmbRecent.AddItem nNumber, 0
    cmbRecent.ItemData(0) = nNumber
    cmbRecent.ListIndex = 0
End If

Exit Sub
error:
Call HandleError
End Sub

Private Sub cmdClearPrevious_Click()
cmbRecent.clear
End Sub

Private Sub cmdColorQ_Click()
MsgBox "Use the Set BG checkbox to set the background color.  Only the left set of boxes" _
    & " can be used as the background color.", vbInformation
End Sub

Private Sub cmdEditCancel_Click()
fraMain.Visible = True
fraEdit.Visible = False
End Sub


Private Sub cmdEditSave_Click()
Dim x As Long, sData As String, nPart As Integer, nStatus As Integer
Dim nPos As Long
On Error GoTo error:

'Call cmdUpdateLine_Click

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

If lvLines.ListItems.Count < 1 Then
    Call cmdEditCancel_Click
    Exit Sub
End If

For x = 1 To lvLines.ListItems.Count
    sData = sData & lvLines.ListItems(x) & Chr(10)
Next x

nPart = 0
TextblockKey.Number = Val(txtNumber.Text)
TextblockKey.PartNum = nPart

'get current record
nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKey, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error getting current record." & BtrieveErrorCode(nStatus)
    bLoaded = False
    Exit Sub
End If
Call TextblockRowToStruct(TextblockDataBuf.buf)

Do While Len(sData) > 2000
    nPos = 1
    x = InStr(nPos, sData, Chr(10))
    Do While x > 0
        If x > 2000 Then Exit Do
        nPos = x
        x = InStr(nPos + 1, sData, Chr(10))
    Loop
    
    'if couldn't split at a linebreak...
    If nPos = 1 Then
        x = InStr(nPos, sData, ":")
        Do While x > 0
            If x > 2000 Then Exit Do
            nPos = x
            x = InStr(nPos + 1, sData, ":")
        Loop
    End If
    
    'if couldn't split at a command break...
    If nPos = 1 Then
        x = InStr(nPos, sData, " ")
        Do While x > 0
            If x > 2000 Then Exit Do
            nPos = x
            x = InStr(nPos + 1, sData, " ")
        Loop
    End If
    
    'if couldn't split at a space...
    If nPos = 1 Then nPos = 1990
    
    'update and save block
    TextblockRec.Data = EncryptTextblock(Left(sData, nPos))
    Call TextblockStructToRow(TextblockDataBuf.buf)
    nStatus = UpdateTextblock
    If Not nStatus = 0 Then
        MsgBox "Error updating textblock: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
    
    nPart = nPart + 1
    
    'get next block
    nStatus = BTRCALL(BGETNEXT, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 And Not nStatus = 9 Then
        MsgBox "Error getting textblock: " & BtrieveErrorCode(nStatus)
        Exit Sub
    ElseIf nStatus = 9 Then
        TextblockRec.PartNum = nPart
        TextblockRec.Number = TextblockKey.Number
        TextblockRec.LinkTo = 0
        TextblockRec.Data = String(2000, Chr(0))
        Call TextblockStructToRow(TextblockDataBuf.buf)
        nStatus = BTRCALL(BINSERT, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
        If Not nStatus = 0 And Not nStatus = 9 Then
            MsgBox "Error inserting textblock part: " & BtrieveErrorCode(nStatus)
            Exit Sub
        End If
    Else
        Call TextblockRowToStruct(TextblockDataBuf.buf)
        If Not TextblockRec.Number = TextblockKey.Number Then
            TextblockRec.PartNum = nPart
            TextblockRec.Number = TextblockKey.Number
            TextblockRec.LinkTo = 0
            TextblockRec.Data = String(2000, Chr(0))
            Call TextblockStructToRow(TextblockDataBuf.buf)
            nStatus = BTRCALL(BINSERT, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
            If Not nStatus = 0 And Not nStatus = 9 Then
                MsgBox "Error inserting textblock part: " & BtrieveErrorCode(nStatus)
                Exit Sub
            End If
        End If
    End If
    
    sData = Mid(sData, nPos + 1)
Loop

'update and save block
TextblockRec.Data = EncryptTextblock(sData)
Call TextblockStructToRow(TextblockDataBuf.buf)
nStatus = UpdateTextblock
If Not nStatus = 0 Then
    MsgBox "Error updating textblock: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

check_part_again:
'check for extra parts
nStatus = BTRCALL(BGETNEXT, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error getting textblock: " & BtrieveErrorCode(nStatus)
    Exit Sub
ElseIf nStatus = 0 Then
    Call TextblockRowToStruct(TextblockDataBuf.buf)
    If TextblockRec.Number = TextblockKey.Number Then
        nStatus = BTRCALL(BDELETE, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
        If Not nStatus = 0 And Not nStatus = 9 Then
            MsgBox "Error deleting extra textblock part: " & BtrieveErrorCode(nStatus)
            Exit Sub
        End If
        
        'TextblockKey.Number = Val(txtNumber.Text)
        TextblockKey.PartNum = nPart
        
        'get current record
        nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKey, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            MsgBox "Error re-getting last part." & BtrieveErrorCode(nStatus)
            Exit Sub
        End If
        GoTo check_part_again:
    End If
End If

'TextblockKey.Number = Val(txtNumber.Text)
TextblockKey.PartNum = 0

'get current record
nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKey, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error re-getting current record." & BtrieveErrorCode(nStatus)
    bLoaded = False
Else
    Call DispTextblockInfo(TextblockDataBuf.buf)
End If

fraMain.Visible = True
fraEdit.Visible = False

Exit Sub
error:
Call HandleError("cmdEditSave_Click")
End Sub

Private Sub cmdGotoInLine_Click()

If Val(txtGotoInLine.Text) < 0 Then Exit Sub
nLastLineBlock = Val(txtNumber.Text)
Call GotoTB_InLine(Val(txtGotoInLine.Text))

End Sub

Public Sub GotoTB_InLine(ByVal nRecnum As Long)
On Error GoTo error:
Dim nStatus As Integer, nNumber As Long, x As Integer, temp As String

If bCheckSave Then
    nStatus = MsgBox("Save current block first?", vbYesNoCancel + vbQuestion + vbDefaultButton1)
    If nStatus = vbYes Then
        Call cmdEditSave_Click
    ElseIf nStatus = vbCancel Then
        Exit Sub
    End If
End If

TextblockKey.PartNum = 0
TextblockKey.Number = nRecnum

nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKey, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Goto Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
Else
    Call AddRecent(nRecnum)
    DispTextblockInfo TextblockDataBuf.buf
End If

Call cmdLineEdit_Click

On Error Resume Next
Me.SetFocus
Exit Sub
error:
Call HandleError
End Sub


Private Sub cmdHelp_Click()
MsgBox "The box on the left lists all the textblock lines.  The controls to the" & vbCrLf _
    & "right of it manipulate the lines by moving them up and down (script order)" & vbCrLf _
    & "and adding/deleting lines.  The box to the right allows you to edit the" & vbCrLf _
    & "the script commands for the selected line.  You do not include the ':'" & vbCrLf _
    & "command separator.  Simply put each command on its own line and the line" & vbCrLf _
    & "editor will automaticly put in the separators.  When you save, it automaticly" & vbCrLf _
    & "parses the lines into parts (and deletes any extra parts if present).", vbInformation
    
End Sub

Private Sub cmdLineEdit_Click()
On Error GoTo error:
Dim sLine As String, x As Long, nPos As Long, sDecrypted As String
Dim nStatus As Integer, oLI As ListItem

If bLoaded = True Then Call SaveTextBlock

lvLines.ListItems.clear

TextblockKey.Number = Val(txtNumber.Text)
TextblockKey.PartNum = Val(txtPart.Text)

'get current record
nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKey, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error getting current record." & BtrieveErrorCode(nStatus)
    bLoaded = False
    Exit Sub
End If
Call TextblockRowToStruct(TextblockDataBuf.buf)

'find first part
prev_part:
If TextblockRec.PartNum > 0 Then
    nStatus = BTRCALL(BGETPREVIOUS, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKey, KEY_BUF_LEN, 0)
    If nStatus = 0 Then
        Call TextblockRowToStruct(TextblockDataBuf.buf)
        If TextblockRec.PartNum > 0 Then GoTo prev_part:
        If Not TextblockRec.Number = Val(txtNumber.Text) Then
            MsgBox "Error getting part 0 (no part 0 found)", vbExclamation
            bLoaded = False
            Exit Sub
        End If
        Call DispTextblockInfo(TextblockDataBuf.buf)
    Else
        MsgBox "Error getting part 0: " & BtrieveErrorCode(nStatus), vbExclamation
        bLoaded = False
        Exit Sub
    End If
End If

'collect data
sDecrypted = DecryptTextblock(TextblockRec.Data)

next_part:
nStatus = BTRCALL(BGETNEXT, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKey, KEY_BUF_LEN, 0)
If nStatus = 0 Then
    Call TextblockRowToStruct(TextblockDataBuf.buf)
    If TextblockRec.Number = Val(txtNumber.Text) Then
        sDecrypted = sDecrypted & DecryptTextblock(TextblockRec.Data)
        GoTo next_part:
    End If
End If

'reget part 0
TextblockKey.Number = Val(txtNumber.Text)
TextblockKey.PartNum = Val(txtPart.Text)

nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKey, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error getting current record." & BtrieveErrorCode(nStatus), vbExclamation
    bLoaded = False
    Exit Sub
End If
Call TextblockRowToStruct(TextblockDataBuf.buf)

If Len(sDecrypted) = 0 Then
    MsgBox "No Lines detected.", vbInformation
    Exit Sub
End If

If TextblockRec.LinkTo > 0 Then
    cmdLineLinkTo.Tag = TextblockRec.LinkTo
    cmdLineLinkTo.Caption = "LinksTo " & TextblockRec.LinkTo
    cmdLineLinkTo.Enabled = True
Else
    cmdLineLinkTo.Tag = 0
    cmdLineLinkTo.Caption = "No LinkTo"
    cmdLineLinkTo.Enabled = False
End If

'create list
x = 1
nPos = 1
Do While InStr(nPos, sDecrypted, Chr(10)) > 0
    x = InStr(nPos, sDecrypted, Chr(10))
    If nPos = x Then
        Set oLI = lvLines.ListItems.add(, , "")
        nPos = x + 1
        GoTo NextLine:
    End If
    'Debug.Print Mid(sDecrypted, nPos, x - nPos)
    sLine = Mid(sDecrypted, nPos, x - nPos)
    If Right(sLine, 1) = Chr(13) Then sLine = Mid(sLine, 1, Len(sLine) - 1)
    Set oLI = lvLines.ListItems.add(, , sLine)
    nPos = x + 1
NextLine:
Loop

If nPos <= Len(sDecrypted) Then
    sLine = Mid(sDecrypted, nPos)
    If Right(sLine, 1) = Chr(13) Then sLine = Mid(sLine, 1, Len(sLine) - 1)
    Set oLI = lvLines.ListItems.add(, , sLine)
End If

If lvLines.ListItems.Count > 0 Then
    Call lvLines_ItemClick(lvLines.ListItems(1))
End If

Call CalcChars

fraEdit.Visible = True
fraMain.Visible = False

out:
bCheckSave = False
Set oLI = Nothing
Exit Sub
error:
Call HandleError("cmdLineEdit_Click")
Resume out:
End Sub

Private Sub cmdMoveLine_Click(Index As Integer)
On Error GoTo error:
Dim nTemp As Integer, sTemp As String, oLI As ListItem

bCheckSave = True

Select Case Index
    Case 0: 'move up
        If lvLines.ListItems.Count < 1 Then Exit Sub
        If lvLines.SelectedItem Is Nothing Then Exit Sub
        If lvLines.SelectedItem.Index > 1 Then
            
            nTemp = lvLines.SelectedItem.Index - 1
            sTemp = lvLines.SelectedItem.Text
            lvLines.ListItems.Remove (lvLines.SelectedItem.Index)
            Set oLI = lvLines.ListItems.add(nTemp, , sTemp)
            
            For Each oLI In lvLines.ListItems
                oLI.Selected = False
            Next oLI
            
            Set lvLines.SelectedItem = lvLines.ListItems(nTemp)
            lvLines.SelectedItem.EnsureVisible
            Call lvLines_ItemClick(lvLines.SelectedItem)
            Set oLI = Nothing
            
        End If
    Case 1: 'move down
        If lvLines.ListItems.Count < 1 Then Exit Sub
        If lvLines.SelectedItem Is Nothing Then Exit Sub
        If lvLines.SelectedItem.Index < lvLines.ListItems.Count Then
        
            nTemp = lvLines.SelectedItem.Index + 1
            sTemp = lvLines.SelectedItem.Text
            lvLines.ListItems.Remove (lvLines.SelectedItem.Index)
            Set oLI = lvLines.ListItems.add(nTemp, , sTemp)
            
            For Each oLI In lvLines.ListItems
                oLI.Selected = False
            Next oLI
            
            Set lvLines.SelectedItem = lvLines.ListItems(nTemp)
            lvLines.SelectedItem.EnsureVisible
            Call lvLines_ItemClick(lvLines.SelectedItem)
            Set oLI = Nothing
            
        End If
    Case 2: 'delete
        If lvLines.ListItems.Count < 1 Then Exit Sub
        If lvLines.SelectedItem Is Nothing Then Exit Sub
        'nTemp = MsgBox("Are you sure you want to delete this line?", vbYesNo + vbQuestion + vbDefaultButton2)
        'If nTemp = vbYes Then
            nTemp = lvLines.SelectedItem.Index
            lvLines.ListItems.Remove (lvLines.SelectedItem.Index)
            
            If lvLines.ListItems.Count > 0 Then
                If nTemp > lvLines.ListItems.Count Then
                    If nTemp - 1 > 1 Then
                        nTemp = nTemp - 1
                    Else
                        nTemp = 1
                    End If
                End If
                Set lvLines.SelectedItem = lvLines.ListItems(nTemp)
                lvLines.SelectedItem.EnsureVisible
                Call lvLines_ItemClick(lvLines.SelectedItem)
            End If
        'End If
        
        Call CalcChars
    Case 3: 'add
        If lvLines.SelectedItem Is Nothing Then
            nTemp = 1
            sTemp = ""
        Else
            nTemp = lvLines.SelectedItem.Index + 1
            sTemp = lvLines.SelectedItem.Text
        End If
        For Each oLI In lvLines.ListItems
            oLI.Selected = False
        Next oLI
        
        Set oLI = lvLines.ListItems.add(nTemp, , sTemp)
        Set lvLines.SelectedItem = lvLines.ListItems(nTemp)
        lvLines.SelectedItem.EnsureVisible
        Call lvLines_ItemClick(lvLines.SelectedItem)
        Set oLI = Nothing
        
        Call CalcChars
End Select

Set oLI = Nothing
Exit Sub
error:
Call HandleError("cmdMoveLine_Click")
Set oLI = Nothing
End Sub


Private Sub CalcChars()
Dim oLI As ListItem, nChars As Long, nLines As Integer
On Error GoTo error:

For Each oLI In lvLines.ListItems
    nLines = nLines + 1
    nChars = nChars + Len(oLI.Text) + 1
Next oLI

lblInfo.Caption = nLines & " lines, " & nChars & " chars" & vbCrLf & "~" & Fix(nChars / 2000) + 1 & " part(s)"

Exit Sub
error:
Call HandleError("CalcChars")
End Sub

Private Sub AddRecent(ByVal nNum As Long) ', ByVal sText As String)
Dim x As Integer, nList(20) As Long

'If Len(sText) > 50 Then sText = Left(sText, 47) & "..."

For x = 0 To cmbRecent.ListCount - 1
    nList(x) = cmbRecent.ItemData(x)
Next
    
For x = 0 To cmbRecent.ListCount - 1
    If nNum = nList(x) Then
        cmbRecent.RemoveItem (x)
        cmbRecent.AddItem nNum, 0 '& " - " & sText, 0
        cmbRecent.ItemData(0) = nNum
        GoTo done:
    End If
Next x

If cmbRecent.ListCount > 20 Then
    For x = 0 To 19
        cmbRecent.List(x + 1) = nList(x)
        cmbRecent.ItemData(x + 1) = nList(x)
    Next
    
    cmbRecent.List(0) = nNum '& " - " & sText
    cmbRecent.ItemData(0) = nNum
Else
    cmbRecent.AddItem nNum, 0 '& " - " & sText, 0
    cmbRecent.ItemData(0) = nNum
End If

done:
'If cmbRecent.ListCount > 0 Then Call AutoSizeDropDownWidth(cmbRecent)

End Sub
Public Sub StopSearch()
bStopSearch = True
End Sub
Private Sub cmdFind_Click(Index As Integer)
On Error GoTo error:
Dim nStatus As Integer, x As Integer, decrypted As String, sTemp As String

sTemp = sFindText

bStopSearch = False
If Index = 0 Then 'find
    sFindText = InputBox("Enter String to search for", "Enter search string", sFindText)
    If sFindText = "" Then
        sFindText = sTemp
        Exit Sub
    End If
    
    nStatus = BTRCALL(BGETFIRST, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "GETFIRST Textblock, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
Else 'find next
    If sFindText = "" Then Exit Sub
    
    TextblockKey.PartNum = Val(txtPart.Text)
    TextblockKey.Number = Val(txtNumber.Text)
    nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKey, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Error getting current record." & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
    
    nStatus = BTRCALL(BGETNEXT, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nStatus = 9 Then
            MsgBox "You are at the last record."
            Exit Sub
        Else
            MsgBox "Couldn't get next record -- Error: " & BtrieveErrorCode(nStatus)
            Exit Sub
        End If
    End If
End If

frmProgressBar.sCaption = "TextBlock Search"
frmProgressBar.lblCaption.Caption = "Searching ..."
frmProgressBar.cmdCancel.Enabled = True
frmProgressBar.ProgressBar.Value = 0
Set frmProgressBar.FormOwner = Me

nStatus = BTRCALL(BSTAT, TextblockPosBlock, DBStatDatabuf, Len(TextblockDataBuf), 0, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    Call frmProgressBar.SetRange(8000)
    frmProgressBar.ProgressBar.Value = 1
Else
    DBStatRowToStruct DBStatDatabuf.buf
    Call frmProgressBar.SetRange(DBStat.nRecords)
End If

frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & "text2.dat"
frmProgressBar.lblPanel(1).Caption = TextblockRec.Number
frmProgressBar.Show
frmMain.Enabled = False
DoEvents


Do While nStatus = 0
    If bStopSearch Then GoTo canceled:
    
    TextblockRowToStruct TextblockDataBuf.buf
    frmProgressBar.IncreaseProgress
    frmProgressBar.lblPanel(1).Caption = TextblockRec.Number
    
    decrypted = DecryptTextblock(TextblockRec.Data)
    
    If Not InStr(1, LCase(decrypted), LCase(sFindText)) = 0 Then
        x = InStr(1, LCase(decrypted), LCase(sFindText))
        GoTo found:
    End If
    
    If Not TextblockRec.LinkTo = 0 And sFindText = CStr(TextblockRec.LinkTo) Then
        x = Len(decrypted) + 1
        GoTo found:
    End If
    
    nStatus = BTRCALL(BGETNEXT, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
    DoEvents
Loop

GoTo notfound:

found:
'MsgBox "Found.", vbInformation
Unload frmProgressBar
DoEvents
Call AddRecent(TextblockRec.Number)
frmMain.Enabled = True
frmMain.SetFocus
Call DispTextblockInfo(TextblockDataBuf.buf)
txtAnsi.SelStart = x - 1
txtAnsi.SelLength = Len(sFindText)
txtAnsi.SetFocus
Exit Sub

notfound:
MsgBox "String not found.", vbInformation
canceled:
Unload frmProgressBar
DoEvents
frmMain.Enabled = True
frmMain.SetFocus
TextblockKey.PartNum = Val(txtPart.Text)
TextblockKey.Number = Val(txtNumber.Text)
nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKey, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error getting current record." & BtrieveErrorCode(nStatus)
End If

Exit Sub
error:
Call HandleError
Unload frmProgressBar
DoEvents
frmMain.Enabled = True

End Sub

Private Sub cmdGotoLinko_Click()
On Error GoTo error:
Dim nStatus As Integer, nNumber As Long, x As Integer, temp As String

If bLoaded Then Call SaveTextBlock

TextblockKey.PartNum = 0
TextblockKey.Number = Val(txtLinkTo.Text)

nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKey, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Goto Error: " & BtrieveErrorCode(nStatus)
Else
    Call AddRecent(Val(txtLinkTo.Text))
    DispTextblockInfo TextblockDataBuf.buf
End If

Exit Sub
error:
Call HandleError
End Sub

Private Sub DispTextblockInfo(row() As Byte)
On Error GoTo error:
Dim x As Integer
bLoaded = True

RowToStruct row, TextblockFldMap, TextblockRec, LenB(TextblockRec)

'For x = 1 To 14
'    MsgBox TextblockRec.LeadIn(x)
'Next

Me.Caption = "Textblock Editor -- " & TextblockRec.Number & "/" & TextblockRec.PartNum

txtPart.Text = TextblockRec.PartNum
txtNumber.Text = TextblockRec.Number
txtLinkTo.Text = TextblockRec.LinkTo
txtAnsi.Text = DecryptTextblock(TextblockRec.Data)

If PreviewLoaded = True Then
    Call DispPreview
    Me.SetFocus
End If

lblCharCount.Caption = 2000 - Len(txtAnsi.Text)

bCheckSave = False

Exit Sub
error:
Call HandleError
MsgBox "Warning, record was not completely displayed." & vbCrLf _
    & "Previous records stats may still be in memory.  Select 'Disable DB Writing'" & vbCrLf _
    & "from the file menu and then reload the editor.", vbExclamation
End Sub

Private Sub SaveTextBlock()
On Error GoTo error:
Dim nStatus As Integer

If bDisableWriting Then Exit Sub

TextblockKey.Number = Val(txtNumber.Text)
TextblockKey.PartNum = Val(txtPart.Text)

nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKey, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error getting current record." & BtrieveErrorCode(nStatus)
    Exit Sub
End If

TextblockRowToStruct TextblockDataBuf.buf

DoEvents
TextblockRec.LinkTo = Val(txtLinkTo.Text)
TextblockRec.Data = EncryptTextblock(txtAnsi.Text)

nStatus = UpdateTextblock
If Not nStatus = 0 Then
    MsgBox "Save Error: " & BtrieveErrorCode(nStatus)
Else
    DispTextblockInfo TextblockDataBuf.buf
End If

Exit Sub
error:
Call HandleError
End Sub

Private Sub cmdFirst_Click()
Dim nStatus As Integer
If bLoaded Then Call SaveTextBlock

nStatus = BTRCALL(BGETFIRST, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "GETFIRST Textblock, Error: " & BtrieveErrorCode(nStatus)
Else
    DispTextblockInfo TextblockDataBuf.buf
End If
End Sub
Public Sub GotoTB(ByVal nTB As Long, Optional ByVal nPart As Long)
On Error GoTo error:
Dim nStatus As Integer, nNumber As Long, x As Integer, temp As String

If fraEdit.Visible = True Then
    Call GotoTB_InLine(nTB)
    Exit Sub
End If

If bLoaded Then SaveTextBlock

Me.Show
Me.SetFocus

TextblockKey.PartNum = nPart
TextblockKey.Number = nTB

nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKey, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    If Not nStatus = 4 Then MsgBox "Goto Error: " & BtrieveErrorCode(nStatus)
Else
    Call AddRecent(nTB)
    DispTextblockInfo TextblockDataBuf.buf
End If

Exit Sub
error:
Call HandleError

End Sub
Private Sub cmdGoto_Click()
On Error GoTo error:
Dim nStatus As Integer, nNumber As Long, x As Integer, temp As String

If bLoaded Then Call SaveTextBlock

TextblockKey.PartNum = Val(txtGotoPart.Text)
TextblockKey.Number = Val(txtGotoBlock.Text)

nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKey, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Goto Error: " & BtrieveErrorCode(nStatus)
Else
    Call AddRecent(Val(txtGotoBlock.Text))
    DispTextblockInfo TextblockDataBuf.buf
End If

Exit Sub
error:
Call HandleError
End Sub


Private Sub cmdLast_Click()
Dim nStatus As Integer
If bLoaded Then Call SaveTextBlock

nStatus = BTRCALL(BGETLAST, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "GETLAST Textblock, Error: " & BtrieveErrorCode(nStatus)
Else
    DispTextblockInfo TextblockDataBuf.buf
End If
End Sub

Private Sub cmdNext_Click()
Dim nStatus As Integer
If bLoaded Then Call SaveTextBlock

nStatus = BTRCALL(BGETNEXT, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Next_Click, BGETNEXT, Textblock, Error: " & BtrieveErrorCode(nStatus)
Else
    DispTextblockInfo TextblockDataBuf.buf
End If
End Sub

Private Sub cmdPrevious_Click()
Dim nStatus As Integer
If bLoaded Then Call SaveTextBlock

nStatus = BTRCALL(BGETPREVIOUS, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdNext_Click, BGETPREVIOUS, Textblock, Error: " & BtrieveErrorCode(nStatus)
Else
    DispTextblockInfo TextblockDataBuf.buf
End If
End Sub

Private Sub cmdDelete_Click()
Dim nStatus As Integer
Dim nDelete As Integer

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

nDelete = MsgBox("Delete this record from database?", vbYesNo, "Delete Record?")
If nDelete = vbNo Then Exit Sub

TextblockKey.Number = Val(txtNumber.Text)
TextblockKey.PartNum = Val(txtPart.Text)

nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKey, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error getting current record." & BtrieveErrorCode(nStatus)
    Exit Sub
End If

nStatus = BTRCALL(BDELETE, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Delete Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
Else
    Call Form_Load
End If

End Sub

Private Sub cmdInsert_Click()
On Error GoTo error:
Dim nStatus As Integer, NewTextBlockNumber As String, x As Integer
Dim NewTextBlockPart As String

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
If bLoaded = True Then SaveTextBlock

NewTextBlockNumber = InputBox("New TextBlock Number:", "Insert", Val(txtNumber.Text) + 1)
If NewTextBlockNumber = "" Then Exit Sub

NewTextBlockPart = InputBox("New TextBlock Part Number" & vbCrLf & vbCrLf & "(First part is actually part 0)", "Insert", 0)
If NewTextBlockPart = "" Then Exit Sub

TextblockKey.Number = Val(txtNumber.Text)
TextblockKey.PartNum = Val(txtPart.Text)

nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKey, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    For x = 1 To 14
        TextblockRec.LeadIn(x) = TextblockKey.LeadIn(x)
    Next
    
    TextblockRec.Data = String(Len(TextblockRec.Data), Chr(0))
End If

TextblockRec.PartNum = Val(NewTextBlockPart)
TextblockRec.Number = Val(NewTextBlockNumber)
TextblockRec.LinkTo = 0
    
TextblockStructToRow TextblockDataBuf.buf

nStatus = BTRCALL(BINSERT, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Insert, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    Else
        DispTextblockInfo TextblockDataBuf.buf
    End If
Exit Sub
error:
Call HandleError
End Sub


Private Sub cmdSeperate_Click()
InsertAnsiCode (Chr(10))
End Sub
Private Sub cmdAnsiResetCode_Click()
InsertAnsiCode ("[0m")
End Sub

Private Sub AnsiCyan_Click()
If chkBGColor.Value = 1 Then
    InsertAnsiCode ("[0;46m")
Else
    InsertAnsiCode ("[0;36m")
End If
End Sub

Private Sub AnsiDGray_Click()
InsertAnsiCode ("[1;30m")
End Sub

Private Sub AnsiGreen_Click()
If chkBGColor.Value = 1 Then
    InsertAnsiCode ("[0;42m")
Else
    InsertAnsiCode ("[0;32m")
End If
End Sub

Private Sub AnsiGrey_Click()
If chkBGColor.Value = 1 Then
    InsertAnsiCode ("[0;47m")
Else
    InsertAnsiCode ("[0;37m")
End If
End Sub

Private Sub AnsiMagenta_Click()
    InsertAnsiCode ("[1;35m")
End Sub

Private Sub AnsiPurple_Click()
If chkBGColor.Value = 1 Then
    InsertAnsiCode ("[0;45m")
Else
    InsertAnsiCode ("[0;35m")
End If
End Sub

Private Sub AnsiRed_Click()
If chkBGColor.Value = 1 Then
    InsertAnsiCode ("[0;41m")
Else
    InsertAnsiCode ("[0;31m")
End If
End Sub

Private Sub AnsiWhite_Click()
    InsertAnsiCode ("[1;37m")
End Sub

Private Sub AnsiYellow_Click()
    InsertAnsiCode ("[1;33m")
End Sub
Private Sub AnsiBBlue_Click()
    InsertAnsiCode ("[1;34m")
End Sub

Private Sub AnsiBCyan_Click()
    InsertAnsiCode ("[1;36m")
End Sub

Private Sub AnsiBGreen_Click()
    InsertAnsiCode ("[1;32m")
End Sub

Private Sub AnsiBlack_Click()
If chkBGColor.Value = 1 Then
    InsertAnsiCode ("[0;40m")
Else
    InsertAnsiCode ("[0;30m")
End If
End Sub

Private Sub AnsiBlue_Click()
If chkBGColor.Value = 1 Then
    InsertAnsiCode ("[0;44m")
Else
    InsertAnsiCode ("[0;34m")
End If
End Sub

Private Sub AnsiBRed_Click()
    InsertAnsiCode ("[1;31m")
End Sub

Private Sub AnsiBrown_Click()
If chkBGColor.Value = 1 Then
    InsertAnsiCode ("[0;43m")
Else
    InsertAnsiCode ("[0;33m")
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next

If Me.WindowState = vbMinimized Then Exit Sub

fraMain.Width = Me.Width
fraMain.Height = Me.Height
fraEdit.Width = Me.Width
fraEdit.Height = Me.Height
txtAnsi.Width = Me.Width - 135
txtAnsi.Height = Me.Height - 2900 - TITLEBAR_OFFSET

'lvLines.Height = fraEdit.Height - lvLines.Top - 450 - TITLEBAR_OFFSET
splSplitter.Width = Me.Width
splSplitter.Height = Me.Height - 950 - TITLEBAR_OFFSET

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If fraEdit.Visible = True And bCheckSave Then
    MsgBox "Close textblock line editor first.", vbInformation
    Cancel = 1
    Exit Sub
End If
    
If bLoaded = True Then SaveTextBlock
Unload frmPreview

'Call WriteINI("Windows", "TextblockAutoUpdateLine", chkAutoLine.Value)

If Me.WindowState = vbMinimized Then Exit Sub

If Me.WindowState = vbMaximized Then
    Call WriteINI("Windows", "TextMaxed", 1)
Else
    Call WriteINI("Windows", "TextMaxed", 0)
    Call WriteINI("Windows", "TextTop", Me.Top)
    Call WriteINI("Windows", "TextLeft", Me.Left)
    Call WriteINI("Windows", "TextWidth", Me.Width)
    Call WriteINI("Windows", "TextHeight", Me.Height)
End If
End Sub

Private Sub InsertAnsiCode(AnsiCode As String)
Dim TextboxPos As Long, TextboxPart1 As String, TextboxPart2 As String
TextboxPos = txtAnsi.SelStart

If TextboxPos = 0 Then
    TextboxPart1 = ""
    TextboxPart2 = Me.txtAnsi.Text
Else
    TextboxPart1 = Left(Me.txtAnsi.Text, TextboxPos)
    TextboxPart2 = Right(Me.txtAnsi.Text, Len(Me.txtAnsi.Text) - TextboxPos)
End If

Me.txtAnsi.Text = TextboxPart1 & AnsiCode & TextboxPart2
Me.txtAnsi.SetFocus
Me.txtAnsi.SelStart = TextboxPos + Len(AnsiCode)
End Sub

Public Sub cmdShowPreview_Click()

Call DispPreview

frmPreview.SetFocus
End Sub
Private Sub DispPreview()
Dim i As Integer

term_Carethide
frmPreview.CursorTimer.Enabled = False
term_reset_matrix
term_eraseSCREEN
modANSIStuff.MyBackColor = 40
modANSIStuff.MyForeColor = 32
modANSIStuff.isBold = False

For i = 1 To Len(Me.txtAnsi.Text)
    term_process_char (Asc(Mid(Me.txtAnsi.Text, i, 1)))
Next i
modANSIStuff.CurState = False
'frmPreview.CursorTimer.Enabled = True
term_Caretshow
frmPreview.refresh

End Sub
'Private Sub PreviouslyViewed()
'Dim nStatus As Integer, nNumber As Long
'
'TextblockKey.PartNum = Val(txtPart.Text)
'TextblockKey.Number = Val(txtNumber.Text)
'
'nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKey, KEY_BUF_LEN, 0)
'If Not nStatus = 0 Then
'    MsgBox "Error going to previously viewed record, close and open the textblock form."
'    Exit Sub
'Else
'    DispTextblockInfo TextblockDataBuf.buf
'End If
'End Sub

Private Sub cmdDiscard_Click()
Dim nStatus As Integer

TextblockKey.Number = Val(txtNumber.Text)
TextblockKey.PartNum = Val(txtPart.Text)

nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKey, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error getting current record." & BtrieveErrorCode(nStatus)
    Exit Sub
Else
    DispTextblockInfo TextblockDataBuf.buf
End If


End Sub

Private Sub cmdSave_Click()

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

Call SaveTextBlock

End Sub

Private Sub HScroll1_Change()

End Sub

Private Sub lvLines_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo error:
Dim sData As String, nPos As Long, x As Long, sNewData As String, sLine As String
Dim bTemp As Boolean

bTemp = bCheckSave
If lvLines.ListItems.Count < 1 Then Exit Sub
If lvLines.SelectedItem Is Nothing Then Exit Sub

sData = lvLines.SelectedItem.Text

x = 1
nPos = 1
Do While InStr(nPos, sData, ":") > 0
    x = InStr(nPos, sData, ":")
    If nPos = x Then
        sNewData = sNewData & vbCrLf
        nPos = x + 1
        GoTo NextLine:
    End If
    'Debug.Print Mid(sData, nPos, x - nPos)
    sLine = Mid(sData, nPos, x - nPos)
    'If Right(sLine, 1) = Chr(13) Then sLine = Mid(sLine, 1, Len(sLine) - 1)
    sNewData = sNewData & sLine & vbCrLf
    nPos = x + 1
NextLine:
Loop

If nPos <= Len(sData) Then
    sLine = Mid(sData, nPos)
    'If Right(sLine, 1) = Chr(13) Then sLine = Mid(sLine, 1, Len(sLine) - 1)
    sNewData = sNewData & sLine & vbCrLf
End If

txtAnsiEdit.Text = sNewData
bCheckSave = bTemp
Exit Sub
error:
Call HandleError("lvLines_ItemClick")
End Sub

Private Sub splSplitter_Resize()
txtAnsiEdit.Width = fraRight.Width - 1000 '- txtAnsiEdit.Left - 200
txtAnsiEdit.Height = fraRight.Height '- txtAnsiEdit.Top - 450 - TITLEBAR_OFFSET
End Sub

Private Sub txtAnsi_Change()
lblCharCount.Caption = 2000 - Len(txtAnsi.Text)
End Sub

Private Sub txtAnsi_KeyUp(KeyCode As Integer, Shift As Integer)
'lblCharCount.Caption = 2000 - Len(txtAnsi.Text)
End Sub

Private Sub txtAnsiEdit_Change()
Dim x As Long, sData As String
On Error GoTo error:

If lvLines.SelectedItem Is Nothing Then Exit Sub

bCheckSave = True

sData = RemoveCharacter(txtAnsiEdit.Text, Chr(13))
x = InStr(1, sData, Chr(10))
Do While x > 0
    Mid(sData, x, 1) = ":"
    x = InStr(1, sData, Chr(10))
Loop

check_again:
If Right(sData, 1) = Chr(10) Then
    sData = Left(sData, Len(sData) - 1)
    GoTo check_again:
ElseIf Right(sData, 1) = ":" Then
    sData = Left(sData, Len(sData) - 1)
    GoTo check_again:
End If

lvLines.SelectedItem.Text = sData
Call CalcChars

Exit Sub
error:
Call HandleError("txtAnsiEdit_Change")
End Sub

Private Sub txtGotoBlock_GotFocus()
Call SelectAll(txtGotoBlock)

End Sub

Private Sub txtGotoBlock_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call cmdGoto_Click
End Sub

Private Sub txtGotoBlock_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtGotoInLine_GotFocus()
Call SelectAll(txtGotoInLine)
End Sub

Private Sub txtGotoInLine_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call cmdGotoInLine_Click
End Sub

Private Sub txtGotoInLine_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtGotoPart_GotFocus()
Call SelectAll(txtGotoPart)

End Sub

Private Sub txtGotoPart_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call cmdGoto_Click
End Sub

Private Sub txtGotoPart_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtLinkTo_GotFocus()
Call SelectAll(txtLinkTo)

End Sub

Private Sub txtNumber_GotFocus()
Call SelectAll(txtNumber)

End Sub

Private Sub txtPart_GotFocus()
Call SelectAll(txtPart)

End Sub
