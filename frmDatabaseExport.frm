VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDatabaseExport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Exporter"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12885
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDatabaseExport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   12885
   Begin VB.CommandButton cmdExportAllToggle 
      Caption         =   "Export All's - OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   11040
      TabIndex        =   64
      Top             =   3120
      Width           =   1755
   End
   Begin VB.CommandButton cmdExportAllToggle 
      Caption         =   "Export All's - ON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   9060
      TabIndex        =   63
      Top             =   3120
      Width           =   1755
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   1335
      Index           =   1
      Left            =   3240
      TabIndex        =   51
      Top             =   1260
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   2355
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   1335
      Index           =   2
      Left            =   6060
      TabIndex        =   52
      Top             =   1260
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   2355
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   1335
      Index           =   5
      Left            =   6060
      TabIndex        =   53
      Top             =   2940
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   2355
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   1335
      Index           =   0
      Left            =   360
      TabIndex        =   54
      Top             =   1260
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   2355
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   1335
      Index           =   3
      Left            =   360
      TabIndex        =   55
      Top             =   2940
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   2355
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   1335
      Index           =   4
      Left            =   3240
      TabIndex        =   56
      Top             =   2940
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   2355
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   1335
      Index           =   6
      Left            =   360
      TabIndex        =   57
      Top             =   4620
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   2355
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   1335
      Index           =   7
      Left            =   3240
      TabIndex        =   58
      Top             =   4620
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   2355
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   1335
      Index           =   8
      Left            =   6060
      TabIndex        =   50
      Top             =   4620
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   2355
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Frame fraConfig 
      Caption         =   "Config File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   9060
      TabIndex        =   46
      Top             =   60
      Width           =   3735
      Begin VB.TextBox txtCustom 
         Height          =   345
         Left            =   120
         MaxLength       =   20
         TabIndex        =   48
         Text            =   "Custom Export"
         Top             =   540
         Width           =   3495
      End
      Begin VB.TextBox txtConfigFile 
         BackColor       =   &H8000000F&
         Height          =   795
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   47
         Top             =   900
         Width           =   3495
      End
      Begin VB.CommandButton cmdSelectConfig 
         Caption         =   "&Load Config ..."
         Height          =   435
         Left            =   120
         TabIndex        =   20
         Top             =   1740
         Width           =   1695
      End
      Begin VB.CommandButton cmdSaveConfig 
         Caption         =   "&Save Config ..."
         Height          =   435
         Left            =   1920
         TabIndex        =   21
         Top             =   1740
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Identifier"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   49
         Top             =   300
         Width           =   765
      End
   End
   Begin VB.Frame fra1 
      Caption         =   "What to Export"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   120
      TabIndex        =   32
      Top             =   60
      Width           =   8775
      Begin VB.CommandButton cmdGetFirstLast 
         Height          =   195
         Index           =   1
         Left            =   5220
         TabIndex        =   61
         Top             =   240
         Width           =   195
      End
      Begin VB.CommandButton cmdGetFirstLast 
         Height          =   195
         Index           =   0
         Left            =   4320
         TabIndex        =   60
         Top             =   240
         Width           =   195
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "Clear All"
         Height          =   315
         Left            =   7680
         TabIndex        =   7
         Top             =   480
         Width           =   915
      End
      Begin VB.CheckBox chkExportAll 
         Caption         =   "Export All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   7440
         TabIndex        =   16
         Top             =   4320
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkExportAll 
         Caption         =   "Export All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   4560
         TabIndex        =   15
         Top             =   4320
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkExportAll 
         Caption         =   "Export All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   1680
         TabIndex        =   14
         Top             =   4320
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkExportAll 
         Caption         =   "Export All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   7440
         TabIndex        =   13
         Top             =   2640
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkExportAll 
         Caption         =   "Export All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   4560
         TabIndex        =   12
         Top             =   2640
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkExportAll 
         Caption         =   "Export All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1680
         TabIndex        =   11
         Top             =   2640
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkExportAll 
         Caption         =   "Export All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   7440
         TabIndex        =   10
         Top             =   960
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkExportAll 
         Caption         =   "Export All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   4560
         TabIndex        =   9
         Top             =   960
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkExportAll 
         Caption         =   "Export All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1680
         TabIndex        =   8
         Top             =   960
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkActions 
         Caption         =   "Actions (all)"
         Height          =   255
         Left            =   720
         TabIndex        =   17
         Top             =   6120
         Width           =   1575
      End
      Begin VB.CheckBox chkUsers 
         Caption         =   "Users (all)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3600
         TabIndex        =   18
         Top             =   6120
         Width           =   1395
      End
      Begin VB.CheckBox chkBankbooks 
         Caption         =   "BankBooks (all)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   6360
         TabIndex        =   19
         Top             =   6120
         Width           =   1695
      End
      Begin VB.TextBox txtMap 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3060
         TabIndex        =   1
         Text            =   "1"
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   315
         Left            =   7680
         TabIndex        =   6
         Top             =   180
         Width           =   915
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   495
         Left            =   6540
         TabIndex        =   5
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   495
         Left            =   5580
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtTo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4680
         TabIndex        =   3
         Text            =   "1"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtFrom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3780
         TabIndex        =   2
         Text            =   "1"
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox cmbDB 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   2595
      End
      Begin VB.Shape shpOutline 
         BorderWidth     =   3
         Height          =   1365
         Index           =   8
         Left            =   5925
         Top             =   4545
         Width           =   2625
      End
      Begin VB.Shape shpOutline 
         BorderWidth     =   3
         Height          =   1365
         Index           =   7
         Left            =   3105
         Top             =   4545
         Width           =   2625
      End
      Begin VB.Shape shpOutline 
         BorderWidth     =   3
         Height          =   1365
         Index           =   6
         Left            =   225
         Top             =   4545
         Width           =   2625
      End
      Begin VB.Shape shpOutline 
         BorderWidth     =   3
         Height          =   1365
         Index           =   5
         Left            =   5925
         Top             =   2865
         Width           =   2625
      End
      Begin VB.Shape shpOutline 
         BorderWidth     =   3
         Height          =   1365
         Index           =   4
         Left            =   3105
         Top             =   2865
         Width           =   2625
      End
      Begin VB.Shape shpOutline 
         BorderWidth     =   3
         Height          =   1365
         Index           =   3
         Left            =   225
         Top             =   2865
         Width           =   2625
      End
      Begin VB.Shape shpOutline 
         BorderWidth     =   3
         Height          =   1365
         Index           =   2
         Left            =   5925
         Top             =   1185
         Width           =   2625
      End
      Begin VB.Shape shpOutline 
         BorderWidth     =   3
         Height          =   1365
         Index           =   1
         Left            =   3105
         Top             =   1185
         Width           =   2625
      End
      Begin VB.Shape shpOutline 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         Height          =   1365
         Index           =   0
         Left            =   225
         Top             =   1185
         Width           =   2625
      End
      Begin VB.Line Line1 
         X1              =   300
         X2              =   8520
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Label lblDB 
         Caption         =   "Database:"
         Height          =   255
         Index           =   8
         Left            =   5940
         TabIndex        =   45
         Top             =   4320
         Width           =   1155
      End
      Begin VB.Label lblDB 
         Caption         =   "Database:"
         Height          =   255
         Index           =   7
         Left            =   3120
         TabIndex        =   44
         Top             =   4320
         Width           =   1155
      End
      Begin VB.Label lblDB 
         Caption         =   "Database:"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   43
         Top             =   4320
         Width           =   1155
      End
      Begin VB.Label lblDB 
         Caption         =   "Database:"
         Height          =   255
         Index           =   5
         Left            =   5940
         TabIndex        =   42
         Top             =   2640
         Width           =   1155
      End
      Begin VB.Label lblDB 
         Caption         =   "Database:"
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   41
         Top             =   2640
         Width           =   1155
      End
      Begin VB.Label lblDB 
         Caption         =   "Database:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   40
         Top             =   2640
         Width           =   1155
      End
      Begin VB.Label lblDB 
         Caption         =   "Database:"
         Height          =   255
         Index           =   2
         Left            =   5940
         TabIndex        =   39
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label lblDB 
         Caption         =   "Database:"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   38
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label lblMap 
         Caption         =   "Map"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3060
         TabIndex        =   37
         Top             =   240
         Width           =   495
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   240
         X2              =   8580
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label lblDB 
         Caption         =   "Database:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   36
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Database:"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   34
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   255
         Index           =   0
         Left            =   3780
         TabIndex        =   33
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdImportRecordNumbers 
      Caption         =   "Import Record Numbers from Database..."
      Height          =   555
      Left            =   9060
      TabIndex        =   22
      Top             =   2460
      Width           =   3735
   End
   Begin VB.Frame fra2 
      Caption         =   "File Format"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   9060
      TabIndex        =   29
      Top             =   3720
      Width           =   3735
      Begin VB.CommandButton cmdUserInteractionQ 
         Caption         =   "?"
         Height          =   315
         Left            =   3180
         TabIndex        =   62
         Top             =   1380
         Width           =   330
      End
      Begin VB.CheckBox chkZeroUserInteraction 
         Caption         =   """Reset"" User Interactable Fields on Export (Cash, Item Uses, Etc)"
         Height          =   735
         Left            =   480
         TabIndex        =   59
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2595
      End
      Begin VB.CommandButton cmdQ 
         Caption         =   "?"
         Height          =   315
         Left            =   2340
         TabIndex        =   28
         Top             =   780
         Width           =   330
      End
      Begin VB.CheckBox chkOneExpField 
         Caption         =   "Use 1 Field for Mon EXP"
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
         Left            =   480
         TabIndex        =   25
         ToolTipText     =   "This is just for people who need the experience in one field when doing advanced operations outside of NMR."
         Top             =   900
         Width           =   1815
      End
      Begin VB.OptionButton optAccessDB 
         Caption         =   "Access Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   600
         Value           =   -1  'True
         Width           =   1635
      End
      Begin VB.OptionButton optTextfile 
         Caption         =   "Textfiles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   300
         Width           =   1035
      End
   End
   Begin MSComctlLib.StatusBar stsStatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   31
      Top             =   6975
      Width           =   12885
      _ExtentX        =   22728
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20108
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
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   315
      Left            =   120
      TabIndex        =   30
      Top             =   6660
      Visible         =   0   'False
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel / Close"
      Height          =   495
      Left            =   11040
      TabIndex        =   27
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Export"
      Height          =   495
      Left            =   9060
      TabIndex        =   26
      Top             =   6000
      Width           =   1515
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10560
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmDatabaseExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim DB As Database
Dim tabActions As Recordset
Dim tabMessages As Recordset
Dim tabTextblocks As Recordset
Dim tabItems As Recordset
Dim tabClasses As Recordset
Dim tabRaces As Recordset
Dim tabSpells As Recordset
Dim tabInfo As Recordset
Dim tabMonsters As Recordset
Dim tabShops As Recordset
Dim tabRooms As Recordset

Dim bStopExport As Boolean
Dim bCheckSave As Boolean

Dim nScale As Integer
Dim nScaleCount As Long

Dim bUpdateExistingADB As Boolean
Dim sDataSource As String
Dim MessagesTextfile As String
Dim ItemsTextfile As String
Dim SpellsTextfile As String
Dim ClassesTextfile As String
Dim RacesTextfile As String
Dim ShopsTextfile As String
Dim RoomsTextfile As String
Dim ActionsTextfile As String
Dim MonstersTextfile As String
Dim UsersTextfile As String
Dim BankbooksTextfile As String
Dim TextblocksTextfile As String
Dim sExportPath As String
Dim sConfigFile As String

Private Sub SetRange(ByVal MaxValue As Double)
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
End Sub

Private Sub chkActions_Click()
bCheckSave = True
End Sub

Private Sub chkBankbooks_Click()
bCheckSave = True
End Sub

Private Sub chkExportAll_Click(Index As Integer)
bCheckSave = True
Call UpdateListStuff
End Sub

Private Sub chkOneExpField_Click()
bCheckSave = True
End Sub

Private Sub chkUsers_Click()
bCheckSave = True
End Sub

Private Sub chkZeroUserInteraction_Click()
bCheckSave = True
End Sub

Private Sub cmbDB_Click()
Dim bEn As Boolean

If cmbDB.ListIndex = 8 Then bEn = True
lblMap.Enabled = bEn
txtMap.Enabled = bEn
Call UpdateListStuff
End Sub

Private Sub cmdAdd_Click()
Dim oLI As ListItem
Dim nFrom As Long, nTo As Long, nMap As Long
On Error GoTo error:

If cmbDB.ListIndex < 0 Then Exit Sub

nFrom = Val(txtFrom.Text)
nTo = Val(txtTo.Text)
nMap = Val(txtMap.Text)

If nTo < 1 Then Exit Sub
If nFrom < 1 And Not cmbDB.ListIndex = 7 Then Exit Sub
If nTo < nFrom Then Exit Sub
If cmbDB.ListIndex = 8 And nMap < 1 Then Exit Sub

Set oLI = lvList(cmbDB.ListIndex).ListItems.add
If cmbDB.ListIndex = 8 Then
    oLI.Text = nMap
    oLI.ListSubItems.add 1, , nFrom
    oLI.ListSubItems.add 2, , nTo
Else
    oLI.Text = nFrom
    oLI.ListSubItems.add 1, , nTo
End If
chkExportAll(cmbDB.ListIndex).Value = 0

Call CombineRanges

bCheckSave = True

out:
Set oLI = Nothing
Exit Sub
error:
Call HandleError("cmdAdd_Click")
Resume out:
End Sub

Private Sub cmdClear_Click()
If cmbDB.ListIndex < 0 Then Exit Sub
lvList(cmbDB.ListIndex).ListItems.clear
bCheckSave = True
End Sub

Private Sub cmdClearAll_Click()
Dim nYesNo As Integer, x As Integer
nYesNo = MsgBox("Clear all lists?", vbQuestion + vbYesNo)
If nYesNo <> vbYes Then Exit Sub
For x = 0 To 8
    lvList(x).ListItems.clear
    chkExportAll(x).Value = 1
Next x
bCheckSave = True
End Sub

Private Sub cmdExportAllToggle_Click(Index As Integer)
Dim x As Integer, nOP As Integer
If Index = 0 Then nOP = 1
For x = 0 To 8
    chkExportAll(x).Value = nOP
Next x
End Sub

Private Sub cmdGetFirstLast_Click(Index As Integer)
On Error GoTo error:
Dim nStatus As Integer, nOP As Integer, nRET As Long

cmdGetFirstLast(0).Enabled = False
cmdGetFirstLast(1).Enabled = False

If Index = 0 Then
    nOP = BGETFIRST
Else
    nOP = BGETLAST
End If

nRET = -1

Select Case cmbDB.ListIndex
    Case 0:
        nStatus = BTRCALL(nOP, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            MsgBox "Error getting first class record: " & BtrieveErrorCode(nStatus)
        Else
            ClassRowToStruct Classdatabuf.buf
            nRET = Classrec.Number
        End If
    Case 1:
        nStatus = BTRCALL(nOP, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            MsgBox "Error getting first race record: " & BtrieveErrorCode(nStatus)
        Else
            RaceRowToStruct Racedatabuf.buf
            nRET = Racerec.Number
        End If
    Case 2:
        nStatus = BTRCALL(nOP, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            MsgBox "Error getting first item record: " & BtrieveErrorCode(nStatus)
        Else
            ItemRowToStruct Itemdatabuf.buf
            nRET = Itemrec.Number
        End If
    Case 3:
        nStatus = BTRCALL(nOP, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            MsgBox "Error getting first message record: " & BtrieveErrorCode(nStatus)
        Else
            MessageRowToStruct Messagedatabuf.buf
            nRET = Messagerec.Number
        End If
    Case 4:
        nStatus = BTRCALL(nOP, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            MsgBox "Error getting first monster record: " & BtrieveErrorCode(nStatus)
        Else
            MonsterRowToStruct Monsterdatabuf.buf
            nRET = Monsterrec.Number
        End If
    Case 5:
        nStatus = BTRCALL(nOP, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            MsgBox "Error getting first shop record: " & BtrieveErrorCode(nStatus)
        Else
            ShopRowToStruct Shopdatabuf.buf
            nRET = Shoprec.Number
        End If
    Case 6:
        nStatus = BTRCALL(nOP, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            MsgBox "Error getting first spell record: " & BtrieveErrorCode(nStatus)
        Else
            SpellRowToStruct Spelldatabuf.buf
            nRET = Spellrec.Number
        End If
    Case 7:
        nStatus = BTRCALL(nOP, TextblockPosBlock, TextblockDataBuf, Len(TextblockDataBuf), ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            MsgBox "Error getting first textblock record: " & BtrieveErrorCode(nStatus)
        Else
            TextblockRowToStruct TextblockDataBuf.buf
            nRET = TextblockRec.Number
        End If
    Case 8:
        If Val(txtMap.Text) > 0 Then
            
            nStatus = BTRCALL(BSTAT, RoomPosBlock, DBStatDatabuf, Len(Roomdatabuf), 0, KEY_BUF_LEN, 0)
            If nStatus = 0 Then
                DBStatRowToStruct DBStatDatabuf.buf
                Call SetRange(DBStat.nRecords)
            Else
                Call SetRange(30000)
            End If
            ProgressBar.Visible = True
            DoEvents
                
            If Index = 0 Then
                nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
                If nStatus = 0 Then
                    Do While nStatus = 0
                        Call IncreaseProgressBar
                        RoomRowToStruct Roomdatabuf.buf
                        If Roomrec.MapNumber = Val(txtMap.Text) Then
                            nRET = Roomrec.RoomNumber
                            Exit Do
                        End If
                        nStatus = BTRCALL(BGETNEXT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
                        If Not bUseCPU Then DoEvents
                    Loop
                Else
                    MsgBox "Error getting first room.", vbExclamation
                End If
            Else
                nStatus = BTRCALL(BGETLAST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
                If nStatus = 0 Then
                    Do While nStatus = 0
                        Call IncreaseProgressBar
                        RoomRowToStruct Roomdatabuf.buf
                        If Roomrec.MapNumber = Val(txtMap.Text) Then
                            nRET = Roomrec.RoomNumber
                            Exit Do
                        End If
                        nStatus = BTRCALL(BGETPREVIOUS, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
                        If Not bUseCPU Then DoEvents
                    Loop
                Else
                    MsgBox "Error getting first room.", vbExclamation
                End If
            End If
        Else
            MsgBox "Enter map number.", vbExclamation
        End If
End Select

If nRET >= 0 Then
    If Index = 0 Then
        txtFrom.Text = nRET
    Else
        txtTo.Text = nRET
    End If
End If

out:
On Error Resume Next
If cmbDB.ListIndex = 8 Then
    ProgressBar.Value = ProgressBar.Max
    ProgressBar.Visible = False
End If
cmdGetFirstLast(0).Enabled = True
cmdGetFirstLast(1).Enabled = True
Exit Sub
error:
Call HandleError("cmdGetFirstLast_Click")
Resume out:
End Sub

Private Sub cmdImportRecordNumbers_Click()
On Error GoTo error:
Dim oLI As ListItem, nTotalRec As Long
Dim sTemp As String, nYesNo As Integer, catDB As ADOX.Catalog
Dim fso As FileSystemObject, x As Integer, y As Integer, nTemp As Integer
Dim nLastMap As Long, nFirstRoom As Long, nLastRoom As Long

Set fso = CreateObject("Scripting.FileSystemObject")
sExportPath = ReadINI("Options", "ExportPath")
If Not fso.FolderExists(sExportPath) Then sExportPath = App.Path

sTemp = ReadINI("Options", "ExportFileName")
If Len(sTemp) < 5 Then sTemp = "NMR-DataExport.mdb"

CommonDialog1.Filter = "MDB Files (*.mdb)|*.mdb"
CommonDialog1.DialogTitle = "Select Export File"
CommonDialog1.FileName = sTemp
CommonDialog1.InitDir = sExportPath

On Error GoTo out:
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then GoTo out:

On Error GoTo error:
sDataSource = CommonDialog1.FileName

If Not LCase(Right(sDataSource, 4)) = ".mdb" Then sDataSource = sDataSource & ".mdb"

sTemp = CommonDialog1.FileTitle
If Not LCase(Right(sTemp, 4)) = ".mdb" Then sTemp = sTemp & ".mdb"

If Not fso.FileExists(sDataSource) = True Then
    MsgBox "File not found!", vbExclamation
    GoTo out:
End If

Set tabRooms = Nothing
Set tabItems = Nothing
Set tabClasses = Nothing
Set tabRaces = Nothing
Set tabSpells = Nothing
Set tabMonsters = Nothing
Set tabShops = Nothing
Set tabMessages = Nothing
Set tabTextblocks = Nothing

Set DB = OpenDatabase(sDataSource)
Call OpenTables

nTotalRec = 0
If Not tabRooms Is Nothing Then nTotalRec = nTotalRec + tabRooms.RecordCount
If Not tabItems Is Nothing Then nTotalRec = nTotalRec + tabItems.RecordCount
If Not tabClasses Is Nothing Then nTotalRec = nTotalRec + tabClasses.RecordCount
If Not tabRaces Is Nothing Then nTotalRec = nTotalRec + tabRaces.RecordCount
If Not tabSpells Is Nothing Then nTotalRec = nTotalRec + tabSpells.RecordCount
If Not tabMonsters Is Nothing Then nTotalRec = nTotalRec + tabMonsters.RecordCount
If Not tabShops Is Nothing Then nTotalRec = nTotalRec + tabShops.RecordCount
If Not tabMessages Is Nothing Then nTotalRec = nTotalRec + tabMessages.RecordCount
If Not tabTextblocks Is Nothing Then nTotalRec = nTotalRec + tabTextblocks.RecordCount

Call SetRange(nTotalRec)
ProgressBar.Visible = True
DoEvents

Call SetRangeFromDB(tabClasses, 0, "pkClasses")
Call SetRangeFromDB(tabRaces, 1, "pkRaces")
Call SetRangeFromDB(tabItems, 2, "pkItems")
Call SetRangeFromDB(tabMessages, 3, "pkMessages")
Call SetRangeFromDB(tabMonsters, 4, "pkMonsters")
Call SetRangeFromDB(tabSpells, 5, "pkSpells")
Call SetRangeFromDB(tabShops, 6, "pkShops")
Call SetRangeFromDB(tabTextblocks, 7, "idxTextblocks")

For x = 0 To 8
    If lvList(x).ListItems.Count > 0 Then SortListView lvList(x), 1, ldtNumber, True
Next x

If tabRooms Is Nothing Then GoTo norooms:
If tabRooms.RecordCount = 0 Then GoTo norooms:

tabRooms.Index = "idxRooms"

nLastMap = 0
nFirstRoom = 0
nLastRoom = 0
tabRooms.MoveFirst
Do While tabRooms.EOF = False
    If nFirstRoom > 0 And nLastMap = tabRooms.Fields("Map Number") _
        And tabRooms.Fields("Room Number") = (nLastRoom + 1) Then
        nLastRoom = tabRooms.Fields("Room Number")
        
    ElseIf nLastMap > 0 And nLastRoom > 0 And nFirstRoom > 0 Then
        Set oLI = lvList(8).ListItems.add
        oLI.Text = nLastMap
        oLI.ListSubItems.add 1, , nFirstRoom
        oLI.ListSubItems.add 2, , nLastRoom
        
        nLastMap = tabRooms.Fields("Map Number")
        nFirstRoom = tabRooms.Fields("Room Number")
        nLastRoom = tabRooms.Fields("Room Number")
    Else
        nLastMap = tabRooms.Fields("Map Number")
        nFirstRoom = tabRooms.Fields("Room Number")
        nLastRoom = tabRooms.Fields("Room Number")
    End If
    tabRooms.MoveNext
    Call IncreaseProgressBar
    If Not bUseCPU Then DoEvents
Loop

If nLastMap > 0 And nLastRoom > 0 And nFirstRoom > 0 Then
    Set oLI = lvList(8).ListItems.add
    oLI.Text = nLastMap
    oLI.ListSubItems.add 1, , nFirstRoom
    oLI.ListSubItems.add 2, , nLastRoom
End If

chkExportAll(8).Value = 0
Call SortListView(lvList(8), 2, ldtNumber, True)
Call SortListView(lvList(8), 1, ldtNumber, True)
GoTo out:

norooms:

out:
On Error Resume Next
Call CombineRanges
ProgressBar.Visible = False
bCheckSave = True
Call CloseAll(True)
Set fso = Nothing
Set catDB = Nothing
Set DB = Nothing
Set oLI = Nothing
Exit Sub

error:
Call HandleError("cmdImportRecordNumbers_Click")
Resume out:
End Sub

Private Sub SetRangeFromDB(ByRef tabTable As Recordset, nIndex As Integer, sIndex As String)
Dim oLI As ListItem, nFirst As Long, nLast As Long
On Error GoTo error:

If tabTable Is Nothing Then GoTo zero:
If tabTable.RecordCount = 0 Then GoTo zero:

tabTable.Index = sIndex

tabTable.MoveFirst
nFirst = tabTable.Fields("Number")
nLast = tabTable.Fields("Number")
Do While Not tabTable.EOF
    If tabTable.Fields("Number") > nLast + 1 Then
        Set oLI = lvList(nIndex).ListItems.add
        oLI.Text = nFirst
        oLI.ListSubItems.add 1, , nLast
        
        nFirst = tabTable.Fields("Number")
        nLast = tabTable.Fields("Number")
    Else
        nLast = tabTable.Fields("Number")
    End If
    Call IncreaseProgressBar
    If Not bUseCPU Then DoEvents
    tabTable.MoveNext
Loop

Set oLI = lvList(nIndex).ListItems.add
oLI.Text = nFirst
oLI.ListSubItems.add 1, , nLast

chkExportAll(nIndex).Value = 0
GoTo out:

zero:

out:
Set oLI = Nothing
Exit Sub
error:
Call HandleError("SetRangeFromDB")
Resume out:
End Sub
Private Sub cmdQ_Click()
MsgBox "The ""Use 1 field for Mon EXP"" setting is for people who need the Monster 'EXP' and 'EXP Multiplier' fields" _
    & "multiplied together in the export for sorting purposes with 3rd party applications.", vbInformation
End Sub

Private Sub cmdRemove_Click()
If cmbDB.ListIndex < 0 Then Exit Sub
If lvList(cmbDB.ListIndex).SelectedItem Is Nothing Then Exit Sub
lvList(cmbDB.ListIndex).ListItems.Remove lvList(cmbDB.ListIndex).SelectedItem.Index
bCheckSave = True
End Sub

Private Sub cmdSaveConfig_Click()
Call SaveConfig(sConfigFile, True)
End Sub

Private Sub cmdSelectConfig_Click()
Dim sTemp As String, nTemp As Integer
On Error GoTo error:

If bCheckSave Then
    nTemp = MsgBox("Save current config file first?", vbYesNoCancel + vbQuestion, "Save Export Config?")
    If nTemp = vbYes Then
        nTemp = SaveConfig(sConfigFile)
        If nTemp = -1 Then Exit Sub
    ElseIf nTemp = vbCancel Then
        Exit Sub
    End If
End If

CommonDialog1.Filter = "INI Files (*.ini)|*.ini"
CommonDialog1.DialogTitle = "Select Export Configuration File ..."
CommonDialog1.FileName = sConfigFile
CommonDialog1.InitDir = sConfigFile

On Error GoTo canceled:
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then GoTo canceled:

On Error GoTo error:

sTemp = CommonDialog1.FileName
If Right(sTemp, 4) <> ".ini" Then sTemp = sTemp & ".ini"

Call LoadConfig(sTemp)

canceled:

out:
Exit Sub
error:
Call HandleError("cmdSelectConfig_Click")
Resume out:
End Sub

Private Sub cmdUserInteractionQ_Click()
Dim sText As String

sText = "Monsters: Active, Date Killed, Time Killed (all set to 0) " & vbCrLf _
& "Rooms: Visible/Hidden Cash and Items (all set to 0)" & vbCrLf _
& "Shops: Current stock set to... " & vbCrLf _
& "       ""ShopRegen%"" > 0 THEN ""NumToRegen"" ELSE 0"

MsgBox "Current fields that get reset--" & vbCrLf & vbCrLf & sText, vbInformation

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim x As Integer
Dim fso As FileSystemObject

Set fso = CreateObject("Scripting.FileSystemObject")

bCheckSave = False

cmbDB.clear
cmbDB.AddItem "Classes", 0
cmbDB.AddItem "Races", 1
cmbDB.AddItem "Items", 2
cmbDB.AddItem "Messages", 3
cmbDB.AddItem "Monsters", 4
cmbDB.AddItem "Shops", 5
cmbDB.AddItem "Spells", 6
cmbDB.AddItem "Textblocks", 7
cmbDB.AddItem "Rooms", 8

For x = 0 To 8
    lblDB(x).Caption = cmbDB.List(x) & ":"
    
    lvList(x).ColumnHeaders.clear
    If x = 8 Then
        lvList(x).ColumnHeaders.add , , "M", 400
        lvList(x).ColumnHeaders.add , , "From", 800
        lvList(x).ColumnHeaders.add , , "To", 800
    Else
        lvList(x).ColumnHeaders.add , , "From", 1000
        lvList(x).ColumnHeaders.add , , "To", 1000
    End If
Next x

If eDatFileVersion < v111j Then
    chkOneExpField.Value = 1
    chkOneExpField.Enabled = False
End If

sConfigFile = ReadINI("Options", "NMR-ExportConfig")
If Not fso.FileExists(sConfigFile) Then sConfigFile = App.Path & "\NMR-ExportConfig.ini"

Me.Top = ReadINI("Windows", "ExportTop")
Me.Left = ReadINI("Windows", "ExportLeft")

Call LoadConfig(sConfigFile)

Me.Show
Me.SetFocus

cmbDB.ListIndex = 0
cmdCancel.SetFocus

Set fso = Nothing
End Sub

Private Sub cmdGo_Click()
Dim objForm As Form
On Error GoTo error:
Dim FilenameArray(0 To 11) As String
Dim sNewPath() As String
Dim x As Integer, sPath As String, nFilesToExport As Long
Dim StartTime As Variant, nTotalTime As Double, sTotalTime As String
'UnloadForms (Me.Name)

nFilesToExport = 0
bStopExport = False
StartTime = Timer

Call SetRange(CalcTotalRecords)
ProgressBar.Visible = True
DoEvents

If optAccessDB.Value = True Then GoTo CreateAccessDB:

sExportPath = ReadINI("Options", "ExportPath")
sPath = BrowseForFolder(Me, "Select a Folder to export to", sExportPath)
If sPath = "" Then GoTo ReEnable:
sExportPath = sPath

If Right(sExportPath, 1) = "\" Then sExportPath = Left(sExportPath, Len(sExportPath) - 1)

Call WriteINI("Options", "ExportPath", sExportPath)

MessagesTextfile = sExportPath & "\NMR-Messages.txt"
ItemsTextfile = sExportPath & "\NMR-Items.txt"
SpellsTextfile = sExportPath & "\NMR-Spells.txt"
ClassesTextfile = sExportPath & "\NMR-Classes.txt"
RacesTextfile = sExportPath & "\NMR-Races.txt"
ShopsTextfile = sExportPath & "\NMR-Shops.txt"
RoomsTextfile = sExportPath & "\NMR-Rooms.txt"
ActionsTextfile = sExportPath & "\NMR-Actions.txt"
MonstersTextfile = sExportPath & "\NMR-Monsters.txt"
UsersTextfile = sExportPath & "\NMR-Users.txt"
BankbooksTextfile = sExportPath & "\NMR-Bankbooks.txt"
TextblocksTextfile = sExportPath & "\NMR-Textblocks.txt"

'If CheckFirstRecords = False Then GoTo ReEnable:

FilenameArray(0) = ClassesTextfile
FilenameArray(1) = RacesTextfile
FilenameArray(2) = ItemsTextfile
FilenameArray(3) = MessagesTextfile
FilenameArray(4) = MonstersTextfile
FilenameArray(5) = ShopsTextfile
FilenameArray(6) = SpellsTextfile
FilenameArray(7) = TextblocksTextfile
FilenameArray(8) = RoomsTextfile
FilenameArray(9) = ActionsTextfile
FilenameArray(10) = UsersTextfile
FilenameArray(11) = BankbooksTextfile

Call HideWindows

DoEvents

For x = 0 To 8
    If x = 8 Then
        SortListView lvList(x), 3, ldtNumber, True
    End If
    SortListView lvList(x), 2, ldtNumber, True
    SortListView lvList(x), 1, ldtNumber, True
    
    If bStopExport Then Exit For
    If chkExportAll(x).Value = 1 Or lvList(x).ListItems.Count > 0 Then
        Call CreateExportFile(FilenameArray(x))
        Select Case x
            Case 0:
                Call ExportClasses("textfile")
            Case 1:
                Call ExportRaces("textfile")
            Case 2:
                Call ExportItems("textfile")
            Case 3:
                Call ExportMessages("textfile")
            Case 4:
                Call ExportMonsters("textfile")
            Case 5:
                Call ExportShops("textfile")
            Case 6:
                Call ExportSpells("textfile")
            Case 7:
                Call ExportTextblocks("textfile")
            Case 8:
                Call ExportRooms("textfile")
        End Select
        DoEvents
    End If
Next
If chkActions.Value = 1 And Not bStopExport Then
    Call CreateExportFile(FilenameArray(9))
    Call ExportActions("textfile")
End If
DoEvents

If chkUsers.Value = 1 And Not bStopExport Then
    Call CreateExportFile(FilenameArray(10))
    Call ExportUsers
End If
DoEvents

If chkBankbooks.Value = 1 And Not bStopExport Then
    Call CreateExportFile(FilenameArray(11))
    Call ExportBankbooks
End If

If bStopExport Then GoTo ReEnable:

ProgressBar.Value = ProgressBar.Max

nTotalTime = Timer - StartTime
sTotalTime = CStr(Round(CDbl(nTotalTime / 60), 2))
sTotalTime = Left(sTotalTime, InStr(1, sTotalTime, ".") + 2)

MsgBox "Export Complete." & vbCrLf & vbCrLf & "Total time: " & sTotalTime & " minutes.", vbInformation

GoTo ReEnable:

CreateAccessDB:

Dim temp1 As Integer

temp1 = CreateDatabase

Select Case temp1
    Case 3: 'cancel
        GoTo ReEnable:
    
    Case 2: 'yes (update existing)
        bUpdateExistingADB = True
        
    Case 1: 'no (create new)
        stsStatusBar.Panels(2).Text = "Creating Tables..."
        If eDatFileVersion >= v111j And chkOneExpField.Value = 0 Then
            temp1 = CreateAccessTables(sDataSource, True)
        Else
            temp1 = CreateAccessTables(sDataSource, False)
        End If
        
        If temp1 = False Then
            MsgBox "Access Database was not created successfully."
            GoTo ReEnable:
        End If
    Case Else: 'else
        MsgBox "Access Database was not created successfully."
        GoTo ReEnable:
End Select

Call HideWindows

If InStr(1, sDataSource, "\") > 0 Then
    sNewPath = Split(sDataSource, "\")
    sExportPath = sNewPath(LBound(sNewPath()))
    For x = LBound(sNewPath()) + 1 To UBound(sNewPath()) - 1
        sExportPath = sExportPath & "\" & sNewPath(x)
    Next x
    'MsgBox sExportPath
    Call WriteINI("Options", "ExportPath", sExportPath)
End If
Erase sNewPath()

Set DB = OpenDatabase(sDataSource)
Call OpenTables

If bUpdateExistingADB = True Then
    If CheckVersion = False Then
        Call CloseAll(True)
        GoTo ReEnable:
    End If
End If
DoEvents

For x = 0 To 8
    If x = 8 Then
        SortListView lvList(x), 3, ldtNumber, True
    End If
    SortListView lvList(x), 2, ldtNumber, True
    SortListView lvList(x), 1, ldtNumber, True
    
    If bStopExport Then Exit For
    If chkExportAll(x).Value = 1 Or lvList(x).ListItems.Count > 0 Then
        Select Case x
            Case 0:
                Call ExportClasses("Access")
            Case 1:
                Call ExportRaces("Access")
            Case 2:
                Call ExportItems("Access")
            Case 3:
                Call ExportMessages("Access")
            Case 4:
                Call ExportMonsters("Access")
            Case 5:
                Call ExportShops("Access")
            Case 6:
                Call ExportSpells("Access")
            Case 7:
                Call ExportTextblocks("Access")
            Case 8:
                Call ExportRooms("Access")
        End Select
        DoEvents
    End If
Next
If chkActions.Value = 1 And Not bStopExport Then
    Call ExportActions("Access")
End If
DoEvents

Call ExportVersionInfo
Call CloseAll

If bStopExport Then GoTo ReEnable:

ProgressBar.Value = ProgressBar.Max

nTotalTime = Timer - StartTime
sTotalTime = CStr(Round(CDbl(nTotalTime / 60), 2))
sTotalTime = Left(sTotalTime, InStr(1, sTotalTime, ".") + 2)

MsgBox "Export Complete." & vbCrLf & vbCrLf & "Total time: " & sTotalTime & " minutes.", vbInformation

ReEnable:
On Error Resume Next

For Each objForm In Forms
    If Not objForm Is Me And Not objForm Is frmMain Then
        objForm.Enabled = True
    End If
Next

Set objForm = Nothing
Call UnLockMenus
fra1.Enabled = True
fra2.Enabled = True

fraConfig.Enabled = True
cmdImportRecordNumbers.Enabled = True
For x = 0 To 8
    lvList(x).Enabled = True
Next x

frmMain.Enabled = True
ProgressBar.Visible = False
stsStatusBar.Panels(1).Text = ""
stsStatusBar.Panels(2).Text = ""
cmdGo.Enabled = True
cmdCancel.Enabled = True
Me.WindowState = vbNormal
'Me.Show
Me.SetFocus

Exit Sub

error:
Call HandleError("cmdGo_Click")
Resume error2:

error2:
On Error Resume Next
Call CloseAll(True)
GoTo ReEnable:

End Sub

Private Function HideWindows()
Dim objForm As Form, x As Integer
On Error Resume Next

For Each objForm In Forms
    If Not objForm Is Me And Not objForm Is frmMain Then
        objForm.WindowState = vbMinimized
        objForm.Enabled = False
        If frmMain.tbTaskBar.Visible Then objForm.Hide
    End If
Next

cmdGo.Enabled = False
fra1.Enabled = False
fra2.Enabled = False

fraConfig.Enabled = False
cmdImportRecordNumbers.Enabled = False
For x = 0 To 8
    lvList(x).Enabled = False
Next x

Call LockMenus

Set objForm = Nothing

DoEvents

End Function
Private Function CheckVersion() As Boolean
On Error GoTo error:
Dim nYesNo As Integer, sVer As String, sCurrentVer As String, sNMRVer As String

CheckVersion = False

If tabInfo.RecordCount = 0 Then
    nYesNo = MsgBox("Unable to verify export file version information, continue anyway?", vbYesNo + vbQuestion)
    If nYesNo = vbYes Then CheckVersion = True
    Exit Function
End If

tabInfo.MoveLast
sVer = tabInfo.Fields("Dat File Version")
sNMRVer = tabInfo.Fields("NMR Version")
sCurrentVer = FriendlyDatVersion(eDatFileVersion)

If Not sVer = sCurrentVer Or Not sNMRVer = sAppVersion Then
    nYesNo = MsgBox("Warning, current NMR Version/Dat File Version does not match the export file's versions." & vbCrLf _
        & "Current: " & sAppVersion & "/" & sCurrentVer & ", Export file: " & sNMRVer & "/" & sVer & vbCrLf & vbCrLf _
        & "Often the export database is updated and changed between releases as new fields are found." & vbCrLf _
        & "Errors may occur, Continue anyway?", vbYesNo + vbQuestion)
    If nYesNo = vbNo Then Exit Function
End If

CheckVersion = True

Exit Function
error:
Call HandleError
nYesNo = MsgBox("Unable to verify export file version information, continue anyway?", vbYesNo + vbQuestion)
If nYesNo = vbYes Then CheckVersion = True
End Function

Private Sub OpenTables()
On Error GoTo error:

Set tabRooms = DB.OpenRecordset("Rooms")
Set tabItems = DB.OpenRecordset("Items")
Set tabClasses = DB.OpenRecordset("Classes")
Set tabRaces = DB.OpenRecordset("Races")
Set tabSpells = DB.OpenRecordset("Spells")
Set tabActions = DB.OpenRecordset("Actions")
Set tabMonsters = DB.OpenRecordset("Monsters")
Set tabShops = DB.OpenRecordset("Shops")
Set tabMessages = DB.OpenRecordset("Messages")
Set tabTextblocks = DB.OpenRecordset("Textblocks")
Set tabInfo = DB.OpenRecordset("Info")

Exit Sub
error:
Call HandleError("OpenTables")
Resume Next

End Sub
Private Sub CloseAll(Optional DontCompact As Boolean)
On Error Resume Next
Dim temp As String
Dim fso As FileSystemObject

tabRooms.Close
tabItems.Close
tabSpells.Close
tabRaces.Close
tabClasses.Close
tabInfo.Close
tabMonsters.Close
tabShops.Close
tabMessages.Close
tabTextblocks.Close
tabActions.Close

DB.Close

Set tabRooms = Nothing
Set tabMonsters = Nothing
Set tabShops = Nothing
Set tabItems = Nothing
Set tabSpells = Nothing
Set tabRaces = Nothing
Set tabClasses = Nothing
Set tabInfo = Nothing
Set tabMessages = Nothing
Set tabTextblocks = Nothing
Set tabActions = Nothing

Set DB = Nothing

If DontCompact Then GoTo finish:

stsStatusBar.Panels(1).Text = ""
stsStatusBar.Panels(2).Text = "Compacting Database ..."

On Error GoTo NoCompact:
Set fso = CreateObject("Scripting.FileSystemObject")

temp = sDataSource & "_temp.mdb"
Call CompactDatabase(sDataSource, temp)

If fso.FileExists(temp) Then
    fso.DeleteFile sDataSource
    fso.CopyFile temp, sDataSource
    fso.DeleteFile temp, True
End If

GoTo finish:

NoCompact:
On Error Resume Next
stsStatusBar.Panels(2).Text = ""

finish:
Set fso = Nothing
DoEvents

End Sub


Private Sub cmdCancel_Click()
Dim nYesNo As Integer

If cmdGo.Enabled = False Then
    nYesNo = MsgBox("Are you sure you want to cancel?", vbYesNo + vbQuestion + vbDefaultButton2)
    If Not nYesNo = vbYes Then Exit Sub

    cmdCancel.Enabled = False
    bStopExport = True
    DoEvents
Else
    Unload Me
End If

End Sub

Private Sub CreateExportFile(fil As String)
Dim fso As FileSystemObject

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(fil) = True Then fso.DeleteFile fil, True

fso.CreateTextFile (fil)

Set fso = Nothing

End Sub
Private Sub ExportVersionInfo()
On Error GoTo error:

tabInfo.AddNew
tabInfo.Fields("NMR Version") = sAppVersion
tabInfo.Fields("Dat File Version") = FriendlyDatVersion(eDatFileVersion)
tabInfo.Fields("Date") = Date
tabInfo.Fields("Time") = Time
tabInfo.Fields("Custom") = txtCustom.Text
tabInfo.Update

Exit Sub
error:
Call HandleError("ExportVersionInfo")
End Sub
Private Sub ExportBankbooks()
Dim nStatus As Integer, recnum As Long
Dim fso As FileSystemObject, ts As TextStream

Set fso = CreateObject("Scripting.FileSystemObject")
Set ts = fso.OpenTextFile(BankbooksTextfile, ForWriting)

recnum = 1
stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_BANKS
stsStatusBar.Panels(2).Text = recnum

nStatus = BTRCALL(BGETFIRST, BankPosBlock, BankDatabuf, Len(BankDatabuf), ByVal BankKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Bankbooks: Could not get first record, Error: " & BtrieveErrorCode(nStatus)
    Set ts = Nothing
    Set fso = Nothing
    Exit Sub
End If

ts.WriteLine ("BBS Name" & vbTab & "Shop Number" & vbTab & "Amount")

Do While nStatus = 0 And Not bStopExport

    Call BankRowToStruct(BankDatabuf.buf)
    
    ts.Write (RTrim(RemoveCharacter(Bankrec.BBSName, vbNull)) & vbTab)
    ts.Write (Bankrec.ShopNumber & vbTab)
    ts.Write (SLong2ULong(Bankrec.Cash) & vbTab)
    
    ts.WriteLine ("")
    
    nStatus = BTRCALL(BGETNEXT, BankPosBlock, BankDatabuf, Len(BankDatabuf), ByVal BankKeyBuffer, KEY_BUF_LEN, 0)
    
    recnum = recnum + 1
    stsStatusBar.Panels(2).Text = recnum
    IncreaseProgressBar
    If Not bUseCPU Then DoEvents

Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error exporting Banks, Btrieve Error: " & BtrieveErrorCode(nStatus, True)
End If

ts.Close
Set fso = Nothing
Set ts = Nothing

End Sub

Private Sub ExportTextblocks(format As String)
Dim nStatus As Integer, decrypted As String, nLastRec(1) As Long
Dim fso As FileSystemObject, ts As TextStream
Dim nRecnum As Long, nLastRecNum As Long
Dim nListNum As Integer, nCurrenListItem As Long

nListNum = 7

If chkExportAll(nListNum).Value = 0 Then
    If lvList(nListNum).ListItems.Count = 0 Then Exit Sub
    nCurrenListItem = 1
    nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
    nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
    
GotoNextTextblockStart:
    TextblockKey.PartNum = 0
    TextblockKey.Number = nRecnum
    nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKeyStructToRow(), KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nRecnum = nLastRecNum Then
            If nCurrenListItem = lvList(nListNum).ListItems.Count Then
                MsgBox "No textblock actually found to export.", vbInformation
                Exit Sub
            End If
            nCurrenListItem = nCurrenListItem + 1
            
            nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
            nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
        Else
            nRecnum = nRecnum + 1
        End If
        GoTo GotoNextTextblockStart:
    End If
Else
    nRecnum = 1
    nStatus = BTRCALL(BGETFIRST, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Textblocks: Could not get first record, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
End If

Set fso = CreateObject("Scripting.FileSystemObject")

stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_TEXT
stsStatusBar.Panels(2).Text = nRecnum

If format = "Access" Then GoTo Access:

Set ts = fso.OpenTextFile(TextblocksTextfile, ForWriting)
ts.WriteLine ("Number" & vbTab & "Part#" & vbTab & "LinkTo" & vbTab & "Data")

Do While nStatus = 0 And Not bStopExport
    
    decrypted = ""
    TextblockRowToStruct TextblockDataBuf.buf

    ts.Write (TextblockRec.Number & vbTab)
    ts.Write (TextblockRec.PartNum & vbTab)
    ts.Write (TextblockRec.LinkTo & vbTab)
    
    decrypted = DecryptTextblock(TextblockRec.Data)
    
    ts.WriteLine ("[TBLOCK]" & decrypted & "[/TBLOCK]")
    
    If chkExportAll(nListNum).Value = 0 Then
GotoNextTextblock:
        Call IncreaseProgressBar
        
        nStatus = BTRCALL(BGETNEXT, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            TextblockRowToStruct TextblockDataBuf.buf
            
            If TextblockRec.Number > nRecnum Then
                'next record is not a new partnum
                If TextblockRec.Number > nLastRecNum Then
                    'now out of range
                    If nCurrenListItem = lvList(nListNum).ListItems.Count Then GoTo finish
                    nCurrenListItem = nCurrenListItem + 1
                        
                    nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
                    nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
                    
                    TextblockKey.PartNum = 0
                    TextblockKey.Number = nRecnum
                    
                    nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKeyStructToRow(), KEY_BUF_LEN, 0)
                    If Not nStatus = 0 Then GoTo GotoNextTextblock:
                Else
                    nRecnum = nRecnum + 1
                End If
            End If
        End If
    Else
        nStatus = BTRCALL(BGETNEXT, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
        nRecnum = nRecnum + 1
        Call IncreaseProgressBar
    End If
    
    stsStatusBar.Panels(2).Text = nRecnum
    If Not bUseCPU Then DoEvents
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error exporting Textblocks, Btrieve Error: " & BtrieveErrorCode(nStatus, True)
End If

finish:
ts.Close

Set ts = Nothing
Set fso = Nothing
Exit Sub

Access:
tabTextblocks.Index = "idxTextblocks"
nLastRec(0) = TextblockRec.Number
nLastRec(1) = TextblockRec.PartNum
Do While nStatus = 0 And Not bStopExport
    RowToStruct TextblockDataBuf.buf, TextblockFldMap, TextblockRec, LenB(TextblockRec)
    
    If bUpdateExistingADB Then
        'check for extra textblock parts
        If nLastRec(0) <> TextblockRec.Number Then
            TextblockKey.Number = nLastRec(0)
            TextblockKey.PartNum = nLastRec(1)
part_check:
            TextblockKey.PartNum = TextblockKey.PartNum + 1
            
            tabTextblocks.Seek "=", TextblockKey.Number, TextblockKey.PartNum
            If tabTextblocks.NoMatch = False Then
                tabTextblocks.Delete
                GoTo part_check:
            End If
        End If
    End If
    nLastRec(0) = TextblockRec.Number
    nLastRec(1) = TextblockRec.PartNum
    
    stsStatusBar.Panels(2).Text = nRecnum
    
    If bUpdateExistingADB = True Then
        If tabTextblocks.RecordCount = 0 Then
            tabTextblocks.AddNew
        Else
            tabTextblocks.Seek "=", TextblockRec.Number, TextblockRec.PartNum
            If tabTextblocks.NoMatch = True Then
                tabTextblocks.AddNew
            Else
                tabTextblocks.Edit
            End If
        End If
    Else
        tabTextblocks.AddNew
    End If
    
    tabTextblocks.Fields("Number") = TextblockRec.Number
    tabTextblocks.Fields("Part #") = TextblockRec.PartNum
    tabTextblocks.Fields("Link To") = TextblockRec.LinkTo
    
    decrypted = DecryptTextblock(TextblockRec.Data)
    
    tabTextblocks.Fields("Data") = decrypted
    
'    For x = 1 To 8
'        If Len(decrypted) <= 250 Then
'            tabTextblocks.Fields(CStr("Data Part " & x)) = decrypted
'            decrypted = ""
'        Else
'            tabTextblocks.Fields(CStr("Data Part " & x)) = Left(decrypted, 250)
'            decrypted = Right(decrypted, Len(decrypted) - 250)
'        End If
'    Next
    
    tabTextblocks.Update
    
    If chkExportAll(nListNum).Value = 0 Then
        nStatus = BTRCALL(BGETNEXT, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            TextblockRowToStruct TextblockDataBuf.buf
            
            If TextblockRec.Number > nRecnum Then
                'next record is not a new partnum
                If TextblockRec.Number > nLastRecNum Then
                    'now out of range
GotoNextTextblock_access:
                    Call IncreaseProgressBar
                    
                    If nCurrenListItem = lvList(nListNum).ListItems.Count Then GoTo FinishedAccess
                    nCurrenListItem = nCurrenListItem + 1
                        
                    nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
                    nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
                    
                    TextblockKey.PartNum = 0
                    TextblockKey.Number = nRecnum
                    
                    nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKeyStructToRow(), KEY_BUF_LEN, 0)
                    If Not nStatus = 0 Then GoTo GotoNextTextblock_access:
                Else
                    nRecnum = nRecnum + 1
                    Call IncreaseProgressBar
                End If
            Else
                'new part
            End If
        End If
    Else
        nStatus = BTRCALL(BGETNEXT, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0)
        nRecnum = nRecnum + 1
        Call IncreaseProgressBar
    End If
    
    If Not bUseCPU Then DoEvents
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error exporting Textblocks, Btrieve Error: " & BtrieveErrorCode(nStatus, True)
End If

FinishedAccess:

Set fso = Nothing
Set ts = Nothing

End Sub

Private Sub ExportMessages(format As String)
Dim nStatus As Integer
Dim fso As FileSystemObject, ts As TextStream, x As Long
Dim nRecnum As Long, nLastRecNum As Long
Dim nListNum As Integer, nCurrenListItem As Long

nListNum = 3

If chkExportAll(nListNum).Value = 0 Then
    If lvList(nListNum).ListItems.Count = 0 Then Exit Sub
    nCurrenListItem = 1
    nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
    nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
    
GotoNextMessageStart:
    nStatus = BTRCALL(BGETEQUAL, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), nRecnum, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nRecnum = nLastRecNum Then
            If nCurrenListItem = lvList(nListNum).ListItems.Count Then
                MsgBox "No message found to export.", vbInformation
                Exit Sub
            End If
            nCurrenListItem = nCurrenListItem + 1
            
            nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
            nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
        Else
            nRecnum = nRecnum + 1
        End If
        GoTo GotoNextMessageStart:
    End If
Else
    nRecnum = 1
    nStatus = BTRCALL(BGETFIRST, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Messages: Could not get first record, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
End If
    
Set fso = CreateObject("Scripting.FileSystemObject")

stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_MSG
stsStatusBar.Panels(2).Text = nRecnum

If format = "Access" Then GoTo Access:
    
Set ts = fso.OpenTextFile(MessagesTextfile, ForWriting)
ts.WriteLine ("Number" & vbTab & "Line1" & vbTab & "Line2" & vbTab & "Line3")

Do While nStatus = 0 And Not bStopExport
    RowToStruct Messagedatabuf.buf, MessageFldMap, Messagerec, LenB(Messagerec)
    
    ts.Write (Messagerec.Number & vbTab)
    ts.Write (RTrim(Messagerec.MessageLine1) & vbTab)
    ts.Write (RTrim(Messagerec.MessageLine2) & vbTab)
    ts.WriteLine (RTrim(Messagerec.MessageLine3))
    
    If chkExportAll(nListNum).Value = 0 Then
GotoNextMessage:
        Call IncreaseProgressBar
        
        nStatus = BTRCALL(BGETNEXT, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            MessageRowToStruct Messagedatabuf.buf
            
            If Messagerec.Number > nLastRecNum Then
                If nCurrenListItem = lvList(nListNum).ListItems.Count Then GoTo Finished
                nCurrenListItem = nCurrenListItem + 1
                    
                nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
                nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
                
                nStatus = BTRCALL(BGETEQUAL, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), nRecnum, KEY_BUF_LEN, 0)
                If Not nStatus = 0 Then GoTo GotoNextMessage:
            Else
                nRecnum = nRecnum + 1
            End If
        End If
    Else
        nStatus = BTRCALL(BGETNEXT, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
        nRecnum = nRecnum + 1
        Call IncreaseProgressBar
    End If
    
    stsStatusBar.Panels(2).Text = nRecnum
    If Not bUseCPU Then DoEvents
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error exporting Messages, Btrieve Error: " & BtrieveErrorCode(nStatus, True)
End If

Finished:
ts.Close
Set fso = Nothing
Set ts = Nothing

Exit Sub

Access:
'Dim adoConnect As Database
'Dim tabMessages As Recordset

'Set adoConnect = OpenDatabase(sDataSource)
'Set tabMessages = adoConnect.OpenRecordset("Messages")

tabMessages.Index = "pkMessages"
Do While nStatus = 0 And Not bStopExport
    
    RowToStruct Messagedatabuf.buf, MessageFldMap, Messagerec, LenB(Messagerec)
    
    If bUpdateExistingADB = True Then
        If tabMessages.RecordCount = 0 Then
            tabMessages.AddNew
        Else
            tabMessages.Seek "=", Messagerec.Number
            If tabMessages.NoMatch = True Then
                tabMessages.AddNew
            Else
                tabMessages.Edit
            End If
        End If
    Else
        tabMessages.AddNew
    End If
    
    tabMessages.Fields("Number") = Messagerec.Number
    tabMessages.Fields("Line 1") = Messagerec.MessageLine1
    tabMessages.Fields("Line 2") = Messagerec.MessageLine2
    tabMessages.Fields("Line 3") = Messagerec.MessageLine3
    
    tabMessages.Update
   
    If chkExportAll(nListNum).Value = 0 Then
GotoNextMessageAccess:
        Call IncreaseProgressBar
        
        nStatus = BTRCALL(BGETNEXT, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            MessageRowToStruct Messagedatabuf.buf
            
            If Messagerec.Number > nLastRecNum Then
                If nCurrenListItem = lvList(nListNum).ListItems.Count Then GoTo FinishedAccess
                nCurrenListItem = nCurrenListItem + 1
                    
                nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
                nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
                
                nStatus = BTRCALL(BGETEQUAL, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), nRecnum, KEY_BUF_LEN, 0)
                If Not nStatus = 0 Then GoTo GotoNextMessageAccess:
            Else
                nRecnum = nRecnum + 1
            End If
        End If
    Else
        nStatus = BTRCALL(BGETNEXT, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
        nRecnum = nRecnum + 1
        Call IncreaseProgressBar
    End If
    
    stsStatusBar.Panels(2).Text = nRecnum
    If Not bUseCPU Then DoEvents
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error exporting Messages, Btrieve Error: " & BtrieveErrorCode(nStatus, True)
End If

FinishedAccess:
Set fso = Nothing
Set ts = Nothing

End Sub

Private Sub ExportItems(format As String)
Dim nStatus As Integer
Dim fso As FileSystemObject, ts As TextStream, x As Long
Dim nRecnum As Long, nLastRecNum As Long
Dim nListNum As Integer, nCurrenListItem As Long

nListNum = 2

If chkExportAll(nListNum).Value = 0 Then
    If lvList(nListNum).ListItems.Count = 0 Then Exit Sub
    nCurrenListItem = 1
    nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
    nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
    
GotoNextItemStart:
    nStatus = BTRCALL(BGETEQUAL, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), nRecnum, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nRecnum = nLastRecNum Then
            If nCurrenListItem = lvList(nListNum).ListItems.Count Then
                MsgBox "No item found to export.", vbInformation
                Exit Sub
            End If
            nCurrenListItem = nCurrenListItem + 1
            
            nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
            nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
        Else
            nRecnum = nRecnum + 1
        End If
        GoTo GotoNextItemStart:
    End If
Else
    nRecnum = 1
    nStatus = BTRCALL(BGETFIRST, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Items: Could not get first record, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
End If
    
Set fso = CreateObject("Scripting.FileSystemObject")

stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_ITEMS
stsStatusBar.Panels(2).Text = nRecnum

If format = "Access" Then GoTo Access:

Set ts = fso.OpenTextFile(ItemsTextfile, ForWriting)
ts.Write ("Number" & vbTab & "Name" & vbTab & "Desc 1" & vbTab & "Desc 2" & vbTab & "Desc 3" & vbTab & "Desc 4" & vbTab & "Desc 5" & vbTab & "Desc 6" & vbTab & "Game Limit" & vbTab & "Weight" & vbTab & "Type" & vbTab & "Uses" & vbTab & "Cost" & vbTab & "Cost Type" & vbTab & "Min Hit" & vbTab & "Max Hit" & vbTab & "Weapon" & vbTab & "Armour" & vbTab & "WornOn" & vbTab & "Accuracy" & vbTab & "DR" & vbTab & "AC" & vbTab)
ts.Write ("Gettable" & vbTab & "Required Strength" & vbTab & "Speed" & vbTab & "Robable" & vbTab & "Miss MSG" & vbTab & "Hit MSG" & vbTab & "Destruct MSG" & vbTab & "Read MSG" & vbTab & "Not Droppable" & vbTab & "Destroy on Death" & vbTab & "RetainAfterUser" & vbTab)
ts.Write ("OpenRunic" & vbTab & "OpenPlat" & vbTab & "OpenGold" & vbTab & "OpenSilver" & vbTab & "OpenCopper" & vbTab)

For x = 0 To 9
    ts.Write ("Class " & x & vbTab)
Next
For x = 0 To 9
    ts.Write ("Race " & x & vbTab)
Next
For x = 0 To 9
    ts.Write ("Negate " & x & vbTab)
Next
For x = 0 To 19
    ts.Write ("Ability " & x & vbTab)
    ts.Write ("AbilVal " & x & vbTab)
Next
ts.WriteLine ("")

Do While nStatus = 0 And Not bStopExport
    RowToStruct Itemdatabuf.buf, ItemFldMap, Itemrec, LenB(Itemrec)
    
    ts.Write (Itemrec.Number & vbTab)
    ts.Write (RTrim(RemoveCharacter(Itemrec.Name, vbNull)) & vbTab)
    ts.Write (RTrim(RemoveCharacter(Itemrec.Desc1, vbNull)) & vbTab)
    ts.Write (RTrim(RemoveCharacter(Itemrec.Desc2, vbNull)) & vbTab)
    ts.Write (RTrim(RemoveCharacter(Itemrec.Desc3, vbNull)) & vbTab)
    ts.Write (RTrim(RemoveCharacter(Itemrec.Desc4, vbNull)) & vbTab)
    ts.Write (RTrim(RemoveCharacter(Itemrec.Desc5, vbNull)) & vbTab)
    ts.Write (RTrim(RemoveCharacter(Itemrec.Desc6, vbNull)) & vbTab)
    ts.Write (Itemrec.GameLimit & vbTab)
    ts.Write (Itemrec.Weight & vbTab)
    ts.Write (Itemrec.Type & vbTab)
    ts.Write (Itemrec.Uses & vbTab)
    ts.Write (Itemrec.Cost & vbTab)
    ts.Write (Itemrec.CostType & vbTab)
    ts.Write (Itemrec.Minhit & vbTab)
    ts.Write (Itemrec.Maxhit & vbTab)
    ts.Write (Itemrec.Weapon & vbTab)
    ts.Write (Itemrec.Armour & vbTab)
    ts.Write (Itemrec.WornOn & vbTab)
    ts.Write (Itemrec.Accuracy & vbTab)
    ts.Write (Itemrec.AC & vbTab)
    ts.Write (Itemrec.DR & vbTab)
    ts.Write (Itemrec.Gettable & vbTab)
    ts.Write (Itemrec.ReqStr & vbTab)
    ts.Write (Itemrec.Speed & vbTab)
    ts.Write (Itemrec.Robable & vbTab)
    ts.Write (Itemrec.MissMsg & vbTab)
    ts.Write (Itemrec.HitMsg & vbTab)
    ts.Write (Itemrec.DistructMsg & vbTab)
    ts.Write (Itemrec.ReadTB & vbTab)
    ts.Write (Itemrec.NotDroppable & vbTab)
    ts.Write (Itemrec.DestroyOnDeath & vbTab)
    ts.Write (Itemrec.RetainAfterUses & vbTab)
    ts.Write (Itemrec.OpenRunic & vbTab)
    ts.Write (Itemrec.OpenPlatinum & vbTab)
    ts.Write (Itemrec.OpenGold & vbTab)
    ts.Write (Itemrec.OpenSilver & vbTab)
    ts.Write (Itemrec.OpenCopper & vbTab)
    
    For x = 0 To 9
        ts.Write (Itemrec.Class(x) & vbTab)
    Next
    
    For x = 0 To 9
        ts.Write (Itemrec.Race(x) & vbTab)
    Next
    
    For x = 0 To 9
        ts.Write (Itemrec.Negate(x * 2) & vbTab)
    Next
    
    For x = 0 To 19
        ts.Write (Itemrec.AbilityA(x) & vbTab)
        ts.Write (Itemrec.AbilityB(x) & vbTab)
    Next

    ts.WriteLine ("")
    If chkExportAll(nListNum).Value = 0 Then
GotoNextItem:
        Call IncreaseProgressBar
        
        nStatus = BTRCALL(BGETNEXT, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            ItemRowToStruct Itemdatabuf.buf
            
            If Itemrec.Number > nLastRecNum Then
                If nCurrenListItem = lvList(nListNum).ListItems.Count Then GoTo Finished
                nCurrenListItem = nCurrenListItem + 1
                    
                nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
                nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
                
                nStatus = BTRCALL(BGETEQUAL, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), nRecnum, KEY_BUF_LEN, 0)
                If Not nStatus = 0 Then GoTo GotoNextItem:
            Else
                nRecnum = nRecnum + 1
            End If
        End If
    Else
        nStatus = BTRCALL(BGETNEXT, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
        nRecnum = nRecnum + 1
        Call IncreaseProgressBar
    End If
    
    stsStatusBar.Panels(2).Text = nRecnum
    If Not bUseCPU Then DoEvents
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error exporting Items, Btrieve Error: " & BtrieveErrorCode(nStatus, True)
End If

Finished:
ts.Close
Set fso = Nothing
Set ts = Nothing

Exit Sub

Access:
'Dim adoConnect As Database
'Dim tabItems As Recordset
'
'Set adoConnect = OpenDatabase(sDataSource)
'Set tabItems = adoConnect.OpenRecordset("Items")

tabItems.Index = "pkItems"
Do While nStatus = 0 And Not bStopExport
    
    RowToStruct Itemdatabuf.buf, ItemFldMap, Itemrec, LenB(Itemrec)
    
    If bUpdateExistingADB = True Then
        If tabItems.RecordCount = 0 Then
            tabItems.AddNew
        Else
            tabItems.Seek "=", Itemrec.Number
            If tabItems.NoMatch = True Then
                tabItems.AddNew
            Else
                tabItems.Edit
            End If
        End If
    Else
        tabItems.AddNew
    End If
    
    tabItems.Fields("Number") = Itemrec.Number
    tabItems.Fields("Name") = Itemrec.Name
    tabItems.Fields("Game Limit") = Itemrec.GameLimit
    tabItems.Fields("Desc1") = Itemrec.Desc1
    tabItems.Fields("Desc2") = Itemrec.Desc2
    tabItems.Fields("Desc3") = Itemrec.Desc3
    tabItems.Fields("Desc4") = Itemrec.Desc4
    tabItems.Fields("Desc5") = Itemrec.Desc5
    tabItems.Fields("Desc6") = Itemrec.Desc6
    tabItems.Fields("Weight") = Itemrec.Weight
    tabItems.Fields("Type") = Itemrec.Type
    tabItems.Fields("Uses") = Itemrec.Uses
    tabItems.Fields("Cost") = Itemrec.Cost
    tabItems.Fields("Cost Type") = Itemrec.CostType
    tabItems.Fields("Min Hit") = Itemrec.Minhit
    tabItems.Fields("Max Hit") = Itemrec.Maxhit
    tabItems.Fields("AC") = Itemrec.AC
    tabItems.Fields("DR") = Itemrec.DR
    tabItems.Fields("Weapon") = Itemrec.Weapon
    tabItems.Fields("Armour") = Itemrec.Armour
    tabItems.Fields("Worn On") = Itemrec.WornOn
    tabItems.Fields("Accuracy") = Itemrec.Accuracy
    tabItems.Fields("Gettable") = Itemrec.Gettable
    tabItems.Fields("Req Str") = Itemrec.ReqStr
    tabItems.Fields("Speed") = Itemrec.Speed
    tabItems.Fields("Robable") = Itemrec.Robable
    tabItems.Fields("Hit Msg") = Itemrec.HitMsg
    tabItems.Fields("Miss Msg") = Itemrec.MissMsg
    tabItems.Fields("Read Msg") = Itemrec.ReadTB
    tabItems.Fields("Distruct Msg") = Itemrec.DistructMsg
    tabItems.Fields("Not Droppable") = Itemrec.NotDroppable
    tabItems.Fields("Destroy On Death") = Itemrec.DestroyOnDeath
    tabItems.Fields("Retain After Uses") = Itemrec.RetainAfterUses
    tabItems.Fields("OpenRunic") = Itemrec.OpenRunic
    tabItems.Fields("OpenPlatinum") = Itemrec.OpenPlatinum
    tabItems.Fields("OpenGold") = Itemrec.OpenGold
    tabItems.Fields("OpenSilver") = Itemrec.OpenSilver
    tabItems.Fields("OpenCopper") = Itemrec.OpenCopper
    
    For x = 0 To 9
        tabItems.Fields("Class " & x) = Itemrec.Class(x)
    Next
    
    For x = 0 To 9
        tabItems.Fields("Negate " & x) = Itemrec.Negate(x * 2)
    Next
    
    For x = 0 To 9
        tabItems.Fields("Race " & x) = Itemrec.Race(x)
    Next

    For x = 0 To 19
        tabItems.Fields("Ability " & x) = Itemrec.AbilityA(x)
        tabItems.Fields("Ability Value " & x) = Itemrec.AbilityB(x)
    Next
    
    tabItems.Update
    
    If chkExportAll(nListNum).Value = 0 Then
GotoNextItemAccess:
        Call IncreaseProgressBar
        
        nStatus = BTRCALL(BGETNEXT, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            ItemRowToStruct Itemdatabuf.buf
            
            If Itemrec.Number > nLastRecNum Then
                If nCurrenListItem = lvList(nListNum).ListItems.Count Then GoTo FinishedAccess
                nCurrenListItem = nCurrenListItem + 1
                    
                nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
                nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
                
                nStatus = BTRCALL(BGETEQUAL, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), nRecnum, KEY_BUF_LEN, 0)
                If Not nStatus = 0 Then GoTo GotoNextItemAccess:
            Else
                nRecnum = nRecnum + 1
            End If
        End If
    Else
        nStatus = BTRCALL(BGETNEXT, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
        nRecnum = nRecnum + 1
        Call IncreaseProgressBar
    End If
    
    stsStatusBar.Panels(2).Text = nRecnum
    If Not bUseCPU Then DoEvents
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error exporting Items, Btrieve Error: " & BtrieveErrorCode(nStatus, True)
End If

FinishedAccess:
Set fso = Nothing
Set ts = Nothing

End Sub
Private Sub ExportRooms(format As String)
Dim nStatus As Integer, x As Integer
Dim fso As FileSystemObject, ts As TextStream
Dim nRecnum As Long, nLastRecNum As Long, nMap As Long
Dim nListNum As Integer, nCurrenListItem As Long

nListNum = 8

If chkExportAll(nListNum).Value = 0 Then
    If lvList(nListNum).ListItems.Count = 0 Then Exit Sub
    nCurrenListItem = 1
    nMap = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
    nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
    nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(2).Text)
    
GotoNextRoomStart:
    RoomKeyStruct.MapNum = nMap
    RoomKeyStruct.RoomNum = nRecnum
    nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), RoomKeyStruct, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nRecnum = nLastRecNum Then
            If nCurrenListItem = lvList(nListNum).ListItems.Count Then
                MsgBox "No rooms actually found to export.", vbInformation
                Exit Sub
            End If
            nCurrenListItem = nCurrenListItem + 1
            nMap = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
            nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
            nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(2).Text)
        Else
            nRecnum = nRecnum + 1
        End If
        GoTo GotoNextRoomStart:
    End If
Else
    nRecnum = 1
    nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Rooms: Could not get first record, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
End If
    
Set fso = CreateObject("Scripting.FileSystemObject")

stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_MP
stsStatusBar.Panels(2).Text = nRecnum

If format = "Access" Then GoTo Access:

Set ts = fso.OpenTextFile(RoomsTextfile, ForWriting)

ts.Write ("Map" & vbTab & "Room" & vbTab & "Name" & vbTab)
For x = 0 To 6
    ts.Write ("Desc " & x & vbTab)
Next
ts.Write ("AnsiMap" & vbTab & "Type" & vbTab & "Shop#" & vbTab & "Gang House #" & vbTab & "Min Index" & vbTab & "Max Index" & vbTab & "Light" & vbTab & "Runic" & vbTab & "Platinum" & vbTab & "Gold" & vbTab & "Silver" & vbTab & "Copper" & vbTab & "InvisRunic" & vbTab & "InvisPlatinum" & vbTab & "InvisGold" & vbTab & "InvisSilver" & vbTab & "InvisCopper" & vbTab & "Max Regen" & vbTab)
ts.Write ("Mon Type" & vbTab & "Attributes" & vbTab & "Death Room" & vbTab & "Exit Room" & vbTab & "Command Text" & vbTab & "Delay" & vbTab & "Max Area" & vbTab & "Control Room" & vbTab & "Perm NPC" & vbTab & "Spell" & vbTab)
For x = 0 To 9
    ts.Write ("Exit " & x & vbTab)
    ts.Write ("Type " & x & vbTab)
    ts.Write ("Para1 " & x & vbTab)
    ts.Write ("Para2 " & x & vbTab)
    ts.Write ("Para3 " & x & vbTab)
    ts.Write ("Para4 " & x & vbTab)
Next
For x = 0 To 16
    ts.Write ("RoomItem " & x & vbTab)
    ts.Write ("RoomItem " & x & " USES" & vbTab)
    ts.Write ("RoomItem " & x & " QTY" & vbTab)
Next
For x = 0 To 14
    ts.Write ("HiddenItem " & x & vbTab)
    ts.Write ("HiddenItem " & x & " USES" & vbTab)
    ts.Write ("HiddenItem " & x & " QTY" & vbTab)
Next
For x = 0 To 9
    ts.Write ("PlacedItem " & x & vbTab)
Next
'For x = 0 To 14
'    ts.Write ("CurMon " & x & vbTab)
'Next
ts.WriteLine ("")

Do While nStatus = 0 And Not bStopExport
    
    RowToStruct Roomdatabuf.buf, RoomFldMap, Roomrec, LenB(Roomrec)
    
    ts.Write (Roomrec.MapNumber & vbTab)
    ts.Write (Roomrec.RoomNumber & vbTab)
    ts.Write (RTrim(RemoveCharacter(Roomrec.Name, Chr(0))) & vbTab)

    For x = 0 To 6
        ts.Write (RTrim(RemoveCharacter(Roomrec.Desc(x), Chr(0))) & vbTab)
    Next

    ts.Write (Roomrec.AnsiMap & vbTab)
    ts.Write (Roomrec.Type & vbTab)
    ts.Write (Roomrec.ShopNum & vbTab)
    ts.Write (Roomrec.GangHouseNumber & vbTab)
    ts.Write (Roomrec.MinIndex & vbTab)
    ts.Write (Roomrec.MaxIndex & vbTab)
    ts.Write (Roomrec.Light & vbTab)
    ts.Write (Roomrec.Runic & vbTab)
    ts.Write (Roomrec.Platinum & vbTab)
    ts.Write (Roomrec.Gold & vbTab)
    ts.Write (Roomrec.Silver & vbTab)
    ts.Write (Roomrec.Copper & vbTab)
    ts.Write (Roomrec.InvisRunic & vbTab)
    ts.Write (Roomrec.InvisPlatinum & vbTab)
    ts.Write (Roomrec.InvisGold & vbTab)
    ts.Write (Roomrec.InvisSilver & vbTab)
    ts.Write (Roomrec.InvisCopper & vbTab)
    ts.Write (Roomrec.MaxRegen & vbTab)
    ts.Write (Roomrec.MonsterType & vbTab)
    ts.Write (Roomrec.Attributes & vbTab)
    ts.Write (Roomrec.DeathRoom & vbTab)
    ts.Write (Roomrec.ExitRoom & vbTab)
    ts.Write (Roomrec.CmdText & vbTab)
    ts.Write (Roomrec.Delay & vbTab)
    ts.Write (Roomrec.MaxArea & vbTab)
    ts.Write (Roomrec.ControlRoom & vbTab)
    ts.Write (Roomrec.PermNPC & vbTab)
    ts.Write (Roomrec.Spell & vbTab)

    For x = 0 To 9
        ts.Write (Roomrec.RoomExit(x) & vbTab)
        ts.Write (Roomrec.RoomType(x) & vbTab)
        ts.Write (Roomrec.Para1(x) & vbTab)
        ts.Write (Roomrec.Para2(x) & vbTab)
        ts.Write (Roomrec.Para3(x) & vbTab)
        ts.Write (Roomrec.Para4(x) & vbTab)
    Next

    For x = 0 To 16
        ts.Write (Roomrec.RoomItems(x) & vbTab)
        ts.Write (Roomrec.RoomItemUses(x) & vbTab)
        ts.Write (Roomrec.RoomItemQty(x) & vbTab)
    Next

    For x = 0 To 14
        ts.Write (Roomrec.InvisItems(x) & vbTab)
        ts.Write (Roomrec.InvisItemUses(x) & vbTab)
        ts.Write (Roomrec.InvisItemQty(x) & vbTab)
    Next

    For x = 0 To 9
        ts.Write (Roomrec.PlacedItems(x) & vbTab)
    Next

    'For x = 0 To 14
    '    ts.Write (Roomrec.CurrentRoomMon(x) & vbTab)
    'Next
    
    ts.WriteLine ("")

    If chkExportAll(nListNum).Value = 0 Then
GotoNextRoom:
        Call IncreaseProgressBar
        
        If nRecnum = nLastRecNum Then
            If nCurrenListItem = lvList(nListNum).ListItems.Count Then GoTo Finished
            nCurrenListItem = nCurrenListItem + 1
            nMap = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
            nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
            nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(2).Text)
        Else
            nRecnum = nRecnum + 1
        End If
        
        RoomKeyStruct.MapNum = nMap
        RoomKeyStruct.RoomNum = nRecnum
        
        nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), RoomKeyStruct, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then GoTo GotoNextRoom:
    Else
        nStatus = BTRCALL(BGETNEXT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
        nRecnum = nRecnum + 1
        IncreaseProgressBar
    End If

    stsStatusBar.Panels(2).Text = nRecnum
    If Not bUseCPU Then DoEvents
Loop

If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error exporting Rooms, Btrieve Error: " & BtrieveErrorCode(nStatus, True)
End If


Finished:
ts.Close
Set fso = Nothing
Set ts = Nothing

Exit Sub

Access:

tabRooms.Index = "idxRooms"
Do While nStatus = 0 And Not bStopExport
    
    RowToStruct Roomdatabuf.buf, RoomFldMap, Roomrec, LenB(Roomrec)
    
    If bUpdateExistingADB = True Then
        If tabRooms.RecordCount = 0 Then
            tabRooms.AddNew
        Else
            tabRooms.Seek "=", Roomrec.MapNumber, Roomrec.RoomNumber
            If tabRooms.NoMatch = True Then
                tabRooms.AddNew
            Else
                tabRooms.Edit
            End If
        End If
    Else
        tabRooms.AddNew
    End If
    
    tabRooms.Fields("Map Number") = Roomrec.MapNumber
    tabRooms.Fields("Room Number") = Roomrec.RoomNumber
    tabRooms.Fields("Name") = Roomrec.Name
    tabRooms.Fields("AnsiMap") = Roomrec.AnsiMap
    tabRooms.Fields("Type") = Roomrec.Type
    tabRooms.Fields("Shop Number") = Roomrec.ShopNum
    tabRooms.Fields("Gang House Number") = Roomrec.GangHouseNumber
    tabRooms.Fields("Min Index") = Roomrec.MinIndex
    tabRooms.Fields("Max Index") = Roomrec.MaxIndex
    tabRooms.Fields("Perm NPC") = Roomrec.PermNPC
    tabRooms.Fields("Light") = Roomrec.Light
    tabRooms.Fields("Mon Type") = Roomrec.MonsterType
    tabRooms.Fields("Max Regen") = Roomrec.MaxRegen
    tabRooms.Fields("Death Room") = Roomrec.DeathRoom
    tabRooms.Fields("Command Text") = Roomrec.CmdText
    tabRooms.Fields("Delay") = Roomrec.Delay
    tabRooms.Fields("Max Area") = Roomrec.MaxArea
    tabRooms.Fields("Control Room") = Roomrec.ControlRoom
    tabRooms.Fields("Spell") = Roomrec.Spell
    tabRooms.Fields("Exit Room") = Roomrec.ExitRoom
    tabRooms.Fields("Attributes") = Roomrec.Attributes
    For x = 0 To 6
        tabRooms.Fields("Desc " & x) = Roomrec.Desc(x)
    Next
    
    If chkZeroUserInteraction.Value = 1 Then
        tabRooms.Fields("Runic") = 0
        tabRooms.Fields("Platinum") = 0
        tabRooms.Fields("Gold") = 0
        tabRooms.Fields("Silver") = 0
        tabRooms.Fields("Copper") = 0
        tabRooms.Fields("InvisRunic") = 0
        tabRooms.Fields("InvisPlatinum") = 0
        tabRooms.Fields("InvisGold") = 0
        tabRooms.Fields("InvisSilver") = 0
        tabRooms.Fields("InvisCopper") = 0
        
        For x = 0 To 16
            tabRooms.Fields("Room Item " & x) = 0
            tabRooms.Fields("Room Item " & x & " QTY") = 0
            tabRooms.Fields("Room Item " & x & " USES") = 0
        Next
        For x = 0 To 14
            tabRooms.Fields("Hidden Item " & x) = 0
            tabRooms.Fields("Hidden Item " & x & " QTY") = 0
            tabRooms.Fields("Hidden Item " & x & " USES") = 0
            tabRooms.Fields("CurrentRoomMon " & x) = 0
        Next
    Else
        tabRooms.Fields("Runic") = Roomrec.Runic
        tabRooms.Fields("Platinum") = Roomrec.Platinum
        tabRooms.Fields("Gold") = Roomrec.Gold
        tabRooms.Fields("Silver") = Roomrec.Silver
        tabRooms.Fields("Copper") = Roomrec.Copper
        tabRooms.Fields("InvisRunic") = Roomrec.InvisRunic
        tabRooms.Fields("InvisPlatinum") = Roomrec.InvisPlatinum
        tabRooms.Fields("InvisGold") = Roomrec.InvisGold
        tabRooms.Fields("InvisSilver") = Roomrec.InvisSilver
        tabRooms.Fields("InvisCopper") = Roomrec.InvisCopper
        
        For x = 0 To 16
            tabRooms.Fields("Room Item " & x) = Roomrec.RoomItems(x)
            tabRooms.Fields("Room Item " & x & " QTY") = Roomrec.RoomItemQty(x)
            tabRooms.Fields("Room Item " & x & " USES") = Roomrec.RoomItemUses(x)
        Next
        For x = 0 To 14
            tabRooms.Fields("Hidden Item " & x) = Roomrec.InvisItems(x)
            tabRooms.Fields("Hidden Item " & x & " QTY") = Roomrec.InvisItemQty(x)
            tabRooms.Fields("Hidden Item " & x & " USES") = Roomrec.InvisItemUses(x)
            tabRooms.Fields("CurrentRoomMon " & x) = 0 'Roomrec.CurrentRoomMon(x)
        Next
    End If
    
    
    For x = 0 To 9
        tabRooms.Fields("Exit " & x) = Roomrec.RoomExit(x)
        tabRooms.Fields("Type " & x) = Roomrec.RoomType(x)
        tabRooms.Fields("Para1 " & x) = Roomrec.Para1(x)
        tabRooms.Fields("Para2 " & x) = Roomrec.Para2(x)
        tabRooms.Fields("Para3 " & x) = Roomrec.Para3(x)
        tabRooms.Fields("Para4 " & x) = Roomrec.Para4(x)
        tabRooms.Fields("Placed Item " & x) = Roomrec.PlacedItems(x)
    Next

    tabRooms.Update
        
    If chkExportAll(nListNum).Value = 0 Then
GotoNextRoomAccess:
        Call IncreaseProgressBar
        
        If nRecnum = nLastRecNum Then
            If nCurrenListItem = lvList(nListNum).ListItems.Count Then GoTo FinishedAccess:
            nCurrenListItem = nCurrenListItem + 1
            nMap = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
            nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
            nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(2).Text)
        Else
            nRecnum = nRecnum + 1
        End If
        
        RoomKeyStruct.MapNum = nMap
        RoomKeyStruct.RoomNum = nRecnum
        
        nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), RoomKeyStruct, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then GoTo GotoNextRoomAccess:
    Else
        nStatus = BTRCALL(BGETNEXT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
        nRecnum = nRecnum + 1
        Call IncreaseProgressBar
    End If

    stsStatusBar.Panels(2).Text = nRecnum
    If Not bUseCPU Then DoEvents
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error exporting Rooms, Btrieve Error: " & BtrieveErrorCode(nStatus, True)
End If

FinishedAccess:
Set fso = Nothing
Set ts = Nothing

End Sub
Private Sub ExportSpells(format As String)
Dim nStatus As Integer, x As Integer
Dim fso As FileSystemObject, ts As TextStream
Dim nRecnum As Long, nLastRecNum As Long
Dim nListNum As Integer, nCurrenListItem As Long

nListNum = 6

If chkExportAll(nListNum).Value = 0 Then
    If lvList(nListNum).ListItems.Count = 0 Then Exit Sub
    nCurrenListItem = 1
    nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
    nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
    
GotoNextSpellStart:
    nStatus = BTRCALL(BGETEQUAL, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), nRecnum, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nRecnum = nLastRecNum Then
            If nCurrenListItem = lvList(nListNum).ListItems.Count Then
                MsgBox "No spell found to export.", vbInformation
                Exit Sub
            End If
            nCurrenListItem = nCurrenListItem + 1
            
            nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
            nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
        Else
            nRecnum = nRecnum + 1
        End If
        GoTo GotoNextSpellStart:
    End If
Else
    nRecnum = 1
    nStatus = BTRCALL(BGETFIRST, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Spells: Could not get first record, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
End If
    
Set fso = CreateObject("Scripting.FileSystemObject")

stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_SPELS
stsStatusBar.Panels(2).Text = nRecnum

If format = "Access" Then GoTo Access:

Set ts = fso.OpenTextFile(SpellsTextfile, ForWriting)
ts.Write ("Number" & vbTab & "Name" & vbTab & "Short Name" & vbTab & "Desc A" & vbTab & "Desc B" & vbTab & "Magery Type" & vbTab & "Magery Level" & vbTab & "Mana" & vbTab & "Energy" & vbTab & "Level" & vbTab & "Min" & vbTab & "Max" & vbTab & "Msg Style" & vbTab & "Level Mod" & vbTab)
ts.Write ("Increase" & vbTab & "Level Cap" & vbTab & "Difficulty" & vbTab & "Length" & vbTab & "Type of Resist" & vbTab & "Spell Type" & vbTab & "Target" & vbTab & "Resist Ability" & vbTab & "Type of Attack" & vbTab & "Cast Msg A" & vbTab & "Cast Msg B" & vbTab)
For x = 0 To 9
    ts.Write ("Ability " & x & vbTab)
    ts.Write ("AbilVal " & x & vbTab)
Next
ts.WriteLine ("UNDEFINED01" & vbTab & "UNDEFINED03" & vbTab & "UNDEFINED04" & vbTab & "UNDEFINED05" & vbTab & "UNDEFINED06")

Do While nStatus = 0 And Not bStopExport
    RowToStruct Spelldatabuf.buf, SpellFldMap, Spellrec, LenB(Spellrec)
    
    ts.Write (Spellrec.Number & vbTab)
    ts.Write (RTrim(RemoveCharacter(Spellrec.Name, vbNull)) & vbTab)
    ts.Write (Spellrec.ShortName & vbTab)
    ts.Write (RTrim(RemoveCharacter(Spellrec.DescA, vbNull)) & vbTab)
    ts.Write (RTrim(RemoveCharacter(Spellrec.DescB, vbNull)) & vbTab)
    ts.Write (Spellrec.MageryA & vbTab)
    ts.Write (Spellrec.MageryB & vbTab)
    ts.Write (Spellrec.Mana & vbTab)
    ts.Write (Spellrec.Energy & vbTab)
    ts.Write (Spellrec.Level & vbTab)
    ts.Write (Spellrec.Min & vbTab)
    ts.Write (Spellrec.Max & vbTab)
    ts.Write (Spellrec.MsgStyle & vbTab)
    ts.Write (Spellrec.LVLSMaxIncr & vbTab)
    ts.Write (Spellrec.MaxIncrease & vbTab)
    ts.Write (Spellrec.LevelCap & vbTab)
    ts.Write (Spellrec.Difficulty & vbTab)
    ts.Write (Spellrec.duration & vbTab)
    ts.Write (Spellrec.TypeOfResists & vbTab)
    ts.Write (Spellrec.SpellType & vbTab)
    ts.Write (Spellrec.Target & vbTab)
    ts.Write (Spellrec.ResistAbility & vbTab)
    ts.Write (Spellrec.TypeOfAttack & vbTab)
    ts.Write (Spellrec.CastMsgA & vbTab)
    ts.Write (Spellrec.CastMsgB & vbTab)
    
    For x = 0 To 9
        ts.Write (Spellrec.AbilityA(x) & vbTab)
        ts.Write (Spellrec.AbilityB(x) & vbTab)
    Next
    
    ts.Write (Spellrec.UNDEFINED01 & vbTab)
    ts.Write (Spellrec.MinIncrease & vbTab)
    ts.Write (Spellrec.LVLSMinIncr & vbTab)
    ts.Write (Spellrec.LVLSDurIncr & vbTab)
    ts.WriteLine (Spellrec.DurIncrease)
    
    If chkExportAll(nListNum).Value = 0 Then
GotoNextSpell:
        Call IncreaseProgressBar
        
        nStatus = BTRCALL(BGETNEXT, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            SpellRowToStruct Spelldatabuf.buf
            
            If Spellrec.Number > nLastRecNum Then
                If nCurrenListItem = lvList(nListNum).ListItems.Count Then GoTo Finished
                nCurrenListItem = nCurrenListItem + 1
                    
                nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
                nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
                
                nStatus = BTRCALL(BGETEQUAL, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), nRecnum, KEY_BUF_LEN, 0)
                If Not nStatus = 0 Then GoTo GotoNextSpell:
            Else
                nRecnum = nRecnum + 1
            End If
        End If
    Else
        nStatus = BTRCALL(BGETNEXT, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
        nRecnum = nRecnum + 1
        Call IncreaseProgressBar
    End If
    
    stsStatusBar.Panels(2).Text = nRecnum
    If Not bUseCPU Then DoEvents
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error exporting Spells, Btrieve Error: " & BtrieveErrorCode(nStatus, True)
End If

Finished:

ts.Close
Set fso = Nothing
Set ts = Nothing

Exit Sub

Access:
'Dim adoConnect As Database
'Dim tabSpells As Recordset
'
'Set adoConnect = OpenDatabase(sDataSource)
'Set tabSpells = adoConnect.OpenRecordset("Spells")

tabSpells.Index = "pkSpells"
Do While nStatus = 0 And Not bStopExport
    
    RowToStruct Spelldatabuf.buf, SpellFldMap, Spellrec, LenB(Spellrec)
    
    If bUpdateExistingADB = True Then
        If tabSpells.RecordCount = 0 Then
            tabSpells.AddNew
        Else
            tabSpells.Seek "=", Spellrec.Number
            If tabSpells.NoMatch = True Then
                tabSpells.AddNew
            Else
                tabSpells.Edit
            End If
        End If
    Else
        tabSpells.AddNew
    End If
    
    tabSpells.Fields("Number") = Spellrec.Number
    tabSpells.Fields("Name") = Spellrec.Name
    tabSpells.Fields("Short Name") = Spellrec.ShortName
    tabSpells.Fields("Level") = Spellrec.Level
    tabSpells.Fields("Desc 1") = Spellrec.DescA
    tabSpells.Fields("Desc 2") = Spellrec.DescB
    tabSpells.Fields("Cast MSG A") = Spellrec.CastMsgA
    tabSpells.Fields("Cast MSG B") = Spellrec.CastMsgB
    tabSpells.Fields("MSG Style") = Spellrec.MsgStyle
    tabSpells.Fields("Energy") = Spellrec.Energy
    tabSpells.Fields("Mana") = Spellrec.Mana
    tabSpells.Fields("Min") = Spellrec.Min
    tabSpells.Fields("Max") = Spellrec.Max
    tabSpells.Fields("Spell Type") = Spellrec.SpellType
    tabSpells.Fields("Type of Resists") = Spellrec.TypeOfResists
    tabSpells.Fields("Difficulty") = Spellrec.Difficulty
    tabSpells.Fields("Target") = Spellrec.Target
    tabSpells.Fields("Duration") = Spellrec.duration
    tabSpells.Fields("Attack Type") = Spellrec.TypeOfAttack
    tabSpells.Fields("Resist Ability") = Spellrec.ResistAbility
    tabSpells.Fields("Magery A") = Spellrec.MageryA
    tabSpells.Fields("Magery B") = Spellrec.MageryB
    tabSpells.Fields("Level Cap") = Spellrec.LevelCap
    tabSpells.Fields("LVLS Max Increase") = Spellrec.LVLSMaxIncr
    tabSpells.Fields("Max Increase") = Spellrec.MaxIncrease
    tabSpells.Fields("LVLS Min Increase") = Spellrec.LVLSMinIncr
    tabSpells.Fields("Min Increase") = Spellrec.MinIncrease
    tabSpells.Fields("LVLS Dur Increase") = Spellrec.LVLSDurIncr
    tabSpells.Fields("Dur Increase") = Spellrec.DurIncrease
    tabSpells.Fields("UNDEFINED01") = Spellrec.UNDEFINED01
    tabSpells.Fields("UNDEFINED02") = Spellrec.UNDEFINED02
    
    For x = 0 To 9
        tabSpells.Fields("Ability " & x) = Spellrec.AbilityA(x)
        tabSpells.Fields("Ability Value " & x) = Spellrec.AbilityB(x)
    Next

    tabSpells.Update
    
    If chkExportAll(nListNum).Value = 0 Then
GotoNextSpellAccess:
        Call IncreaseProgressBar
        
        nStatus = BTRCALL(BGETNEXT, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            SpellRowToStruct Spelldatabuf.buf
            
            If Spellrec.Number > nLastRecNum Then
                If nCurrenListItem = lvList(nListNum).ListItems.Count Then GoTo FinishedAccess
                nCurrenListItem = nCurrenListItem + 1
                    
                nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
                nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
                
                nStatus = BTRCALL(BGETEQUAL, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), nRecnum, KEY_BUF_LEN, 0)
                If Not nStatus = 0 Then GoTo GotoNextSpellAccess:
            Else
                nRecnum = nRecnum + 1
            End If
        End If
    Else
        nStatus = BTRCALL(BGETNEXT, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
        nRecnum = nRecnum + 1
        Call IncreaseProgressBar
    End If
    
    stsStatusBar.Panels(2).Text = nRecnum
    If Not bUseCPU Then DoEvents
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error exporting Spells, Btrieve Error: " & BtrieveErrorCode(nStatus, True)
End If

FinishedAccess:
Set fso = Nothing
Set ts = Nothing

End Sub

Private Sub ExportActions(format As String)
Dim nStatus As Integer, recnum As Long
Dim fso As FileSystemObject, ts As TextStream

recnum = 1
stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_ACTS
stsStatusBar.Panels(2).Text = recnum

nStatus = BTRCALL(BGETFIRST, ActionPosBlock, ActionDatabuf, Len(ActionDatabuf), ByVal ActionKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Actions: Couldn't get first record, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
    
Set fso = CreateObject("Scripting.FileSystemObject")

If format = "Access" Then GoTo Access:

Set ts = fso.OpenTextFile(ActionsTextfile, ForWriting)
ts.Write ("Name" & vbTab & "SingleToUser" & vbTab & "SingleToRoom" & vbTab & "UserToUser" & vbTab & "UserToOtherUser" & vbTab & "UserToRoom" & vbTab & "MonsterToUser" & vbTab & "MonsterToRoom" & vbTab)
ts.WriteLine ("InventoryToUser" & vbTab & "InventoryToRoom" & vbTab & "FloorItemToUser" & vbTab & "FloorItemToRoom" & vbTab)

Do While nStatus = 0 And Not bStopExport
    RowToStruct ActionDatabuf.buf, ActionFldMap, Actionrec, LenB(Actionrec)
    
    ts.Write (RTrim(ClipNull(Actionrec.Name)) & vbTab)
    ts.Write (RTrim(ClipNull(Actionrec.SingleToUser)) & vbTab)
    ts.Write (RTrim(ClipNull(Actionrec.SingleToRoom)) & vbTab)
    ts.Write (RTrim(ClipNull(Actionrec.UserToUser)) & vbTab)
    ts.Write (RTrim(ClipNull(Actionrec.UserToOtherUser)) & vbTab)
    ts.Write (RTrim(ClipNull(Actionrec.UserToRoom)) & vbTab)
    ts.Write (RTrim(ClipNull(Actionrec.MonsterToUser)) & vbTab)
    ts.Write (RTrim(ClipNull(Actionrec.MonsterToRoom)) & vbTab)
    ts.Write (RTrim(ClipNull(Actionrec.InventoryToUser)) & vbTab)
    ts.Write (RTrim(ClipNull(Actionrec.InventoryToRoom)) & vbTab)
    ts.Write (RTrim(ClipNull(Actionrec.FloorItemToUser)) & vbTab)
    ts.WriteLine (RTrim(ClipNull(Actionrec.FloorItemToRoom)))
    
    nStatus = BTRCALL(BGETNEXT, ActionPosBlock, ActionDatabuf, Len(ActionDatabuf), ByVal ActionKeyBuffer, KEY_BUF_LEN, 0)
    
    recnum = recnum + 1
    stsStatusBar.Panels(2).Text = recnum
    IncreaseProgressBar
    If Not bUseCPU Then DoEvents
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error exporting Actions, Btrieve Error: " & BtrieveErrorCode(nStatus, True)
End If

ts.Close
Set fso = Nothing
Set ts = Nothing

Exit Sub

Access:
'Dim adoConnect As Database
'Dim tabActions As Recordset
'
'Set adoConnect = OpenDatabase(sDataSource)
'Set tabActions = adoConnect.OpenRecordset("Actions")
recnum = 0

tabActions.Index = "pkActions"
Do While nStatus = 0 And Not bStopExport
    
    RowToStruct ActionDatabuf.buf, ActionFldMap, Actionrec, LenB(Actionrec)
        
    recnum = recnum + 1
    stsStatusBar.Panels(2).Text = recnum
    IncreaseProgressBar
    
    If bUpdateExistingADB = True Then
        If tabActions.RecordCount = 0 Then
            tabActions.AddNew
        Else
            tabActions.Seek "=", Actionrec.Name
            If tabActions.NoMatch = True Then
                tabActions.Seek "=", RTrim(RemoveCharacter(Actionrec.Name, Chr(0)))
                If tabActions.NoMatch = True Then
                    tabActions.AddNew
                Else
                    tabActions.Edit
                End If
            Else
                tabActions.Edit
            End If
        End If
    Else
        tabActions.AddNew
    End If

    tabActions.Fields("Action") = Actionrec.Name
    tabActions.Fields("Single to User") = Actionrec.SingleToUser
    tabActions.Fields("Single to Room") = Actionrec.SingleToRoom
    tabActions.Fields("User to User") = Actionrec.UserToUser
    tabActions.Fields("User to Other User") = Actionrec.UserToOtherUser
    tabActions.Fields("User to Room") = Actionrec.UserToRoom
    tabActions.Fields("Monster to User") = Actionrec.MonsterToUser
    tabActions.Fields("Monster to Room") = Actionrec.MonsterToRoom
    tabActions.Fields("Inventory to User") = Actionrec.InventoryToUser
    tabActions.Fields("Inventory to Room") = Actionrec.InventoryToRoom
    tabActions.Fields("Floor Item to User") = Actionrec.FloorItemToUser
    tabActions.Fields("Floor Item to Room") = Actionrec.FloorItemToRoom

    tabActions.Update
    
    nStatus = BTRCALL(BGETNEXT, ActionPosBlock, ActionDatabuf, Len(ActionDatabuf), ByVal ActionKeyBuffer, KEY_BUF_LEN, 0)

    If Not bUseCPU Then DoEvents
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error exporting Actions, Btrieve Error: " & BtrieveErrorCode(nStatus, True)
End If

Set fso = Nothing
Set ts = Nothing

End Sub
Private Sub ExportClasses(format As String)
Dim nStatus As Integer, x As Integer
Dim fso As FileSystemObject, ts As TextStream
Dim nRecnum As Long, nLastRecNum As Long
Dim nListNum As Integer, nCurrenListItem As Long

nListNum = 0

If chkExportAll(nListNum).Value = 0 Then
    If lvList(nListNum).ListItems.Count = 0 Then Exit Sub
    nCurrenListItem = 1
    nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
    nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
    
GotoNextClassStart:
    nStatus = BTRCALL(BGETEQUAL, ClassPosBlock, Classdatabuf, Len(Classdatabuf), nRecnum, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nRecnum = nLastRecNum Then
            If nCurrenListItem = lvList(nListNum).ListItems.Count Then
                MsgBox "No class found to export.", vbInformation
                Exit Sub
            End If
            nCurrenListItem = nCurrenListItem + 1
            
            nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
            nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
        Else
            nRecnum = nRecnum + 1
        End If
        GoTo GotoNextClassStart:
    End If
Else
    nRecnum = 1
    nStatus = BTRCALL(BGETFIRST, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Classs: Could not get first record, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
End If
    
Set fso = CreateObject("Scripting.FileSystemObject")

stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_CLASS
stsStatusBar.Panels(2).Text = nRecnum

If format = "Access" Then GoTo Access:

Set ts = fso.OpenTextFile(ClassesTextfile, ForWriting)
ts.Write ("Number" & vbTab & "Name" & vbTab & "Min HP" & vbTab & "Max HP" & vbTab & "EXP %" & vbTab & "Magic LVL" & vbTab & "Combat Level" & vbTab & "Title Text" & vbTab & "Magic Type" & vbTab & "Weapon" & vbTab & "Armour" & vbTab)
For x = 0 To 9
    ts.Write ("Ability " & x & vbTab)
    ts.Write ("AbilVal " & x & vbTab)
Next x
ts.WriteLine ("")
    
Do While nStatus = 0 And Not bStopExport

    RowToStruct Classdatabuf.buf, ClassFldMap, Classrec, LenB(Classrec)

    ts.Write (Classrec.Number & vbTab)
    ts.Write (RTrim(RemoveCharacter(Classrec.Name, vbNull)) & vbTab)
    ts.Write (Classrec.MinHp & vbTab)
    ts.Write ((Classrec.MinHp + Classrec.MaxHP) & vbTab)
    ts.Write ((Classrec.Exp + 100) & vbTab)
    ts.Write (Classrec.MagicLvL & vbTab)
    ts.Write ((Classrec.Combat - 2) & vbTab)
    ts.Write (Classrec.TitleText & vbTab)
    ts.Write (Classrec.MagicType & vbTab)
    ts.Write (Classrec.Weapon & vbTab)
    ts.Write (Classrec.Armour & vbTab)
    
    For x = 0 To 9
        ts.Write (Classrec.AbilityA(x) & vbTab)
        ts.Write (Classrec.AbilityB(x) & vbTab)
    Next x

    ts.WriteLine ("")

    If chkExportAll(nListNum).Value = 0 Then
GotoNextClass:
        Call IncreaseProgressBar
        
        nStatus = BTRCALL(BGETNEXT, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            ClassRowToStruct Classdatabuf.buf
            
            If Classrec.Number > nLastRecNum Then
                If nCurrenListItem = lvList(nListNum).ListItems.Count Then GoTo Finished
                nCurrenListItem = nCurrenListItem + 1
                    
                nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
                nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
                
                nStatus = BTRCALL(BGETEQUAL, ClassPosBlock, Classdatabuf, Len(Classdatabuf), nRecnum, KEY_BUF_LEN, 0)
                If Not nStatus = 0 Then GoTo GotoNextClass:
            Else
                nRecnum = nRecnum + 1
            End If
        End If
    Else
        nStatus = BTRCALL(BGETNEXT, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
        nRecnum = nRecnum + 1
        Call IncreaseProgressBar
    End If
    
    stsStatusBar.Panels(2).Text = nRecnum
    If Not bUseCPU Then DoEvents
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error exporting Classes, Btrieve Error: " & BtrieveErrorCode(nStatus, True)
End If

Finished:

ts.Close
Set fso = Nothing
Set ts = Nothing

Exit Sub

Access:
'Dim adoConnect As Database
'Dim tabClasses As Recordset
'
'Set adoConnect = OpenDatabase(sDataSource)
'Set tabClasses = adoConnect.OpenRecordset("Classes")

tabClasses.Index = "pkClasses"
Do While nStatus = 0 And Not bStopExport
    
    RowToStruct Classdatabuf.buf, ClassFldMap, Classrec, LenB(Classrec)
    
    If bUpdateExistingADB = True Then
        If tabClasses.RecordCount = 0 Then
            tabClasses.AddNew
        Else
            tabClasses.Seek "=", Classrec.Number
            If tabClasses.NoMatch = True Then
                tabClasses.AddNew
            Else
                tabClasses.Edit
            End If
        End If
    Else
        tabClasses.AddNew
    End If
    
    tabClasses.Fields("Number") = Classrec.Number
    tabClasses.Fields("Name") = Classrec.Name
    tabClasses.Fields("Min HP") = Classrec.MinHp
    tabClasses.Fields("Max HP") = Classrec.MaxHP
    tabClasses.Fields("EXP %") = Classrec.Exp
    tabClasses.Fields("Magic Type") = Classrec.MagicType
    tabClasses.Fields("Magic LVL") = Classrec.MagicLvL
    tabClasses.Fields("Weapon") = Classrec.Weapon
    tabClasses.Fields("Armour") = Classrec.Armour
    tabClasses.Fields("Combat") = Classrec.Combat
    tabClasses.Fields("Title Text") = Classrec.TitleText
    
    For x = 0 To 9
        tabClasses.Fields("Ability " & x) = Classrec.AbilityA(x)
        tabClasses.Fields("Ability Value " & x) = Classrec.AbilityB(x)
    Next

    tabClasses.Update
    
    If chkExportAll(nListNum).Value = 0 Then
GotoNextClassAccess:
        Call IncreaseProgressBar
        
        nStatus = BTRCALL(BGETNEXT, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            ClassRowToStruct Classdatabuf.buf
            
            If Classrec.Number > nLastRecNum Then
                If nCurrenListItem = lvList(nListNum).ListItems.Count Then GoTo FinishedAccess
                nCurrenListItem = nCurrenListItem + 1
                    
                nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
                nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
                
                nStatus = BTRCALL(BGETEQUAL, ClassPosBlock, Classdatabuf, Len(Classdatabuf), nRecnum, KEY_BUF_LEN, 0)
                If Not nStatus = 0 Then GoTo GotoNextClassAccess:
            Else
                nRecnum = nRecnum + 1
            End If
        End If
    Else
        nStatus = BTRCALL(BGETNEXT, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
        nRecnum = nRecnum + 1
        Call IncreaseProgressBar
    End If
    
    stsStatusBar.Panels(2).Text = nRecnum
    If Not bUseCPU Then DoEvents
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error exporting Classes, Btrieve Error: " & BtrieveErrorCode(nStatus, True)
End If

FinishedAccess:
Set fso = Nothing
Set ts = Nothing

End Sub
Private Sub ExportRaces(format As String)
Dim nStatus As Integer, x As Integer
Dim fso As FileSystemObject, ts As TextStream
Dim nRecnum As Long, nLastRecNum As Long
Dim nListNum As Integer, nCurrenListItem As Long

nListNum = 1

If chkExportAll(nListNum).Value = 0 Then
    If lvList(nListNum).ListItems.Count = 0 Then Exit Sub
    nCurrenListItem = 1
    nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
    nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
    
GotoNextRaceStart:
    nStatus = BTRCALL(BGETEQUAL, RacePosBlock, Racedatabuf, Len(Racedatabuf), nRecnum, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nRecnum = nLastRecNum Then
            If nCurrenListItem = lvList(nListNum).ListItems.Count Then
                MsgBox "No race found to export.", vbInformation
                Exit Sub
            End If
            nCurrenListItem = nCurrenListItem + 1
            
            nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
            nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
        Else
            nRecnum = nRecnum + 1
        End If
        GoTo GotoNextRaceStart:
    End If
Else
    nRecnum = 1
    nStatus = BTRCALL(BGETFIRST, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Races: Could not get first record, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
End If
    
Set fso = CreateObject("Scripting.FileSystemObject")

stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_RACE
stsStatusBar.Panels(2).Text = nRecnum

If format = "Access" Then GoTo Access:

Set ts = fso.OpenTextFile(RacesTextfile, ForWriting)
ts.Write ("Number" & vbTab & "Name" & vbTab & "EXP %" & vbTab & "HP Bonus" & vbTab & "CP" & vbTab & "Min INT" & vbTab & "Min AGL" & vbTab & "Min STR" & vbTab & "Min WIL" & vbTab & "Min CHM" & vbTab & "Min HEA" & vbTab)
ts.Write ("Max INT" & vbTab & "Max AGL" & vbTab & "Max STR" & vbTab & "Max WIL" & vbTab & "Max CHM" & vbTab & "Max HEA" & vbTab)
For x = 0 To 9
    ts.Write ("Ability " & x & vbTab)
    ts.Write ("AbilVal " & x & vbTab)
Next
ts.WriteLine ("")
    
Do While nStatus = 0 And Not bStopExport

    RowToStruct Racedatabuf.buf, RaceFldMap, Racerec, LenB(Racerec)

    ts.Write (Racerec.Number & vbTab)
    ts.Write (RTrim(RemoveCharacter(Racerec.Name, vbNull)) & vbTab)
    ts.Write (Racerec.ExpChart & vbTab)
    ts.Write (Racerec.HPBonus & vbTab)
    ts.Write (Racerec.CP & vbTab)
    ts.Write (Racerec.MinInt & vbTab)
    ts.Write (Racerec.MinAgl & vbTab)
    ts.Write (Racerec.MinStr & vbTab)
    ts.Write (Racerec.MinWil & vbTab)
    ts.Write (Racerec.MinChm & vbTab)
    ts.Write (Racerec.MinHea & vbTab)
    ts.Write (Racerec.MaxInt & vbTab)
    ts.Write (Racerec.MaxAgl & vbTab)
    ts.Write (Racerec.MaxStr & vbTab)
    ts.Write (Racerec.MaxWil & vbTab)
    ts.Write (Racerec.MaxChm & vbTab)
    ts.Write (Racerec.MaxHea & vbTab)
    For x = 0 To 9
        ts.Write (Racerec.AbilityA(x) & vbTab)
        ts.Write (Racerec.AbilityB(x) & vbTab)
    Next

    ts.WriteLine ("")

    If chkExportAll(nListNum).Value = 0 Then
GotoNextRace:
        Call IncreaseProgressBar
        
        nStatus = BTRCALL(BGETNEXT, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            RaceRowToStruct Racedatabuf.buf
            
            If Racerec.Number > nLastRecNum Then
                If nCurrenListItem = lvList(nListNum).ListItems.Count Then GoTo Finished
                nCurrenListItem = nCurrenListItem + 1
                    
                nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
                nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
                
                nStatus = BTRCALL(BGETEQUAL, RacePosBlock, Racedatabuf, Len(Racedatabuf), nRecnum, KEY_BUF_LEN, 0)
                If Not nStatus = 0 Then GoTo GotoNextRace:
            Else
                nRecnum = nRecnum + 1
            End If
        End If
    Else
        nStatus = BTRCALL(BGETNEXT, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
        nRecnum = nRecnum + 1
        Call IncreaseProgressBar
    End If
    
    stsStatusBar.Panels(2).Text = nRecnum
    If Not bUseCPU Then DoEvents
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error exporting Races, Btrieve Error: " & BtrieveErrorCode(nStatus, True)
End If

Finished:

ts.Close
Set fso = Nothing
Set ts = Nothing

Exit Sub

Access:
'Dim adoConnect As Database
'Dim tabRaces As Recordset
'
'Set adoConnect = OpenDatabase(sDataSource)
'Set tabRaces = adoConnect.OpenRecordset("Races")

tabRaces.Index = "pkRaces"
Do While nStatus = 0 And Not bStopExport
    
    RowToStruct Racedatabuf.buf, RaceFldMap, Racerec, LenB(Racerec)
    
    If bUpdateExistingADB = True Then
        If tabRaces.RecordCount = 0 Then
            tabRaces.AddNew
        Else
            tabRaces.Seek "=", Racerec.Number
            If tabRaces.NoMatch = True Then
                tabRaces.AddNew
            Else
                tabRaces.Edit
            End If
        End If
    Else
        tabRaces.AddNew
    End If
    
    tabRaces.Fields("Number") = Racerec.Number
    tabRaces.Fields("Name") = Racerec.Name
    tabRaces.Fields("Min INT") = Racerec.MinInt
    tabRaces.Fields("Min WIL") = Racerec.MinWil
    tabRaces.Fields("Min STR") = Racerec.MinStr
    tabRaces.Fields("Min HEA") = Racerec.MinHea
    tabRaces.Fields("Min AGL") = Racerec.MinAgl
    tabRaces.Fields("Min CHM") = Racerec.MinChm
    tabRaces.Fields("Max INT") = Racerec.MaxInt
    tabRaces.Fields("Max WIL") = Racerec.MaxWil
    tabRaces.Fields("Max STR") = Racerec.MaxStr
    tabRaces.Fields("Max HEA") = Racerec.MaxHea
    tabRaces.Fields("Max AGL") = Racerec.MaxAgl
    tabRaces.Fields("Max CHM") = Racerec.MaxChm
    tabRaces.Fields("HP Bonus") = Racerec.HPBonus
    tabRaces.Fields("CP") = Racerec.CP
    tabRaces.Fields("EXP %") = Racerec.ExpChart
    
    For x = 0 To 9
        tabRaces.Fields("Ability " & x) = Racerec.AbilityA(x)
        tabRaces.Fields("Ability Value " & x) = Racerec.AbilityB(x)
    Next

    tabRaces.Update
    
    If chkExportAll(nListNum).Value = 0 Then
GotoNextRaceAccess:
        Call IncreaseProgressBar
        
        nStatus = BTRCALL(BGETNEXT, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            RaceRowToStruct Racedatabuf.buf
            
            If Racerec.Number > nLastRecNum Then
                If nCurrenListItem = lvList(nListNum).ListItems.Count Then GoTo FinishedAccess
                nCurrenListItem = nCurrenListItem + 1
                    
                nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
                nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
                
                nStatus = BTRCALL(BGETEQUAL, RacePosBlock, Racedatabuf, Len(Racedatabuf), nRecnum, KEY_BUF_LEN, 0)
                If Not nStatus = 0 Then GoTo GotoNextRaceAccess:
            Else
                nRecnum = nRecnum + 1
            End If
        End If
    Else
        nStatus = BTRCALL(BGETNEXT, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
        nRecnum = nRecnum + 1
        Call IncreaseProgressBar
    End If
    
    stsStatusBar.Panels(2).Text = nRecnum
    If Not bUseCPU Then DoEvents
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error exporting Races, Btrieve Error: " & BtrieveErrorCode(nStatus, True)
End If

FinishedAccess:
Set fso = Nothing
Set ts = Nothing

End Sub
Private Sub ExportShops(format As String)
Dim nStatus As Integer, x As Long
Dim fso As FileSystemObject, ts As TextStream
Dim nRecnum As Long, nLastRecNum As Long
Dim nListNum As Integer, nCurrenListItem As Long

nListNum = 5

If chkExportAll(nListNum).Value = 0 Then
    If lvList(nListNum).ListItems.Count = 0 Then Exit Sub
    nCurrenListItem = 1
    nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
    nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
    
GotoNextShopStart:
    nStatus = BTRCALL(BGETEQUAL, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), nRecnum, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nRecnum = nLastRecNum Then
            If nCurrenListItem = lvList(nListNum).ListItems.Count Then
                MsgBox "No shop found to export.", vbInformation
                Exit Sub
            End If
            nCurrenListItem = nCurrenListItem + 1
            
            nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
            nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
        Else
            nRecnum = nRecnum + 1
        End If
        GoTo GotoNextShopStart:
    End If
Else
    nRecnum = 1
    nStatus = BTRCALL(BGETFIRST, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Shops: Could not get first record, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
End If

Set fso = CreateObject("Scripting.FileSystemObject")

stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_SHOPS
stsStatusBar.Panels(2).Text = nRecnum

If format = "Access" Then GoTo Access:

Set ts = fso.OpenTextFile(ShopsTextfile, ForWriting)
ts.Write ("Number" & vbTab & "Name" & vbTab & "Shop Type" & vbTab & "Min LVL" & vbTab & "Max LVL" & vbTab & "Markup" & vbTab & "Class Limit" & vbTab)
For x = 0 To 19
    ts.Write ("Item " & x & vbTab)
    ts.Write ("Max " & x & vbTab)
    ts.Write ("Normal " & x & vbTab)
    ts.Write ("Regen Time " & x & vbTab)
    ts.Write ("Regen Number " & x & vbTab)
    ts.Write ("Regen %" & x & vbTab)
Next
ts.WriteLine ("")
    
Do While nStatus = 0 And Not bStopExport

    RowToStruct Shopdatabuf.buf, ShopFldMap, Shoprec, LenB(Shoprec)
    
    ts.Write (Shoprec.Number & vbTab)
    ts.Write (RTrim(RemoveCharacter(Shoprec.Name, vbNull)) & vbTab)
    ts.Write (Shoprec.ShopType & vbTab)
    ts.Write (Shoprec.ShopMinLvL & vbTab)
    ts.Write (Shoprec.ShopMaxLvl & vbTab)
    ts.Write (Shoprec.ShopMarkUp & vbTab)
    ts.Write (Shoprec.ShopClassLimit & vbTab)
    For x = 0 To 19
        ts.Write (Shoprec.ShopItemNumber(x) & vbTab)
        ts.Write (Shoprec.ShopMax(x) & vbTab)
        ts.Write (Shoprec.ShopNow(x) & vbTab)
        ts.Write (Shoprec.ShopRgnTime(x) & vbTab)
        ts.Write (Shoprec.ShopRgnNumber(x) & vbTab)
        ts.Write (Shoprec.ShopRgnPercentage(x) & vbTab)
    Next
    ts.WriteLine ("")

    If chkExportAll(nListNum).Value = 0 Then
GotoNextShop:
        Call IncreaseProgressBar
        
        nStatus = BTRCALL(BGETNEXT, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            ShopRowToStruct Shopdatabuf.buf
            
            If Shoprec.Number > nLastRecNum Then
                If nCurrenListItem = lvList(nListNum).ListItems.Count Then GoTo Finished
                nCurrenListItem = nCurrenListItem + 1
                    
                nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
                nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
                
                nStatus = BTRCALL(BGETEQUAL, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), nRecnum, KEY_BUF_LEN, 0)
                If Not nStatus = 0 Then GoTo GotoNextShop:
            Else
                nRecnum = nRecnum + 1
            End If
        End If
    Else
        nStatus = BTRCALL(BGETNEXT, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
        nRecnum = nRecnum + 1
        Call IncreaseProgressBar
    End If
    
    stsStatusBar.Panels(2).Text = nRecnum
    If Not bUseCPU Then DoEvents
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error exporting Shops, Btrieve Error: " & BtrieveErrorCode(nStatus, True)
End If

Finished:

ts.Close
Set fso = Nothing
Set ts = Nothing

Exit Sub

Access:
'Dim adoConnect As Database
'Dim tabShops As Recordset
'
'Set adoConnect = OpenDatabase(sDataSource)
'Set tabShops = adoConnect.OpenRecordset("Shops")

tabShops.Index = "pkShops"
Do While nStatus = 0 And Not bStopExport
    
    RowToStruct Shopdatabuf.buf, ShopFldMap, Shoprec, LenB(Shoprec)
    
    If bUpdateExistingADB = True Then
        If tabShops.RecordCount = 0 Then
            tabShops.AddNew
        Else
            tabShops.Seek "=", Shoprec.Number
            If tabShops.NoMatch = True Then
                tabShops.AddNew
            Else
                tabShops.Edit
            End If
        End If
    Else
        tabShops.AddNew
    End If
    
    tabShops.Fields("Number") = Shoprec.Number
    tabShops.Fields("Name") = Shoprec.Name
    tabShops.Fields("Desc A") = Shoprec.ShopDescriptionA
    tabShops.Fields("Desc B") = Shoprec.ShopDescriptionB
    tabShops.Fields("Desc C") = Shoprec.ShopDescriptionC
    tabShops.Fields("Type") = Shoprec.ShopType
    tabShops.Fields("Min Lvl") = Shoprec.ShopMinLvL
    tabShops.Fields("Max Lvl") = Shoprec.ShopMaxLvl
    tabShops.Fields("MarkUp") = Shoprec.ShopMarkUp
    tabShops.Fields("Class Limit") = Shoprec.ShopClassLimit
    
    For x = 0 To 19
        tabShops.Fields("Item " & x) = Shoprec.ShopItemNumber(x)
        
        If chkZeroUserInteraction.Value = 1 Then
            If Shoprec.ShopRgnPercentage(x) > 0 Then
                tabShops.Fields("Normal " & x) = Shoprec.ShopRgnNumber(x)
            Else
                tabShops.Fields("Normal " & x) = 0
            End If
        Else
            tabShops.Fields("Normal " & x) = Shoprec.ShopNow(x)
        End If
        
        tabShops.Fields("Max " & x) = Shoprec.ShopMax(x)
        tabShops.Fields("Regen Time " & x) = Shoprec.ShopRgnTime(x)
        tabShops.Fields("Regen Number" & x) = Shoprec.ShopRgnNumber(x)
        tabShops.Fields("Regen %" & x) = Shoprec.ShopRgnPercentage(x)
    Next

    tabShops.Update
    
    If chkExportAll(nListNum).Value = 0 Then
GotoNextShopAccess:
        Call IncreaseProgressBar
        
        nStatus = BTRCALL(BGETNEXT, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            ShopRowToStruct Shopdatabuf.buf
            
            If Shoprec.Number > nLastRecNum Then
                If nCurrenListItem = lvList(nListNum).ListItems.Count Then GoTo FinishedAccess
                nCurrenListItem = nCurrenListItem + 1
                    
                nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
                nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
                
                nStatus = BTRCALL(BGETEQUAL, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), nRecnum, KEY_BUF_LEN, 0)
                If Not nStatus = 0 Then GoTo GotoNextShopAccess:
            Else
                nRecnum = nRecnum + 1
            End If
        End If
    Else
        nStatus = BTRCALL(BGETNEXT, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
        nRecnum = nRecnum + 1
        Call IncreaseProgressBar
    End If
    
    stsStatusBar.Panels(2).Text = nRecnum
    If Not bUseCPU Then DoEvents
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error exporting Shops, Btrieve Error: " & BtrieveErrorCode(nStatus, True)
End If

FinishedAccess:
Set fso = Nothing
Set ts = Nothing

End Sub
Private Sub ExportMonsters(format As String)
Dim nStatus As Integer, x As Long
Dim fso As FileSystemObject, ts As TextStream
Dim nRecnum As Long, nLastRecNum As Long
Dim nListNum As Integer, nCurrenListItem As Long

nListNum = 4

If chkExportAll(nListNum).Value = 0 Then
    If lvList(nListNum).ListItems.Count = 0 Then Exit Sub
    nCurrenListItem = 1
    nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
    nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
    
GotoNextMonsterStart:
    nStatus = BTRCALL(BGETEQUAL, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), nRecnum, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nRecnum = nLastRecNum Then
            If nCurrenListItem = lvList(nListNum).ListItems.Count Then
                MsgBox "No monster found to export.", vbInformation
                Exit Sub
            End If
            nCurrenListItem = nCurrenListItem + 1
            
            nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
            nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
        Else
            nRecnum = nRecnum + 1
        End If
        GoTo GotoNextMonsterStart:
    End If
Else
    nRecnum = 1
    nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Monsters: Could not get first record, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
End If
    
Set fso = CreateObject("Scripting.FileSystemObject")

stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_KNMSR
stsStatusBar.Panels(2).Text = nRecnum

If format = "Access" Then GoTo Access:

Set ts = fso.OpenTextFile(MonstersTextfile, ForWriting)
ts.Write ("Number" & vbTab & "Name" & vbTab & "Experience" & vbTab & "Index" & vbTab & "Weapon" & vbTab & "AC" & vbTab & "DR" & vbTab & "Follow" & vbTab & "MR" & vbTab & "HPs" & vbTab & "BS Defense" & vbTab & "Energy" & vbTab & "HP Regen" & vbTab & "Game Limit" & vbTab & "Active" & vbTab & "Type" & vbTab & "Alignment" & vbTab & "Gender" & vbTab & "Group" & vbTab & "Regen Time" & vbTab)
ts.Write ("Date Killed" & vbTab & "Time Killed" & vbTab & "Charm LVL" & vbTab & "Charm RES" & vbTab & "Undead" & vbTab & "Move MSG" & vbTab & "Death MSG" & vbTab & "Greet TXT" & vbTab & "Desc TXT" & vbTab & "Talk TXT" & vbTab & "Death Spell" & vbTab & "Create Spell" & vbTab & "Runic" & vbTab & "Platinum" & vbTab & "Gold" & vbTab & "Silver" & vbTab & "Copper" & vbTab & "Desc 1" & vbTab & "Desc 2" & vbTab & "Desc 3" & vbTab & "Desc 4" & vbTab)
For x = 0 To 4
    ts.Write ("Attack Type " & x & vbTab)
    ts.Write ("Attack Accu/Spell " & x & vbTab)
    ts.Write ("Attack % " & x & vbTab)
    ts.Write ("Attack Min Hit/Cast % " & x & vbTab)
    ts.Write ("Attack Max Hit/Cast LVL " & x & vbTab)
    ts.Write ("Attack Hit Msg " & x & vbTab)
    ts.Write ("Attack Dodge Msg " & x & vbTab)
    ts.Write ("Attack Miss Msg " & x & vbTab)
    ts.Write ("Attack Energy " & x & vbTab)
    ts.Write ("Attack Hit Spell " & x & vbTab)
    ts.Write ("Spell Number " & x & vbTab)
    ts.Write ("Spell Cast % " & x & vbTab)
    ts.Write ("Spell Cast LVL " & x & vbTab)
Next

For x = 0 To 9
    ts.Write ("Item Number " & x & vbTab)
    ts.Write ("Item Uses " & x & vbTab)
    ts.Write ("Item Drop % " & x & vbTab)
Next

For x = 0 To 9
    ts.Write ("Ability " & x & vbTab)
    ts.Write ("AbilVal " & x & vbTab)
Next
ts.WriteLine ("")
    
Do While nStatus = 0 And Not bStopExport

    RowToStruct Monsterdatabuf.buf, MonsterFldMap, Monsterrec, LenB(Monsterrec)
    
    ts.Write (Monsterrec.Number & vbTab)
    ts.Write (RTrim(RemoveCharacter(Monsterrec.Name, vbNull)) & vbTab)
    
    If eDatFileVersion >= v111j Then
        ts.Write ((CDbl(SLong2ULong(Monsterrec.Experience)) * CDbl(SLong2ULong(Monsterrec.ExpMulti))) & vbTab)
    Else
        ts.Write (SLong2ULong(Monsterrec.Experience) & vbTab)
    End If
    
    ts.Write (Monsterrec.Index & vbTab)
    ts.Write (Monsterrec.WeaponNumber & vbTab)
    ts.Write (Monsterrec.AC & vbTab)
    ts.Write (Monsterrec.DR & vbTab)
    ts.Write (Monsterrec.Follow & vbTab)
    ts.Write (Monsterrec.MR & vbTab)
    ts.Write (Monsterrec.Hitpoints & vbTab)
    ts.Write (Monsterrec.BSDefence & vbTab)
    ts.Write (Monsterrec.Energy & vbTab)
    ts.Write (Monsterrec.HPRegen & vbTab)
    ts.Write (Monsterrec.GameLimit & vbTab)
    ts.Write (Monsterrec.Active & vbTab)
    ts.Write (Monsterrec.Type & vbTab)
    ts.Write (Monsterrec.Alignment & vbTab)
    ts.Write (Monsterrec.Gender & vbTab)
    ts.Write (Monsterrec.Group & vbTab)
    ts.Write (Monsterrec.RegenTime & vbTab)
    ts.Write (DOSDate2Date(SInt2UInt(Monsterrec.DateKilled)) & vbTab)
    ts.Write (DOSTime2Time(SInt2UInt(Monsterrec.TimeKilled)) & vbTab)
    ts.Write (Monsterrec.CharmLvL & vbTab)
    ts.Write (Monsterrec.CharmRes & vbTab)
    If Monsterrec.Undead > 1 Then Monsterrec.Undead = 0
    ts.Write (Monsterrec.Undead & vbTab)
    ts.Write (Monsterrec.MoveMsg & vbTab)
    ts.Write (Monsterrec.DeathMsg & vbTab)
    ts.Write (Monsterrec.GreetTxt & vbTab)
    ts.Write (Monsterrec.DescTxt & vbTab)
    ts.Write (Monsterrec.TalkTxt & vbTab)
    ts.Write (Monsterrec.DeathSpellNumber & vbTab)
    ts.Write (Monsterrec.CreateSpellNumber & vbTab)
    ts.Write (Monsterrec.Runic & vbTab)
    ts.Write (Monsterrec.Platinum & vbTab)
    ts.Write (Monsterrec.Gold & vbTab)
    ts.Write (Monsterrec.Silver & vbTab)
    ts.Write (Monsterrec.Copper & vbTab)
    
    ts.Write (RTrim(RemoveCharacter(Monsterrec.DescLine1, vbNull)) & vbTab)
    ts.Write (RTrim(RemoveCharacter(Monsterrec.DescLine2, vbNull)) & vbTab)
    ts.Write (RTrim(RemoveCharacter(Monsterrec.DescLine3, vbNull)) & vbTab)
    ts.Write (RTrim(RemoveCharacter(Monsterrec.DescLine4, vbNull)) & vbTab)
    
    For x = 0 To 4
        If Monsterrec.AttackType(x) > 3 Then Monsterrec.AttackType(x) = 0
        ts.Write (Monsterrec.AttackType(x) & vbTab)
        ts.Write (Monsterrec.AttackAccuSpell(x) & vbTab)
        ts.Write (Monsterrec.AttackPer(x) & vbTab)
        ts.Write (Monsterrec.AttackMinHCastPer(x) & vbTab)
        ts.Write (Monsterrec.AttackMaxHCastLvl(x) & vbTab)
        ts.Write (Monsterrec.AttackHitMsg(x) & vbTab)
        ts.Write (Monsterrec.AttackDodgeMsg(x) & vbTab)
        ts.Write (Monsterrec.AttackMissMsg(x) & vbTab)
        ts.Write (Monsterrec.AttackEnergy(x) & vbTab)
        ts.Write (Monsterrec.AttackHitSpell(x) & vbTab)
        ts.Write (Monsterrec.SpellNumber(x) & vbTab)
        ts.Write (Monsterrec.SpellCastPer(x) & vbTab)
        ts.Write (Monsterrec.SpellCastLvl(x) & vbTab)
    Next
    
    For x = 0 To 9
        ts.Write (Monsterrec.ItemNumber(x) & vbTab)
        ts.Write (Monsterrec.ItemUses(x) & vbTab)
        ts.Write (Monsterrec.ItemDropPer(x) & vbTab)
    Next
    
    For x = 0 To 9
        ts.Write (Monsterrec.AbilityA(x) & vbTab)
        ts.Write (Monsterrec.AbilityB(x) & vbTab)
    Next
    ts.WriteLine ("")
    
    If chkExportAll(nListNum).Value = 0 Then
GotoNextMonster:
        Call IncreaseProgressBar
        
        nStatus = BTRCALL(BGETNEXT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            MonsterRowToStruct Monsterdatabuf.buf
            
            If Monsterrec.Number > nLastRecNum Then
                If nCurrenListItem = lvList(nListNum).ListItems.Count Then GoTo Finished
                nCurrenListItem = nCurrenListItem + 1
                    
                nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
                nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
                
                nStatus = BTRCALL(BGETEQUAL, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), nRecnum, KEY_BUF_LEN, 0)
                If Not nStatus = 0 Then GoTo GotoNextMonster:
            Else
                nRecnum = nRecnum + 1
            End If
        End If
    Else
        nStatus = BTRCALL(BGETNEXT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
        nRecnum = nRecnum + 1
        Call IncreaseProgressBar
    End If
    
    stsStatusBar.Panels(2).Text = nRecnum
    If Not bUseCPU Then DoEvents
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error exporting Monsters, Btrieve Error: " & BtrieveErrorCode(nStatus, True)
End If

Finished:

ts.Close
Set fso = Nothing
Set ts = Nothing

Exit Sub

Access:
'Dim adoConnect As Database
'Dim tabMonsters As Recordset
'
'Set adoConnect = OpenDatabase(sDataSource)
'Set tabMonsters = adoConnect.OpenRecordset("Monsters")

Dim FieldTest As Boolean

If eDatFileVersion >= v111j And chkOneExpField.Value <> 1 Then
    FieldTest = TestMonsterFields
    If FieldTest = False Then
        MsgBox "Exported monster table does not contain the 'Exp Multiplier' Field, cannot export monsters to this file this way.", vbOKOnly
        GoTo FinishedAccess:
    End If
End If

tabMonsters.Index = "pkMonsters"
Do While nStatus = 0 And Not bStopExport
    
    RowToStruct Monsterdatabuf.buf, MonsterFldMap, Monsterrec, LenB(Monsterrec)
    
    If bUpdateExistingADB = True Then
        If tabMonsters.RecordCount = 0 Then
            tabMonsters.AddNew
        Else
            tabMonsters.Seek "=", Monsterrec.Number
            If tabMonsters.NoMatch = True Then
                tabMonsters.AddNew
            Else
                tabMonsters.Edit
            End If
        End If
    Else
        tabMonsters.AddNew
    End If
    
    tabMonsters.Fields("Number") = Monsterrec.Number
    tabMonsters.Fields("Name") = Monsterrec.Name
    tabMonsters.Fields("Group") = Monsterrec.Group
    tabMonsters.Fields("Index") = Monsterrec.Index
    tabMonsters.Fields("Weapon Number") = Monsterrec.WeaponNumber
    tabMonsters.Fields("AC") = Monsterrec.AC
    tabMonsters.Fields("DR") = Monsterrec.DR
    tabMonsters.Fields("Follow") = Monsterrec.Follow
    tabMonsters.Fields("MR") = Monsterrec.MR
    tabMonsters.Fields("Experience") = SLong2ULong(Monsterrec.Experience)
    
    If eDatFileVersion >= v111j Then
        If chkOneExpField.Value = 0 Then
            tabMonsters.Fields("Exp Multiplier") = SLong2ULong(Monsterrec.ExpMulti)
        Else
            tabMonsters.Fields("Experience") = CDbl(SLong2ULong(Monsterrec.Experience)) * CDbl(SLong2ULong(Monsterrec.ExpMulti))
        End If
    End If
    
    tabMonsters.Fields("Hit Points") = Monsterrec.Hitpoints
    tabMonsters.Fields("Energy") = Monsterrec.Energy
    tabMonsters.Fields("HP Regen") = Monsterrec.HPRegen
    tabMonsters.Fields("Game Limit") = Monsterrec.GameLimit
    tabMonsters.Fields("Charm LvL") = Monsterrec.CharmLvL
    tabMonsters.Fields("Charm RES") = Monsterrec.CharmRes
    tabMonsters.Fields("BS Defense") = Monsterrec.BSDefence
    
    tabMonsters.Fields("Type") = Monsterrec.Type
    tabMonsters.Fields("Undead") = Monsterrec.Undead
    tabMonsters.Fields("Alignment") = Monsterrec.Alignment
    tabMonsters.Fields("Regen Time") = Monsterrec.RegenTime
    
    tabMonsters.Fields("Move Msg") = Monsterrec.MoveMsg
    tabMonsters.Fields("Death Msg") = Monsterrec.DeathMsg
    tabMonsters.Fields("Runic") = Monsterrec.Runic
    tabMonsters.Fields("Platinum") = Monsterrec.Platinum
    tabMonsters.Fields("Gold") = Monsterrec.Gold
    tabMonsters.Fields("Silver") = Monsterrec.Silver
    tabMonsters.Fields("Copper") = Monsterrec.Copper
    tabMonsters.Fields("Greet Txt") = Monsterrec.GreetTxt
    tabMonsters.Fields("Desc Txt") = Monsterrec.DescTxt
    tabMonsters.Fields("Talk Txt") = Monsterrec.TalkTxt
    tabMonsters.Fields("Death Spell") = Monsterrec.DeathSpellNumber
    tabMonsters.Fields("Create Spell") = Monsterrec.CreateSpellNumber
    tabMonsters.Fields("Desc 1") = Monsterrec.DescLine1
    tabMonsters.Fields("Desc 2") = Monsterrec.DescLine2
    tabMonsters.Fields("Desc 3") = Monsterrec.DescLine3
    tabMonsters.Fields("Desc 4") = Monsterrec.DescLine4
    tabMonsters.Fields("Gender") = Monsterrec.Gender
    
    If chkZeroUserInteraction.Value = 1 Then
        tabMonsters.Fields("Active") = 0
        tabMonsters.Fields("Date Killed") = 0
        tabMonsters.Fields("Time Killed") = 0
    Else
        tabMonsters.Fields("Active") = Monsterrec.Active
        tabMonsters.Fields("Date Killed") = Monsterrec.DateKilled
        tabMonsters.Fields("Time Killed") = Monsterrec.TimeKilled
    End If
    
    For x = 0 To 4
        tabMonsters.Fields("Attack Type " & x) = Monsterrec.AttackType(x)
        tabMonsters.Fields("Attack Accu/Spell " & x) = Monsterrec.AttackAccuSpell(x)
        tabMonsters.Fields("Attack % " & x) = Monsterrec.AttackPer(x)
        tabMonsters.Fields("Attack Min Hit/Cast % " & x) = Monsterrec.AttackMinHCastPer(x)
        tabMonsters.Fields("Attack Max Hit/Cast LVL " & x) = Monsterrec.AttackMaxHCastLvl(x)
        tabMonsters.Fields("Attack Hit Msg " & x) = Monsterrec.AttackHitMsg(x)
        tabMonsters.Fields("Attack Dodge Msg " & x) = Monsterrec.AttackDodgeMsg(x)
        tabMonsters.Fields("Attack Miss Msg " & x) = Monsterrec.AttackMissMsg(x)
        tabMonsters.Fields("Attack Energy " & x) = Monsterrec.AttackEnergy(x)
        tabMonsters.Fields("Attack Hit Spell " & x) = Monsterrec.AttackHitSpell(x)
        tabMonsters.Fields("Spell Number " & x) = Monsterrec.SpellNumber(x)
        tabMonsters.Fields("Spell Cast % " & x) = Monsterrec.SpellCastPer(x)
        tabMonsters.Fields("Spell Cast LVL " & x) = Monsterrec.SpellCastLvl(x)
    Next
    
    For x = 0 To 9
        tabMonsters.Fields("Item Number " & x) = Monsterrec.ItemNumber(x)
        tabMonsters.Fields("Item Uses " & x) = Monsterrec.ItemUses(x)
        tabMonsters.Fields("Item Drop % " & x) = Monsterrec.ItemDropPer(x)
    Next
    
    For x = 0 To 9
        tabMonsters.Fields("Ability " & x) = Monsterrec.AbilityA(x)
        tabMonsters.Fields("Ability Value " & x) = Monsterrec.AbilityB(x)
    Next

    tabMonsters.Update
    
    If chkExportAll(nListNum).Value = 0 Then
GotoNextMonsterAccess:
        Call IncreaseProgressBar
        
        nStatus = BTRCALL(BGETNEXT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
        If nStatus = 0 Then
            MonsterRowToStruct Monsterdatabuf.buf
            
            If Monsterrec.Number > nLastRecNum Then
                If nCurrenListItem = lvList(nListNum).ListItems.Count Then GoTo FinishedAccess
                nCurrenListItem = nCurrenListItem + 1
                    
                nRecnum = Val(lvList(nListNum).ListItems(nCurrenListItem).Text)
                nLastRecNum = Val(lvList(nListNum).ListItems(nCurrenListItem).ListSubItems(1).Text)
                
                nStatus = BTRCALL(BGETEQUAL, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), nRecnum, KEY_BUF_LEN, 0)
                If Not nStatus = 0 Then GoTo GotoNextMonsterAccess:
            Else
                nRecnum = nRecnum + 1
            End If
        End If
    Else
        nStatus = BTRCALL(BGETNEXT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
        nRecnum = nRecnum + 1
        Call IncreaseProgressBar
    End If
    
    stsStatusBar.Panels(2).Text = nRecnum
    If Not bUseCPU Then DoEvents
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error exporting Monsters, Btrieve Error: " & BtrieveErrorCode(nStatus, True)
End If

FinishedAccess:
Set fso = Nothing
Set ts = Nothing

End Sub
Private Function TestMonsterFields() As Boolean
On Error GoTo error:
'Dim adoConnect As Database, tabMonsters As Recordset
Dim nTemp As Integer, fldTemp As field

'this function is just to test if the "Exp Multiplier" field exists. if not, it errors out

TestMonsterFields = False

'Set adoConnect = OpenDatabase(sDataSource)
'Set tabMonsters = adoConnect.OpenRecordset("Monsters")

nTemp = 0
For Each fldTemp In tabMonsters.Fields()
    If fldTemp.Name = "Exp Multiplier" Then nTemp = 1
Next

If nTemp = 1 Then TestMonsterFields = True

'tabMonsters.Close
'adoConnect.Close
'
'Set tabMonsters = Nothing
'Set adoConnect = Nothing

Exit Function
error:
End Function

Private Sub ExportUsers()
Dim nStatus As Integer, recnum As Long, x As Integer
Dim fso As FileSystemObject, ts As TextStream

Set fso = CreateObject("Scripting.FileSystemObject")
Set ts = fso.OpenTextFile(UsersTextfile, ForWriting)

recnum = 1
stsStatusBar.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_USERS
stsStatusBar.Panels(2).Text = recnum

nStatus = BTRCALL(BGETFIRST, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "User, BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

ts.Write ("BBS Name" & vbTab & "First Name" & vbTab & "Last Name" & vbTab & "Race" & vbTab & "Class" & vbTab & "LVL" & vbTab & "EXP" & vbTab & "Max HP" & vbTab & "HP" & vbTab & "Max Mana" & vbTab & "Mana" & vbTab & "SC" & vbTab & "Lives" & vbTab & "CP" & vbTab)
ts.Write ("Perception" & vbTab & "Stealth" & vbTab & "Thievery" & vbTab & "Traps" & vbTab & "Picklocks" & vbTab & "Tracking" & vbTab & "MA" & vbTab & "MR" & vbTab & "MR2" & vbTab & "Broadcast" & vbTab & "Runic" & vbTab & "Platinum" & vbTab & "Gold" & vbTab & "Silver" & vbTab & "Copper" & vbTab)
ts.Write ("Max ENC" & vbTab & "ENC" & vbTab & "EPs" & vbTab & "Gang" & vbTab & "Suicide Pass" & vbTab & "Title" & vbTab & "Room" & vbTab & "Map" & vbTab & "Weapon" & vbTab)

For x = 0 To 11
    ts.Write ("Stat " & x & vbTab)
Next

For x = 0 To 29
    ts.Write ("Ability(value) " & x & vbTab)
    'ts.Write ("AbilVal " & x & vbTab)
Next x

For x = 0 To 9
    ts.Write ("Spell Casted " & x & vbTab)
    ts.Write ("Spell Value " & x & vbTab)
    ts.Write ("Spell Rounds " & x & vbTab)
Next

ts.Write ("Last Map/Rooms" & vbTab)
ts.Write ("Worn Items" & vbTab)
ts.Write ("Items(uses)" & vbTab)
ts.Write ("Keys(uses) " & vbTab)
ts.WriteLine ("Spells Learned " & vbTab)
    
Do While nStatus = 0 And Not bStopExport

    RowToStruct Userdatabuf.buf, UserFldMap, Userrec, LenB(Userrec)
    
    ts.Write (ClipNull(Userrec.BBSName) & vbTab)
    ts.Write (ClipNull(Userrec.FirstName) & vbTab)
    ts.Write (ClipNull(Userrec.LastName) & vbTab)
    ts.Write (Userrec.Race & vbTab)
    ts.Write (Userrec.Class & vbTab)
    ts.Write (Userrec.Level & vbTab)
    ts.Write (((SLong2ULong(Userrec.BillionsOfExperience) * 1000000000#) + SLong2ULong(Userrec.MillionsOfExperience)) & vbTab)
    ts.Write (Userrec.MaxHP & vbTab)
    ts.Write (Userrec.CurrentHP & vbTab)
    ts.Write (Userrec.MaxMana & vbTab)
    ts.Write (Userrec.CurrentMana & vbTab)
    ts.Write (Userrec.SpellCasting & vbTab)
    ts.Write (Userrec.LivesRemaining & vbTab)
    ts.Write (Userrec.CPRemaining & vbTab)
    ts.Write (Userrec.Perception & vbTab)
    ts.Write (Userrec.Stealth & vbTab)
    ts.Write (Userrec.Thievery & vbTab)
    ts.Write (Userrec.Traps & vbTab)
    ts.Write (Userrec.Picklocks & vbTab)
    ts.Write (Userrec.Tracking & vbTab)
    ts.Write (Userrec.MartialArts & vbTab)
    ts.Write (Userrec.MagicRes & vbTab)
    ts.Write (Userrec.MagicRes2 & vbTab)
    ts.Write (Userrec.BroadcastChan & vbTab)
    ts.Write (SLong2ULong(Userrec.Runic) & vbTab)
    ts.Write (SLong2ULong(Userrec.Platinum) & vbTab)
    ts.Write (SLong2ULong(Userrec.Gold) & vbTab)
    ts.Write (SLong2ULong(Userrec.Silver) & vbTab)
    ts.Write (SLong2ULong(Userrec.Copper) & vbTab)
    ts.Write (Userrec.MaxENC & vbTab)
    ts.Write (Userrec.CurrentENC & vbTab)
    ts.Write (Userrec.EvilPoints & vbTab)
    ts.Write (RTrim(RemoveCharacter(Userrec.GangName, vbNull)) & vbTab)
    ts.Write (RTrim(RemoveCharacter(Userrec.SuicidePassword, vbNull)) & vbTab)
    ts.Write (RTrim(RemoveCharacter(Userrec.Title, vbNull)) & vbTab)
    ts.Write (Userrec.RoomNum & vbTab)
    ts.Write (Userrec.MapNumber & vbTab)
    ts.Write (Userrec.WeaponHand & vbTab)
    
    For x = 0 To 11
        ts.Write (Userrec.Stat(x) & vbTab)
    Next
    
    For x = 0 To 29
        ts.Write (Userrec.Ability(x) & "(" & Userrec.AbilityModifier(x) & ")" & vbTab)
    Next x
    
    For x = 0 To 9
        ts.Write (Userrec.SpellCasted(x) & vbTab)
        ts.Write (Userrec.SpellValue(x) & vbTab)
        ts.Write (Userrec.SpellRoundsLeft(x) & vbTab)
    Next
    
    For x = 0 To 19
        If x = 19 Then
            ts.Write (Userrec.LastMap(x) & "/" & Userrec.LastRoom(x) & vbTab)
        Else
            ts.Write (Userrec.LastMap(x) & "/" & Userrec.LastRoom(x) & ", ")
        End If
    Next
    
    For x = 0 To 19
        If x = 19 Then
            ts.Write (Userrec.WornItem(x) & vbTab)
        Else
            ts.Write (Userrec.WornItem(x) & ", ")
        End If
    Next
    
    For x = 0 To 99
        If x = 99 Then
            ts.Write (Userrec.Item(x) & "(" & Userrec.ItemUses(x) & ")" & vbTab)
        Else
            ts.Write (Userrec.Item(x) & "(" & Userrec.ItemUses(x) & "), ")
        End If
    Next
    
    For x = 0 To 49
        If x = 49 Then
            ts.Write (Userrec.Key(x) & "(" & Userrec.KeyUses(x) & ")" & vbTab)
        Else
            ts.Write (Userrec.Key(x) & "(" & Userrec.KeyUses(x) & "), ")
        End If
    Next
    
    For x = 0 To 99
        If x = 99 Then
            ts.Write (Userrec.Spell(x))
        Else
            ts.Write (Userrec.Spell(x) & ", ")
        End If
    Next

    ts.WriteLine ("")
    
    nStatus = BTRCALL(BGETNEXT, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0)
    
    recnum = recnum + 1
    stsStatusBar.Panels(2).Text = recnum
    IncreaseProgressBar
    If Not bUseCPU Then DoEvents

Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error exporting Users, Btrieve Error: " & BtrieveErrorCode(nStatus, True)
End If

ts.Close
Set fso = Nothing
Set ts = Nothing

End Sub
Private Function CreateDatabase() As Integer
On Error GoTo error:
Dim sTemp As String, nYesNo As Integer, catDB As ADOX.Catalog
Dim fso As FileSystemObject, x As Integer, y As Integer, nTemp As Integer

'0=not created
'1=created ok
'2=update existing
'3=cancel

CreateDatabase = 0

Set fso = CreateObject("Scripting.FileSystemObject")
sExportPath = ReadINI("Options", "ExportPath")
If Not fso.FolderExists(sExportPath) Then sExportPath = App.Path

sTemp = ReadINI("Options", "ExportFileName")
If Len(sTemp) < 5 Then sTemp = "NMR-DataExport.mdb"

CommonDialog1.Filter = "MDB Files (*.mdb)|*.mdb"
CommonDialog1.DialogTitle = "Select Export File/Enter New File Name"
CommonDialog1.FileName = sTemp
CommonDialog1.InitDir = sExportPath

On Error GoTo canceled:
CommonDialog1.ShowSave
If CommonDialog1.FileName = "" Then GoTo canceled:

On Error GoTo error:
sDataSource = CommonDialog1.FileName

If Not LCase(Right(sDataSource, 4)) = ".mdb" Then sDataSource = sDataSource & ".mdb"

sTemp = CommonDialog1.FileTitle
If Not LCase(Right(sTemp, 4)) = ".mdb" Then sTemp = sTemp & ".mdb"
Call WriteINI("Options", "ExportFileName", sTemp)

If fso.FileExists(sDataSource) = True Then
    nYesNo = MsgBox("'" & sDataSource & "' already exists." & vbCrLf & vbCrLf _
        & "Attempt to add to and/or update it?" & vbCrLf & vbCrLf _
        & "NOTE: When exporting dats with a high number of records (ie. rooms), " & vbCrLf _
        & "adding/updating will significantly slow down the exporting process." & vbCrLf & vbCrLf _
        & "Click Yes to update, No to delete and create new, or Cancel to cancel.", vbYesNoCancel + vbQuestion, "File already exits...")
    
    If nYesNo = vbNo Then
        fso.DeleteFile sDataSource, True
    ElseIf nYesNo = vbYes Then
        CreateDatabase = 2
        Set fso = Nothing
        Exit Function
    Else
        CreateDatabase = 3
        Set fso = Nothing
        Exit Function
    End If
End If

bUpdateExistingADB = False

'create database
stsStatusBar.Panels(2).Text = "Creating Database..."
DoEvents
Set catDB = New ADOX.Catalog
catDB.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sDataSource
Set catDB = Nothing
DoEvents

CreateDatabase = 1

Set fso = Nothing
Set catDB = Nothing

Exit Function

canceled:
CreateDatabase = 3
Set fso = Nothing
Set catDB = Nothing
Exit Function

error:
Call HandleError
Set fso = Nothing
Set catDB = Nothing

End Function


Private Sub Form_Unload(Cancel As Integer)
Dim nTemp As Integer
    If cmdGo.Enabled = False Then
        Cancel = 1
        Exit Sub
    End If
    
    If bCheckSave Then
        nTemp = MsgBox("Save current config file first?", vbYesNoCancel + vbQuestion, "Save Config?")
        If nTemp = vbYes Then
            nTemp = SaveConfig(sConfigFile)
            If nTemp = -1 Then
                Cancel = 1
                Exit Sub
            End If
        ElseIf nTemp = vbCancel Then
            Cancel = 1
            Exit Sub
        End If
    End If

    Call WriteINI("Options", "NMR-ExportConfig", sConfigFile)
    
    If Not Me.WindowState = vbMinimized Then
        Call WriteINI("Windows", "ExportTop", Me.Top)
        Call WriteINI("Windows", "ExportLeft", Me.Left)
    End If
    
    Call CloseAll(True)
End Sub


Private Sub lvList_Click(Index As Integer)
cmbDB.ListIndex = Index
End Sub

Private Sub lvList_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
SortListView lvList(Index), ColumnHeader.Index, ldtNumber, True
End Sub

Private Sub optAccessDB_Click()
    chkBankbooks.Value = 0
    chkBankbooks.Enabled = False
    chkUsers.Value = 0
    chkUsers.Enabled = False
    chkZeroUserInteraction.Enabled = True
    If eDatFileVersion >= v111j Then chkOneExpField.Enabled = True
    bCheckSave = True
End Sub

Private Sub optTextfile_Click()
    chkBankbooks.Enabled = True
    chkUsers.Enabled = True
    chkOneExpField.Enabled = False
    chkZeroUserInteraction.Enabled = False
    If eDatFileVersion >= v111j Then chkOneExpField.Value = 0
    bCheckSave = True
End Sub

Private Function CalcTotalRecords() As Long
On Error GoTo error:
Dim nStatus As Integer, x As Long, y As Integer

CalcTotalRecords = 0

For x = 0 To 8
    If chkExportAll(x).Value = 1 Then
        Select Case x
            Case 0:
                nStatus = BTRCALL(BSTAT, ClassPosBlock, DBStatDatabuf, Len(Classdatabuf), 0, KEY_BUF_LEN, 0)
            Case 1:
                nStatus = BTRCALL(BSTAT, RacePosBlock, DBStatDatabuf, Len(Racedatabuf), 0, KEY_BUF_LEN, 0)
            Case 2:
                nStatus = BTRCALL(BSTAT, ItemPosBlock, DBStatDatabuf, Len(Itemdatabuf), 0, KEY_BUF_LEN, 0)
            Case 3:
                nStatus = BTRCALL(BSTAT, MessagePosBlock, DBStatDatabuf, Len(Messagedatabuf), 0, KEY_BUF_LEN, 0)
            Case 4:
                nStatus = BTRCALL(BSTAT, MonsterPosBlock, DBStatDatabuf, Len(Monsterdatabuf), 0, KEY_BUF_LEN, 0)
            Case 5:
                nStatus = BTRCALL(BSTAT, ShopPosBlock, DBStatDatabuf, Len(Shopdatabuf), 0, KEY_BUF_LEN, 0)
            Case 6:
                nStatus = BTRCALL(BSTAT, SpellPosBlock, DBStatDatabuf, Len(Spelldatabuf), 0, KEY_BUF_LEN, 0)
            Case 7:
                nStatus = BTRCALL(BSTAT, TextblockPosBlock, DBStatDatabuf, Len(TextblockDataBuf), 0, KEY_BUF_LEN, 0)
            Case 8:
                nStatus = BTRCALL(BSTAT, RoomPosBlock, DBStatDatabuf, Len(Roomdatabuf), 0, KEY_BUF_LEN, 0)
        End Select
        
        If nStatus = 0 Then
            DBStatRowToStruct DBStatDatabuf.buf
            CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
        End If
    Else
        If lvList(x).ListItems.Count > 0 Then
            If x = 8 Then
                For y = 1 To lvList(x).ListItems.Count
                    CalcTotalRecords = CalcTotalRecords + Val(lvList(x).ListItems(y).ListSubItems(2).Text) - Val(lvList(x).ListItems(y).ListSubItems(1).Text) + 1
                Next y
            Else
                For y = 1 To lvList(x).ListItems.Count
                    CalcTotalRecords = CalcTotalRecords + Val(lvList(x).ListItems(y).ListSubItems(1).Text) - Val(lvList(x).ListItems(y).Text) + 1
                Next y
            End If
        End If
    End If
Next x

If chkActions.Value = 1 Then
    nStatus = BTRCALL(BSTAT, ActionPosBlock, DBStatDatabuf, Len(ActionDatabuf), 0, KEY_BUF_LEN, 0)
    If nStatus = 0 Then
        DBStatRowToStruct DBStatDatabuf.buf
        CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
    End If
End If

If chkUsers.Value = 1 Then
    nStatus = BTRCALL(BSTAT, UserPosBlock, DBStatDatabuf, Len(Userdatabuf), 0, KEY_BUF_LEN, 0)
    If nStatus = 0 Then
        DBStatRowToStruct DBStatDatabuf.buf
        CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
    End If
End If

If chkBankbooks.Value = 1 Then
    nStatus = BTRCALL(BSTAT, BankPosBlock, DBStatDatabuf, Len(BankDatabuf), 0, KEY_BUF_LEN, 0)
    If nStatus = 0 Then
        DBStatRowToStruct DBStatDatabuf.buf
        CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
    End If
End If

If CalcTotalRecords <= 0 Then CalcTotalRecords = 1

Exit Function

error:
Call HandleError("CalcTotalRecords")
End Function
Private Sub IncreaseProgressBar()
On Error Resume Next
'If ProgressBar.Value + 1 < ProgressBar.Max Then ProgressBar.Value = ProgressBar.Value + 1

If nScale > 0 Then
    If nScaleCount = nScale Then
        If ProgressBar.Value + 1 < ProgressBar.Max Then ProgressBar.Value = ProgressBar.Value + 1
        nScaleCount = 1
    Else
        nScaleCount = nScaleCount + 1
    End If
Else
    If ProgressBar.Value + 1 < ProgressBar.Max Then ProgressBar.Value = ProgressBar.Value + 1
End If

End Sub

Private Sub txtCustom_GotFocus()
Call SelectAll(txtCustom)
End Sub

Private Sub txtFrom_GotFocus()
Call SelectAll(txtFrom)
End Sub


Private Sub txtMap_GotFocus()
Call SelectAll(txtMap)
End Sub

Private Sub txtTo_GotFocus()
Call SelectAll(txtTo)
End Sub
Private Sub LoadConfig(ByVal sFile As String)
Dim sLine As String, x As Long, y As Long, oLI As ListItem
Dim sTemp As String, sArray() As String, fso As FileSystemObject
On Error GoTo error:

sConfigFile = sFile
txtConfigFile.Text = sConfigFile

Set fso = CreateObject("Scripting.FileSystemObject")

sExportPath = ReadINI("Settings", "ExportPath", sFile)
If Not fso.FolderExists(sExportPath) Then sExportPath = App.Path

sLine = ReadINI("Settings", "CustomName", sFile, "Custom Export")
If Not sLine = "0" Then
    txtCustom.Text = sLine
Else
    txtCustom.Text = "Custom Export"
End If

chkZeroUserInteraction.Value = ReadINI("Settings", "ZeroUserInteraction", sFile, "1")

For x = 0 To 8
    lvList(x).ListItems.clear
    chkExportAll(x).Value = ReadINI("Records_" & x, "List_All", sFile, "1")
    
    y = 1
    sLine = ReadINI("Records_" & x, "List_" & y, sFile)
    Do While InStr(1, sLine, "/") > 0
        sArray() = Split(sLine, "/", 3)
        
        If UBound(sArray()) >= 1 Then
            Set oLI = lvList(x).ListItems.add
            oLI.Text = sArray(0)
            oLI.ListSubItems.add 1, , sArray(1)
            If x = 8 And UBound(sArray()) >= 2 Then
                oLI.ListSubItems.add 2, , sArray(2)
            End If
        End If
        
        y = y + 1
        sLine = ReadINI("Records_" & x, "List_" & y, sFile)
    Loop
Next x

If ReadINI("Settings", "Format", sFile, "Access") = "Textfiles" Then
    optTextfile.Value = True
    Call optTextfile_Click
    chkActions.Value = ReadINI("Settings", "Actions", sFile, "1")
    chkUsers.Value = ReadINI("Settings", "Users", sFile, "0")
    chkBankbooks.Value = ReadINI("Settings", "Bankbooks", sFile, "0")
Else
    optAccessDB.Value = True
    chkOneExpField.Value = ReadINI("Settings", "OneExpField", sFile, "0")
    Call optAccessDB_Click
    chkActions.Value = ReadINI("Settings", "Actions", sFile, "1")
End If

If eDatFileVersion < v111j Then
    chkOneExpField.Value = 1
    chkOneExpField.Enabled = False
End If

Call UpdateListStuff
bCheckSave = False

out:
Erase sArray()
Set oLI = Nothing
Set fso = Nothing
Exit Sub
error:
Call HandleError("LoadConfig")
Resume out:
    
End Sub

Private Function SaveConfig(ByVal sFile As String, _
    Optional ByVal bPromptFile As Boolean) As Integer
Dim sTemp As String, x As Integer, y As Long
On Error GoTo error:

If bPromptFile Then
    CommonDialog1.Filter = "INI Files (*.ini)|*.ini"
    CommonDialog1.DialogTitle = "Select Export Configuration File ..."
    CommonDialog1.FileName = sConfigFile
    CommonDialog1.InitDir = sConfigFile
    
    On Error GoTo canceled:
    CommonDialog1.ShowSave
    If CommonDialog1.FileName = "" Then GoTo canceled:
    
    On Error GoTo error:
    
    sTemp = CommonDialog1.FileName
    If Right(sTemp, 4) <> ".ini" Then sTemp = sTemp & ".ini"
    
    sFile = sTemp
End If

sConfigFile = sFile
txtConfigFile.Text = sConfigFile

Call WriteINI("Settings", "CustomName", txtCustom.Text, sFile)
Call WriteINI("Settings", "ExportPath", sExportPath, sFile)

If optAccessDB.Value = True Then
    Call WriteINI("Settings", "Format", "Access", sFile)
Else
    Call WriteINI("Settings", "Format", "Textfiles", sFile)
End If

Call WriteINI("Settings", "Actions", chkActions.Value, sFile)
Call WriteINI("Settings", "Users", chkUsers.Value, sFile)
Call WriteINI("Settings", "Bankbooks", chkBankbooks.Value, sFile)

Call WriteINI("Settings", "OneExpField", chkOneExpField.Value, sFile)
Call WriteINI("Settings", "ZeroUserInteraction", chkZeroUserInteraction.Value, sFile)

For x = 0 To 8
    If chkExportAll(x).Value = 1 Then
        Call WriteINI("Records_" & x, "List_All", "1", sFile)
    Else
        Call WriteINI("Records_" & x, "List_All", "0", sFile)
    End If
    
    y = 1
    If lvList(x).ListItems.Count > 0 Then
        For y = 1 To lvList(x).ListItems.Count
            sTemp = lvList(x).ListItems(y).Text & "/" & lvList(x).ListItems(y).SubItems(1)
            If x = 8 Then
                sTemp = sTemp & "/" & lvList(x).ListItems(y).SubItems(2)
            End If
            Call WriteINI("Records_" & x, "List_" & y, sTemp, sFile)
        Next y
    End If
    Call WriteINI("Records_" & x, "List_" & y, "End", sFile)
Next x

bCheckSave = False

GoTo out:

canceled:
SaveConfig = -1

out:
Exit Function
error:
Call HandleError("SaveConfig")
Resume out:

End Function

Private Sub UpdateListStuff()
Dim x As Integer
On Error GoTo error:

For x = 0 To 8
    If chkExportAll(x).Value = 1 Then
        lvList(x).BackColor = &H80000010
        If lvList(x).ListItems.Count > 0 Then
            If Not lvList(x).SelectedItem Is Nothing Then
                lvList(x).SelectedItem.Selected = False
                Set lvList(x).SelectedItem = Nothing
                lvList(x).HideSelection = True
            End If
        End If
    Else
        lvList(x).BackColor = &H80000005
    End If
    
    If cmbDB.ListIndex = x Then
        shpOutline(x).BorderColor = &H8000000D
    Else
        shpOutline(x).BorderColor = &H8000000F
    End If
Next x

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("UpdateListStuff")
End Sub

Private Sub CombineRanges()
Dim nRecIndex As Integer, y As Long
Dim nLastMap As Long, nLastLow As Long, nLastHigh As Long
On Error GoTo error:

nLastMap = 0
nLastLow = 0
nLastHigh = 0

For nRecIndex = 0 To 8
    If lvList(nRecIndex).ListItems.Count > 0 Then
        If nRecIndex = 8 Then
            SortListView lvList(nRecIndex), 3, ldtNumber, True
        End If
        SortListView lvList(nRecIndex), 2, ldtNumber, True
        SortListView lvList(nRecIndex), 1, ldtNumber, True
        DoEvents
        
start_over:
        For y = 1 To lvList(nRecIndex).ListItems.Count
            If y = 1 Then
                If nRecIndex = 8 Then
                    nLastMap = Val(lvList(nRecIndex).ListItems(y).Text)
                    nLastLow = Val(lvList(nRecIndex).ListItems(y).ListSubItems(1))
                    nLastHigh = Val(lvList(nRecIndex).ListItems(y).ListSubItems(2))
                Else
                    nLastLow = Val(lvList(nRecIndex).ListItems(y).Text)
                    nLastHigh = Val(lvList(nRecIndex).ListItems(y).ListSubItems(1))
                End If
            Else
                If nRecIndex = 8 Then
                    'ROOM
                    If nLastMap = Val(lvList(nRecIndex).ListItems(y).Text) Then
                        If Val(lvList(nRecIndex).ListItems(y).ListSubItems(1)) >= nLastLow _
                            And Val(lvList(nRecIndex).ListItems(y).ListSubItems(1)) <= nLastHigh + 1 Then
                            
                            If Val(lvList(nRecIndex).ListItems(y).ListSubItems(2)) >= Val(lvList(nRecIndex).ListItems(y - 1).ListSubItems(2)) Then
                                lvList(nRecIndex).ListItems(y - 1).ListSubItems(2) = lvList(nRecIndex).ListItems(y).ListSubItems(2)
                            End If
                            
                            lvList(nRecIndex).ListItems.Remove y
                            bCheckSave = True
                            GoTo start_over:
                        Else
                            nLastLow = Val(lvList(nRecIndex).ListItems(y).ListSubItems(1))
                            nLastHigh = Val(lvList(nRecIndex).ListItems(y).ListSubItems(2))
                        End If
                    Else
                        nLastMap = Val(lvList(nRecIndex).ListItems(y).Text)
                        nLastLow = Val(lvList(nRecIndex).ListItems(y).ListSubItems(1))
                        nLastHigh = Val(lvList(nRecIndex).ListItems(y).ListSubItems(2))
                    End If
                    
                Else
                    'NON-ROOM
                    If Val(lvList(nRecIndex).ListItems(y).Text) >= nLastLow _
                        And Val(lvList(nRecIndex).ListItems(y).Text) <= nLastHigh + 1 Then
                        
                        If Val(lvList(nRecIndex).ListItems(y).ListSubItems(1)) >= Val(lvList(nRecIndex).ListItems(y - 1).ListSubItems(1)) Then
                            lvList(nRecIndex).ListItems(y - 1).ListSubItems(1) = lvList(nRecIndex).ListItems(y).ListSubItems(1)
                        End If
                        
                        lvList(nRecIndex).ListItems.Remove y
                        bCheckSave = True
                        GoTo start_over:
                    Else
                        nLastLow = Val(lvList(nRecIndex).ListItems(y).Text)
                        nLastHigh = Val(lvList(nRecIndex).ListItems(y).ListSubItems(1))
                    End If
                    
                End If
            End If
        Next y
    End If
Next nRecIndex

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("CombineRanges")
Resume out:
End Sub


