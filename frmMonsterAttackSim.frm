VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMonsterAttackSim 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monster Attack Simulator"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   14685
   Icon            =   "frmMonsterAttackSim.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   14685
   Begin VB.CheckBox chkDynamicRounds 
      Alignment       =   1  'Right Justify
      Caption         =   "or Dynamic:"
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
      Left            =   12900
      TabIndex        =   221
      ToolTipText     =   "This will run the sim in 1,000 round increments untl the change in result is < 0.001%"
      Top             =   6360
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CommandButton cmdGotoMon 
      Caption         =   ">"
      Height          =   375
      Left            =   5100
      TabIndex        =   2
      ToolTipText     =   "Goto Monster"
      Top             =   60
      Width           =   375
   End
   Begin VB.Frame fraAttacks 
      Caption         =   "Regular Attacks"
      Height          =   2895
      Left            =   60
      TabIndex        =   6
      Top             =   540
      Width           =   9915
      Begin VB.TextBox txtAtkDur 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   8580
         TabIndex        =   82
         Top             =   2400
         Width           =   675
      End
      Begin VB.TextBox txtAtkDur 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   8580
         TabIndex        =   69
         Top             =   1980
         Width           =   675
      End
      Begin VB.TextBox txtAtkDur 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   8580
         TabIndex        =   56
         Top             =   1560
         Width           =   675
      End
      Begin VB.TextBox txtAtkDur 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   8580
         TabIndex        =   43
         Top             =   1140
         Width           =   675
      End
      Begin VB.TextBox txtAtkDur 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   8580
         TabIndex        =   30
         Top             =   720
         Width           =   675
      End
      Begin VB.TextBox txtAtkHitSpellMax 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   6960
         TabIndex        =   80
         Top             =   2400
         Width           =   555
      End
      Begin VB.TextBox txtAtkHitSpellMin 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   6420
         TabIndex        =   79
         Top             =   2400
         Width           =   555
      End
      Begin VB.TextBox txtAtkHitSpellMax 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   6960
         TabIndex        =   67
         Top             =   1980
         Width           =   555
      End
      Begin VB.TextBox txtAtkHitSpellMin 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   6420
         TabIndex        =   66
         Top             =   1980
         Width           =   555
      End
      Begin VB.TextBox txtAtkHitSpellMax 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   6960
         TabIndex        =   54
         Top             =   1560
         Width           =   555
      End
      Begin VB.TextBox txtAtkHitSpellMin 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   6420
         TabIndex        =   53
         Top             =   1560
         Width           =   555
      End
      Begin VB.TextBox txtAtkHitSpellMax 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   6960
         TabIndex        =   40
         Top             =   1140
         Width           =   555
      End
      Begin VB.TextBox txtAtkHitSpellMin 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   6420
         TabIndex        =   39
         Top             =   1140
         Width           =   555
      End
      Begin VB.TextBox txtAtkHitSpellMax 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   6960
         TabIndex        =   28
         Top             =   720
         Width           =   555
      End
      Begin VB.TextBox txtAtkHitSpellMin 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   6420
         TabIndex        =   27
         Top             =   720
         Width           =   555
      End
      Begin VB.CheckBox chkAtkDmgResist 
         Height          =   255
         Index           =   4
         Left            =   9480
         TabIndex        =   83
         Top             =   2460
         Width           =   255
      End
      Begin VB.CheckBox chkAtkDmgResist 
         Height          =   255
         Index           =   3
         Left            =   9480
         TabIndex        =   70
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox chkAtkDmgResist 
         Height          =   255
         Index           =   2
         Left            =   9480
         TabIndex        =   57
         Top             =   1620
         Width           =   255
      End
      Begin VB.CheckBox chkAtkDmgResist 
         Height          =   255
         Index           =   1
         Left            =   9480
         TabIndex        =   44
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox chkAtkDmgResist 
         Height          =   255
         Index           =   0
         Left            =   9480
         TabIndex        =   31
         Top             =   780
         Width           =   255
      End
      Begin VB.TextBox txtAtkSuccess 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   4080
         TabIndex        =   76
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtAtkChance 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   2520
         TabIndex        =   74
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtAtkMax 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   5640
         TabIndex        =   78
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtAtkMin 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   4860
         TabIndex        =   77
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtAtkEnergy 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   3300
         TabIndex        =   75
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtAtkName 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   300
         MaxLength       =   20
         TabIndex        =   71
         Top             =   2400
         Width           =   1095
      End
      Begin VB.ComboBox cmbAtkType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         ItemData        =   "frmMonsterAttackSim.frx":08CA
         Left            =   1440
         List            =   "frmMonsterAttackSim.frx":08DA
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   2400
         Width           =   1035
      End
      Begin VB.ComboBox cmbAtkResist 
         Height          =   315
         Index           =   4
         ItemData        =   "frmMonsterAttackSim.frx":08F8
         Left            =   7560
         List            =   "frmMonsterAttackSim.frx":0905
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtAtkSuccess 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   4080
         TabIndex        =   63
         Top             =   1980
         Width           =   735
      End
      Begin VB.TextBox txtAtkChance 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   2520
         TabIndex        =   61
         Top             =   1980
         Width           =   735
      End
      Begin VB.TextBox txtAtkMax 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   5640
         TabIndex        =   65
         Top             =   1980
         Width           =   735
      End
      Begin VB.TextBox txtAtkMin 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   4860
         TabIndex        =   64
         Top             =   1980
         Width           =   735
      End
      Begin VB.TextBox txtAtkEnergy 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   3300
         TabIndex        =   62
         Top             =   1980
         Width           =   735
      End
      Begin VB.TextBox txtAtkName 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   300
         MaxLength       =   20
         TabIndex        =   58
         Top             =   1980
         Width           =   1095
      End
      Begin VB.ComboBox cmbAtkType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         ItemData        =   "frmMonsterAttackSim.frx":0926
         Left            =   1440
         List            =   "frmMonsterAttackSim.frx":0936
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   1980
         Width           =   1035
      End
      Begin VB.ComboBox cmbAtkResist 
         Height          =   315
         Index           =   3
         ItemData        =   "frmMonsterAttackSim.frx":0954
         Left            =   7560
         List            =   "frmMonsterAttackSim.frx":0961
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   1980
         Width           =   975
      End
      Begin VB.TextBox txtAtkSuccess 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   4080
         TabIndex        =   50
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtAtkChance 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   2520
         TabIndex        =   48
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtAtkMax 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   5640
         TabIndex        =   52
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtAtkMin 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   4860
         TabIndex        =   51
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtAtkEnergy 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3300
         TabIndex        =   49
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtAtkName 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   300
         MaxLength       =   20
         TabIndex        =   45
         Top             =   1560
         Width           =   1095
      End
      Begin VB.ComboBox cmbAtkType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         ItemData        =   "frmMonsterAttackSim.frx":0982
         Left            =   1440
         List            =   "frmMonsterAttackSim.frx":0992
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   1560
         Width           =   1035
      End
      Begin VB.ComboBox cmbAtkResist 
         Height          =   315
         Index           =   2
         ItemData        =   "frmMonsterAttackSim.frx":09B0
         Left            =   7560
         List            =   "frmMonsterAttackSim.frx":09BD
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtAtkSuccess 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   4080
         TabIndex        =   36
         Top             =   1140
         Width           =   735
      End
      Begin VB.TextBox txtAtkChance 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   2520
         TabIndex        =   34
         Top             =   1140
         Width           =   735
      End
      Begin VB.TextBox txtAtkMax 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   5640
         TabIndex        =   38
         Top             =   1140
         Width           =   735
      End
      Begin VB.TextBox txtAtkMin 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   4860
         TabIndex        =   37
         Top             =   1140
         Width           =   735
      End
      Begin VB.TextBox txtAtkEnergy 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   3300
         TabIndex        =   35
         Top             =   1140
         Width           =   735
      End
      Begin VB.TextBox txtAtkName 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   300
         MaxLength       =   20
         TabIndex        =   32
         Top             =   1140
         Width           =   1095
      End
      Begin VB.ComboBox cmbAtkType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         ItemData        =   "frmMonsterAttackSim.frx":09DE
         Left            =   1440
         List            =   "frmMonsterAttackSim.frx":09EE
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1140
         Width           =   1035
      End
      Begin VB.ComboBox cmbAtkResist 
         Height          =   315
         Index           =   1
         ItemData        =   "frmMonsterAttackSim.frx":0A0C
         Left            =   7560
         List            =   "frmMonsterAttackSim.frx":0A19
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   1140
         Width           =   975
      End
      Begin VB.TextBox txtAtkSuccess 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   4080
         TabIndex        =   24
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtAtkChance 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   2520
         TabIndex        =   22
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtAtkMax 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   5640
         TabIndex        =   26
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtAtkMin 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   4860
         TabIndex        =   25
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtAtkEnergy 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   3300
         TabIndex        =   23
         Top             =   720
         Width           =   735
      End
      Begin VB.ComboBox cmbAtkResist 
         Height          =   315
         Index           =   0
         ItemData        =   "frmMonsterAttackSim.frx":0A3A
         Left            =   7560
         List            =   "frmMonsterAttackSim.frx":0A47
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox cmbAtkType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         ItemData        =   "frmMonsterAttackSim.frx":0A68
         Left            =   1440
         List            =   "frmMonsterAttackSim.frx":0A78
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   720
         Width           =   1035
      End
      Begin VB.TextBox txtAtkName 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   300
         MaxLength       =   20
         TabIndex        =   19
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblHeadings 
         Alignment       =   2  'Center
         Caption         =   "[ ---- Attack Spell / Hit Spell ---- ]"
         Height          =   195
         Index           =   25
         Left            =   7500
         TabIndex        =   7
         Top             =   120
         Width           =   2355
      End
      Begin VB.Label lblHeadings 
         Alignment       =   2  'Center
         Caption         =   "dmg -MR?"
         Height          =   375
         Index           =   24
         Left            =   9300
         TabIndex        =   18
         Top             =   300
         Width           =   555
      End
      Begin VB.Label lblHeadings 
         Alignment       =   2  'Center
         Caption         =   "Resist All"
         Height          =   255
         Index           =   2
         Left            =   7620
         TabIndex        =   16
         Top             =   420
         Width           =   795
      End
      Begin VB.Label lblHeadings 
         Alignment       =   2  'Center
         Caption         =   "Duration"
         Height          =   255
         Index           =   6
         Left            =   8580
         TabIndex        =   17
         Top             =   420
         Width           =   675
      End
      Begin VB.Label lblHitSpell 
         Alignment       =   2  'Center
         Caption         =   "Hit Spell Min - Max"
         Height          =   435
         Left            =   6480
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblHeadings 
         Caption         =   "5."
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
         Index           =   12
         Left            =   60
         TabIndex        =   73
         Top             =   2430
         Width           =   255
      End
      Begin VB.Label lblHeadings 
         Caption         =   "4."
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
         Index           =   11
         Left            =   60
         TabIndex        =   60
         Top             =   2010
         Width           =   255
      End
      Begin VB.Label lblHeadings 
         Caption         =   "3."
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
         Index           =   10
         Left            =   60
         TabIndex        =   47
         Top             =   1590
         Width           =   255
      End
      Begin VB.Label lblHeadings 
         Caption         =   "2."
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
         Index           =   9
         Left            =   60
         TabIndex        =   42
         Top             =   1170
         Width           =   255
      End
      Begin VB.Label lblHeadings 
         Caption         =   "1."
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
         Index           =   8
         Left            =   60
         TabIndex        =   21
         Top             =   750
         Width           =   255
      End
      Begin VB.Label lblHeadings 
         Alignment       =   2  'Center
         Caption         =   "Accy. or Cast%"
         Height          =   435
         Index           =   7
         Left            =   4080
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblAtkChance 
         Alignment       =   2  'Center
         Caption         =   "Attack Chance%"
         Height          =   435
         Left            =   2460
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblHeadings 
         Alignment       =   2  'Center
         Caption         =   "Max"
         Height          =   255
         Index           =   5
         Left            =   5700
         TabIndex        =   14
         Top             =   420
         Width           =   615
      End
      Begin VB.Label lblHeadings 
         Alignment       =   2  'Center
         Caption         =   "Min"
         Height          =   255
         Index           =   4
         Left            =   4860
         TabIndex        =   13
         Top             =   420
         Width           =   735
      End
      Begin VB.Label lblHeadings 
         Alignment       =   2  'Center
         Caption         =   "Energy"
         Height          =   255
         Index           =   3
         Left            =   3300
         TabIndex        =   11
         Top             =   420
         Width           =   735
      End
      Begin VB.Label lblHeadings 
         Alignment       =   2  'Center
         Caption         =   "Type"
         Height          =   255
         Index           =   1
         Left            =   1500
         TabIndex        =   9
         Top             =   420
         Width           =   915
      End
      Begin VB.Label lblHeadings 
         Alignment       =   2  'Center
         Caption         =   "Name"
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   8
         Top             =   420
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   10080
      TabIndex        =   84
      Top             =   180
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   5741
      _Version        =   393216
      Tab             =   1
      TabHeight       =   882
      TabCaption(0)   =   "Between Round Spells"
      TabPicture(0)   =   "frmMonsterAttackSim.frx":0A96
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdBetweenRoundSpellGoto(4)"
      Tab(0).Control(1)=   "cmdBetweenRoundSpellGoto(3)"
      Tab(0).Control(2)=   "cmdBetweenRoundSpellGoto(2)"
      Tab(0).Control(3)=   "cmdBetweenRoundSpellGoto(1)"
      Tab(0).Control(4)=   "cmdBetweenRoundSpellGoto(0)"
      Tab(0).Control(5)=   "txtBetweenSpellNumber(0)"
      Tab(0).Control(6)=   "txtBetweenSpellName(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtBetweenSpellCastPer(0)"
      Tab(0).Control(8)=   "txtBetweenSpellCastLvL(0)"
      Tab(0).Control(9)=   "txtBetweenSpellNumber(1)"
      Tab(0).Control(10)=   "txtBetweenSpellName(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtBetweenSpellCastPer(1)"
      Tab(0).Control(12)=   "txtBetweenSpellCastLvL(1)"
      Tab(0).Control(13)=   "txtBetweenSpellNumber(2)"
      Tab(0).Control(14)=   "txtBetweenSpellName(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtBetweenSpellCastPer(2)"
      Tab(0).Control(16)=   "txtBetweenSpellCastLvL(2)"
      Tab(0).Control(17)=   "txtBetweenSpellNumber(3)"
      Tab(0).Control(18)=   "txtBetweenSpellName(3)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtBetweenSpellCastPer(3)"
      Tab(0).Control(20)=   "txtBetweenSpellCastLvL(3)"
      Tab(0).Control(21)=   "txtBetweenSpellNumber(4)"
      Tab(0).Control(22)=   "txtBetweenSpellName(4)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtBetweenSpellCastPer(4)"
      Tab(0).Control(24)=   "txtBetweenSpellCastLvL(4)"
      Tab(0).Control(25)=   "label(31)"
      Tab(0).Control(26)=   "label(32)"
      Tab(0).Control(27)=   "label(33)"
      Tab(0).Control(28)=   "label(34)"
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "Attack Statistics"
      TabPicture(1)   =   "frmMonsterAttackSim.frx":0AB2
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblHeadings(23)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblHeadings(22)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblHeadings(21)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblHeadings(20)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblHeadings(19)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblHeadings(18)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblHeadings(17)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblHeadings(16)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblHeadings(15)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblHeadings(14)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lblHeadings(13)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtStatResistDodge(4)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtStatDmgResist(4)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtStatSuccess(4)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtStatAvgRound(4)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtStatAttRound(4)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtStatTrueCast(4)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtStatResistDodge(3)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtStatDmgResist(3)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtStatSuccess(3)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtStatAvgRound(3)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtStatAttRound(3)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtStatTrueCast(3)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtStatResistDodge(2)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "txtStatDmgResist(2)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "txtStatSuccess(2)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "txtStatAvgRound(2)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "txtStatAttRound(2)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "txtStatTrueCast(2)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "txtStatResistDodge(1)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "txtStatDmgResist(1)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "txtStatSuccess(1)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "txtStatAvgRound(1)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "txtStatAttRound(1)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "txtStatTrueCast(1)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "txtStatResistDodge(0)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "txtStatDmgResist(0)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "txtStatSuccess(0)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "txtStatAvgRound(0)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "txtStatAttRound(0)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "txtStatTrueCast(0)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).ControlCount=   41
      TabCaption(2)   =   "Items"
      TabPicture(2)   =   "frmMonsterAttackSim.frx":0ACE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtItemDropPer(5)"
      Tab(2).Control(1)=   "txtItemDropPer(7)"
      Tab(2).Control(2)=   "txtItemDropPer(8)"
      Tab(2).Control(3)=   "txtItemDropPer(9)"
      Tab(2).Control(4)=   "txtItemDropPer(6)"
      Tab(2).Control(5)=   "txtItemDropPer(2)"
      Tab(2).Control(6)=   "txtItemDropPer(3)"
      Tab(2).Control(7)=   "txtItemDropPer(4)"
      Tab(2).Control(8)=   "txtItemDropPer(0)"
      Tab(2).Control(9)=   "txtItemDropPer(1)"
      Tab(2).Control(10)=   "txtItemNumber(9)"
      Tab(2).Control(11)=   "txtItemName(9)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "txtItemNumber(8)"
      Tab(2).Control(13)=   "txtItemName(8)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "txtItemNumber(7)"
      Tab(2).Control(15)=   "txtItemName(7)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "txtItemName(5)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "txtItemNumber(6)"
      Tab(2).Control(18)=   "txtItemName(6)"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "txtItemNumber(5)"
      Tab(2).Control(20)=   "cmdEditItemDrop(5)"
      Tab(2).Control(21)=   "cmdEditItemDrop(6)"
      Tab(2).Control(22)=   "cmdEditItemDrop(7)"
      Tab(2).Control(23)=   "cmdEditItemDrop(8)"
      Tab(2).Control(24)=   "cmdEditItemDrop(9)"
      Tab(2).Control(25)=   "txtItemName(4)"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "txtItemNumber(4)"
      Tab(2).Control(27)=   "txtItemName(3)"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "txtItemNumber(3)"
      Tab(2).Control(29)=   "txtItemName(2)"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "txtItemNumber(2)"
      Tab(2).Control(31)=   "txtItemName(0)"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "txtItemNumber(0)"
      Tab(2).Control(33)=   "txtItemName(1)"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "txtItemNumber(1)"
      Tab(2).Control(35)=   "cmdEditItemDrop(0)"
      Tab(2).Control(36)=   "cmdEditItemDrop(1)"
      Tab(2).Control(37)=   "cmdEditItemDrop(2)"
      Tab(2).Control(38)=   "cmdEditItemDrop(3)"
      Tab(2).Control(39)=   "cmdEditItemDrop(4)"
      Tab(2).Control(40)=   "cmdGotoWeapon"
      Tab(2).Control(41)=   "txtWeaponName"
      Tab(2).Control(41).Enabled=   0   'False
      Tab(2).Control(42)=   "txtWeaponNumber"
      Tab(2).Control(43)=   "cmdItemNote"
      Tab(2).Control(44)=   "lblItemBonus"
      Tab(2).Control(45)=   "Label5"
      Tab(2).Control(46)=   "Label4"
      Tab(2).Control(47)=   "Label3"
      Tab(2).Control(48)=   "label(17)"
      Tab(2).ControlCount=   49
      Begin VB.TextBox txtItemDropPer 
         Height          =   285
         Index           =   5
         Left            =   -71100
         TabIndex        =   122
         Top             =   1260
         Width           =   495
      End
      Begin VB.TextBox txtItemDropPer 
         Height          =   285
         Index           =   7
         Left            =   -71100
         TabIndex        =   155
         Top             =   1830
         Width           =   495
      End
      Begin VB.TextBox txtItemDropPer 
         Height          =   285
         Index           =   8
         Left            =   -71100
         TabIndex        =   170
         Top             =   2115
         Width           =   495
      End
      Begin VB.TextBox txtItemDropPer 
         Height          =   285
         Index           =   9
         Left            =   -71100
         TabIndex        =   183
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txtItemDropPer 
         Height          =   285
         Index           =   6
         Left            =   -71100
         TabIndex        =   135
         Top             =   1545
         Width           =   495
      End
      Begin VB.TextBox txtItemDropPer 
         Height          =   285
         Index           =   2
         Left            =   -73200
         TabIndex        =   149
         Top             =   1830
         Width           =   435
      End
      Begin VB.TextBox txtItemDropPer 
         Height          =   285
         Index           =   3
         Left            =   -73200
         TabIndex        =   166
         Top             =   2115
         Width           =   435
      End
      Begin VB.TextBox txtItemDropPer 
         Height          =   285
         Index           =   4
         Left            =   -73200
         TabIndex        =   179
         Top             =   2400
         Width           =   435
      End
      Begin VB.TextBox txtItemDropPer 
         Height          =   285
         Index           =   0
         Left            =   -73200
         TabIndex        =   119
         Top             =   1260
         Width           =   435
      End
      Begin VB.TextBox txtItemDropPer 
         Height          =   285
         Index           =   1
         Left            =   -73200
         TabIndex        =   131
         Top             =   1545
         Width           =   435
      End
      Begin VB.TextBox txtItemNumber 
         Height          =   285
         Index           =   9
         Left            =   -72480
         TabIndex        =   181
         Top             =   2400
         Width           =   555
      End
      Begin VB.TextBox txtItemName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   9
         Left            =   -71925
         Locked          =   -1  'True
         TabIndex        =   182
         TabStop         =   0   'False
         Top             =   2400
         Width           =   795
      End
      Begin VB.TextBox txtItemNumber 
         Height          =   285
         Index           =   8
         Left            =   -72480
         TabIndex        =   168
         Top             =   2115
         Width           =   555
      End
      Begin VB.TextBox txtItemName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   8
         Left            =   -71925
         Locked          =   -1  'True
         TabIndex        =   169
         TabStop         =   0   'False
         Top             =   2115
         Width           =   795
      End
      Begin VB.TextBox txtItemNumber 
         Height          =   285
         Index           =   7
         Left            =   -72480
         TabIndex        =   151
         Top             =   1830
         Width           =   555
      End
      Begin VB.TextBox txtItemName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   7
         Left            =   -71925
         Locked          =   -1  'True
         TabIndex        =   154
         TabStop         =   0   'False
         Top             =   1830
         Width           =   795
      End
      Begin VB.TextBox txtItemName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   5
         Left            =   -71925
         Locked          =   -1  'True
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   1260
         Width           =   795
      End
      Begin VB.TextBox txtItemNumber 
         Height          =   285
         Index           =   6
         Left            =   -72480
         TabIndex        =   133
         Top             =   1545
         Width           =   555
      End
      Begin VB.TextBox txtItemName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   6
         Left            =   -71925
         Locked          =   -1  'True
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   1545
         Width           =   795
      End
      Begin VB.TextBox txtItemNumber 
         Height          =   285
         Index           =   5
         Left            =   -72480
         TabIndex        =   120
         Top             =   1260
         Width           =   555
      End
      Begin VB.CommandButton cmdEditItemDrop 
         Height          =   195
         Index           =   5
         Left            =   -72720
         TabIndex        =   109
         Top             =   1260
         Width           =   195
      End
      Begin VB.CommandButton cmdEditItemDrop 
         Height          =   195
         Index           =   6
         Left            =   -72720
         TabIndex        =   132
         Top             =   1545
         Width           =   195
      End
      Begin VB.CommandButton cmdEditItemDrop 
         Height          =   195
         Index           =   7
         Left            =   -72720
         TabIndex        =   150
         Top             =   1830
         Width           =   195
      End
      Begin VB.CommandButton cmdEditItemDrop 
         Height          =   195
         Index           =   8
         Left            =   -72720
         TabIndex        =   167
         Top             =   2115
         Width           =   195
      End
      Begin VB.CommandButton cmdEditItemDrop 
         Height          =   195
         Index           =   9
         Left            =   -72720
         TabIndex        =   180
         Top             =   2400
         Width           =   195
      End
      Begin VB.TextBox txtItemName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   4
         Left            =   -74085
         Locked          =   -1  'True
         TabIndex        =   178
         TabStop         =   0   'False
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtItemNumber 
         Height          =   285
         Index           =   4
         Left            =   -74640
         TabIndex        =   177
         Top             =   2400
         Width           =   555
      End
      Begin VB.TextBox txtItemName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   -74085
         Locked          =   -1  'True
         TabIndex        =   165
         TabStop         =   0   'False
         Top             =   2115
         Width           =   855
      End
      Begin VB.TextBox txtItemNumber 
         Height          =   285
         Index           =   3
         Left            =   -74640
         TabIndex        =   164
         Top             =   2115
         Width           =   555
      End
      Begin VB.TextBox txtItemName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   -74085
         Locked          =   -1  'True
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   1830
         Width           =   855
      End
      Begin VB.TextBox txtItemNumber 
         Height          =   285
         Index           =   2
         Left            =   -74640
         TabIndex        =   145
         Top             =   1830
         Width           =   555
      End
      Begin VB.TextBox txtItemName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   -74085
         Locked          =   -1  'True
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   1260
         Width           =   855
      End
      Begin VB.TextBox txtItemNumber 
         Height          =   285
         Index           =   0
         Left            =   -74640
         TabIndex        =   117
         Top             =   1260
         Width           =   555
      End
      Begin VB.TextBox txtItemName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   -74085
         Locked          =   -1  'True
         TabIndex        =   130
         TabStop         =   0   'False
         Top             =   1545
         Width           =   855
      End
      Begin VB.TextBox txtItemNumber 
         Height          =   285
         Index           =   1
         Left            =   -74640
         TabIndex        =   129
         Top             =   1545
         Width           =   555
      End
      Begin VB.CommandButton cmdEditItemDrop 
         Height          =   195
         Index           =   0
         Left            =   -74880
         TabIndex        =   108
         Top             =   1260
         Width           =   195
      End
      Begin VB.CommandButton cmdEditItemDrop 
         Height          =   195
         Index           =   1
         Left            =   -74880
         TabIndex        =   124
         Top             =   1545
         Width           =   195
      End
      Begin VB.CommandButton cmdEditItemDrop 
         Height          =   195
         Index           =   2
         Left            =   -74880
         TabIndex        =   143
         Top             =   1830
         Width           =   195
      End
      Begin VB.CommandButton cmdEditItemDrop 
         Height          =   195
         Index           =   3
         Left            =   -74880
         TabIndex        =   162
         Top             =   2115
         Width           =   195
      End
      Begin VB.CommandButton cmdEditItemDrop 
         Height          =   195
         Index           =   4
         Left            =   -74880
         TabIndex        =   172
         Top             =   2400
         Width           =   195
      End
      Begin VB.CommandButton cmdGotoWeapon 
         Height          =   195
         Left            =   -74880
         TabIndex        =   85
         Top             =   660
         Width           =   195
      End
      Begin VB.TextBox txtWeaponName 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   -73200
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   660
         Width           =   1755
      End
      Begin VB.TextBox txtWeaponNumber 
         Height          =   285
         Left            =   -73860
         TabIndex        =   87
         Top             =   660
         Width           =   615
      End
      Begin VB.CommandButton cmdItemNote 
         Caption         =   "Note"
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
         Left            =   -71340
         TabIndex        =   89
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdBetweenRoundSpellGoto 
         Height          =   195
         Index           =   4
         Left            =   -74820
         TabIndex        =   191
         Top             =   2760
         Width           =   195
      End
      Begin VB.CommandButton cmdBetweenRoundSpellGoto 
         Height          =   195
         Index           =   3
         Left            =   -74820
         TabIndex        =   171
         Top             =   2340
         Width           =   195
      End
      Begin VB.CommandButton cmdBetweenRoundSpellGoto 
         Height          =   195
         Index           =   2
         Left            =   -74820
         TabIndex        =   146
         Top             =   1920
         Width           =   195
      End
      Begin VB.CommandButton cmdBetweenRoundSpellGoto 
         Height          =   195
         Index           =   1
         Left            =   -74820
         TabIndex        =   123
         Top             =   1500
         Width           =   195
      End
      Begin VB.CommandButton cmdBetweenRoundSpellGoto 
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   102
         Top             =   1080
         Width           =   195
      End
      Begin VB.TextBox txtBetweenSpellNumber 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   -74460
         TabIndex        =   103
         Top             =   975
         Width           =   615
      End
      Begin VB.TextBox txtBetweenSpellName 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   -73740
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   975
         Width           =   1815
      End
      Begin VB.TextBox txtBetweenSpellCastPer 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   -71820
         TabIndex        =   105
         Text            =   "99"
         Top             =   975
         Width           =   435
      End
      Begin VB.TextBox txtBetweenSpellCastLvL 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   -71280
         TabIndex        =   106
         Top             =   975
         Width           =   615
      End
      Begin VB.TextBox txtBetweenSpellNumber 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   -74460
         TabIndex        =   125
         Top             =   1395
         Width           =   615
      End
      Begin VB.TextBox txtBetweenSpellName 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   -73740
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   1395
         Width           =   1815
      End
      Begin VB.TextBox txtBetweenSpellCastPer 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   -71820
         TabIndex        =   127
         Top             =   1395
         Width           =   435
      End
      Begin VB.TextBox txtBetweenSpellCastLvL 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   -71280
         TabIndex        =   128
         Top             =   1395
         Width           =   615
      End
      Begin VB.TextBox txtBetweenSpellNumber 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   -74460
         TabIndex        =   144
         Top             =   1815
         Width           =   615
      End
      Begin VB.TextBox txtBetweenSpellName 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   -73740
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   147
         TabStop         =   0   'False
         Top             =   1815
         Width           =   1815
      End
      Begin VB.TextBox txtBetweenSpellCastPer 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   -71820
         TabIndex        =   152
         Top             =   1815
         Width           =   435
      End
      Begin VB.TextBox txtBetweenSpellCastLvL 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   -71280
         TabIndex        =   153
         Top             =   1815
         Width           =   615
      End
      Begin VB.TextBox txtBetweenSpellNumber 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   -74460
         TabIndex        =   173
         Top             =   2235
         Width           =   615
      End
      Begin VB.TextBox txtBetweenSpellName 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   -73740
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   174
         TabStop         =   0   'False
         Top             =   2235
         Width           =   1815
      End
      Begin VB.TextBox txtBetweenSpellCastPer 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   -71820
         TabIndex        =   175
         Top             =   2235
         Width           =   435
      End
      Begin VB.TextBox txtBetweenSpellCastLvL 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   -71280
         TabIndex        =   176
         Top             =   2235
         Width           =   615
      End
      Begin VB.TextBox txtBetweenSpellNumber 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   -74460
         TabIndex        =   192
         Top             =   2655
         Width           =   615
      End
      Begin VB.TextBox txtBetweenSpellName 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   -73740
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   193
         TabStop         =   0   'False
         Top             =   2655
         Width           =   1815
      End
      Begin VB.TextBox txtBetweenSpellCastPer 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   -71820
         TabIndex        =   194
         Top             =   2655
         Width           =   435
      End
      Begin VB.TextBox txtBetweenSpellCastLvL 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   -71280
         TabIndex        =   195
         Top             =   2655
         Width           =   615
      End
      Begin VB.TextBox txtStatTrueCast 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   285
         Locked          =   -1  'True
         TabIndex        =   110
         Text            =   "45.5%"
         Top             =   1095
         Width           =   675
      End
      Begin VB.TextBox txtStatAttRound 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   111
         Text            =   "5.05"
         Top             =   1095
         Width           =   675
      End
      Begin VB.TextBox txtStatAvgRound 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   112
         Text            =   "99999"
         Top             =   1095
         Width           =   675
      End
      Begin VB.TextBox txtStatSuccess 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   113
         Text            =   "100%"
         Top             =   1095
         Width           =   615
      End
      Begin VB.TextBox txtStatDmgResist 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   114
         Text            =   "100%"
         Top             =   1095
         Width           =   615
      End
      Begin VB.TextBox txtStatResistDodge 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   115
         Text            =   "100%"
         Top             =   1095
         Width           =   615
      End
      Begin VB.TextBox txtStatTrueCast 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   285
         Locked          =   -1  'True
         TabIndex        =   136
         Top             =   1515
         Width           =   675
      End
      Begin VB.TextBox txtStatAttRound 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   137
         Top             =   1515
         Width           =   675
      End
      Begin VB.TextBox txtStatAvgRound 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   138
         Top             =   1515
         Width           =   675
      End
      Begin VB.TextBox txtStatSuccess 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   139
         Top             =   1515
         Width           =   615
      End
      Begin VB.TextBox txtStatDmgResist 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   140
         Top             =   1515
         Width           =   615
      End
      Begin VB.TextBox txtStatResistDodge 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   141
         Top             =   1515
         Width           =   615
      End
      Begin VB.TextBox txtStatTrueCast 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   285
         Locked          =   -1  'True
         TabIndex        =   156
         Top             =   1935
         Width           =   675
      End
      Begin VB.TextBox txtStatAttRound 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   157
         Top             =   1935
         Width           =   675
      End
      Begin VB.TextBox txtStatAvgRound 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   158
         Top             =   1935
         Width           =   675
      End
      Begin VB.TextBox txtStatSuccess 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   159
         Top             =   1935
         Width           =   615
      End
      Begin VB.TextBox txtStatDmgResist 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   160
         Top             =   1935
         Width           =   615
      End
      Begin VB.TextBox txtStatResistDodge 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   161
         Top             =   1935
         Width           =   615
      End
      Begin VB.TextBox txtStatTrueCast 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   285
         Locked          =   -1  'True
         TabIndex        =   184
         Top             =   2355
         Width           =   675
      End
      Begin VB.TextBox txtStatAttRound 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   185
         Top             =   2355
         Width           =   675
      End
      Begin VB.TextBox txtStatAvgRound 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   186
         Top             =   2355
         Width           =   675
      End
      Begin VB.TextBox txtStatSuccess 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   187
         Top             =   2355
         Width           =   615
      End
      Begin VB.TextBox txtStatDmgResist 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   188
         Top             =   2355
         Width           =   615
      End
      Begin VB.TextBox txtStatResistDodge 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   189
         Top             =   2355
         Width           =   615
      End
      Begin VB.TextBox txtStatTrueCast 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   285
         Locked          =   -1  'True
         TabIndex        =   197
         Top             =   2775
         Width           =   675
      End
      Begin VB.TextBox txtStatAttRound 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   198
         Top             =   2775
         Width           =   675
      End
      Begin VB.TextBox txtStatAvgRound 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   199
         Top             =   2775
         Width           =   675
      End
      Begin VB.TextBox txtStatSuccess 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   200
         Top             =   2775
         Width           =   615
      End
      Begin VB.TextBox txtStatDmgResist 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   201
         Top             =   2775
         Width           =   615
      End
      Begin VB.TextBox txtStatResistDodge 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   202
         Top             =   2775
         Width           =   615
      End
      Begin VB.Label lblItemBonus 
         Height          =   435
         Left            =   -74820
         TabIndex        =   196
         Top             =   2700
         Width           =   4215
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "%"
         Height          =   195
         Left            =   -70980
         TabIndex        =   107
         Top             =   1020
         Width           =   315
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "%"
         Height          =   195
         Left            =   -73140
         TabIndex        =   101
         Top             =   1020
         Width           =   315
      End
      Begin VB.Label Label3 
         Caption         =   "Drops:"
         Height          =   195
         Left            =   -74640
         TabIndex        =   100
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label label 
         Caption         =   "Weapon"
         Height          =   255
         Index           =   17
         Left            =   -74580
         TabIndex        =   86
         Top             =   660
         Width           =   675
      End
      Begin VB.Label label 
         Alignment       =   2  'Center
         Caption         =   "Cast LVL"
         Height          =   255
         Index           =   31
         Left            =   -71340
         TabIndex        =   93
         Top             =   735
         Width           =   735
      End
      Begin VB.Label label 
         Alignment       =   2  'Center
         Caption         =   "%"
         Height          =   255
         Index           =   32
         Left            =   -71820
         TabIndex        =   92
         Top             =   735
         Width           =   435
      End
      Begin VB.Label label 
         Alignment       =   2  'Center
         Caption         =   "Name"
         Height          =   255
         Index           =   33
         Left            =   -73680
         TabIndex        =   91
         Top             =   735
         Width           =   1695
      End
      Begin VB.Label label 
         Alignment       =   2  'Center
         Caption         =   "#"
         Height          =   255
         Index           =   34
         Left            =   -74460
         TabIndex        =   90
         Top             =   735
         Width           =   615
      End
      Begin VB.Label lblHeadings 
         Alignment       =   2  'Center
         Caption         =   "True Attk%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   13
         Left            =   345
         TabIndex        =   94
         Top             =   675
         Width           =   615
      End
      Begin VB.Label lblHeadings 
         Caption         =   "Attempt /Round"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   14
         Left            =   1005
         TabIndex        =   95
         Top             =   675
         Width           =   735
      End
      Begin VB.Label lblHeadings 
         Alignment       =   2  'Center
         Caption         =   "Avg Round"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   15
         Left            =   1785
         TabIndex        =   96
         Top             =   675
         Width           =   555
      End
      Begin VB.Label lblHeadings 
         Alignment       =   2  'Center
         Caption         =   "% Hit"
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
         Index           =   16
         Left            =   2445
         TabIndex        =   97
         Top             =   855
         Width           =   615
      End
      Begin VB.Label lblHeadings 
         Alignment       =   2  'Center
         Caption         =   "%dmg Resist"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   17
         Left            =   3105
         TabIndex        =   98
         Top             =   675
         Width           =   615
      End
      Begin VB.Label lblHeadings 
         Alignment       =   2  'Center
         Caption         =   "%resist /dodge"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   18
         Left            =   3705
         TabIndex        =   99
         Top             =   675
         Width           =   735
      End
      Begin VB.Label lblHeadings 
         Caption         =   "5"
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
         Index           =   19
         Left            =   120
         TabIndex        =   203
         Top             =   2835
         Width           =   255
      End
      Begin VB.Label lblHeadings 
         Caption         =   "4"
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
         Index           =   20
         Left            =   120
         TabIndex        =   190
         Top             =   2415
         Width           =   255
      End
      Begin VB.Label lblHeadings 
         Caption         =   "3"
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
         Index           =   21
         Left            =   120
         TabIndex        =   163
         Top             =   1995
         Width           =   255
      End
      Begin VB.Label lblHeadings 
         Caption         =   "2"
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
         Index           =   22
         Left            =   120
         TabIndex        =   142
         Top             =   1575
         Width           =   255
      End
      Begin VB.Label lblHeadings 
         Caption         =   "1"
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
         Index           =   23
         Left            =   120
         TabIndex        =   116
         Top             =   1155
         Width           =   255
      End
   End
   Begin VB.Frame fraResults 
      Caption         =   "Results"
      Height          =   1575
      Left            =   9600
      TabIndex        =   216
      Top             =   4620
      Width           =   4995
      Begin VB.Label lblResultsMaxRound 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   180
         TabIndex        =   218
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label lblResultsAvgDmg 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   180
         TabIndex        =   217
         Top             =   300
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdSim 
      Caption         =   "Run Sim"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9660
      TabIndex        =   222
      Top             =   6780
      Width           =   4875
   End
   Begin VB.Frame fraChar 
      Caption         =   "Character Defenses"
      Height          =   975
      Left            =   9600
      TabIndex        =   205
      Top             =   3540
      Width           =   4995
      Begin VB.TextBox txtUserDR 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1140
         TabIndex        =   212
         Top             =   480
         Width           =   735
      End
      Begin VB.CheckBox chkUserAntiMagic 
         Height          =   255
         Left            =   4020
         TabIndex        =   215
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox txtUserMR 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2820
         TabIndex        =   214
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtUserDodge 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1980
         TabIndex        =   213
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtUserAC 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   300
         TabIndex        =   211
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblAntiMagic 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Anti-Magic"
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
         Left            =   3570
         TabIndex        =   210
         Top             =   240
         Width           =   1155
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "DR"
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
         Index           =   4
         Left            =   1230
         TabIndex        =   207
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "MR"
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
         Index           =   3
         Left            =   2850
         TabIndex        =   209
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "AC"
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
         Index           =   2
         Left            =   360
         TabIndex        =   206
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Dodge%"
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
         Index           =   1
         Left            =   1980
         TabIndex        =   208
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.TextBox txtNumRounds 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "M/dd/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11640
      TabIndex        =   220
      Text            =   "2000"
      Top             =   6300
      Width           =   915
   End
   Begin VB.TextBox txtMonsterEnergy 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "M/dd/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9240
      TabIndex        =   5
      Text            =   "1000"
      Top             =   120
      Width           =   675
   End
   Begin VB.TextBox txtCombatLog 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   204
      Top             =   3540
      Width           =   9435
   End
   Begin VB.CommandButton cmbRefresh 
      Caption         =   "Refresh List"
      Height          =   375
      Left            =   5580
      TabIndex        =   3
      Top             =   60
      Width           =   1155
   End
   Begin VB.ComboBox cmbMonster 
      Height          =   315
      ItemData        =   "frmMonsterAttackSim.frx":0AEA
      Left            =   1620
      List            =   "frmMonsterAttackSim.frx":0AEC
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   60
      TabIndex        =   223
      Top             =   7380
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "# Rounds to Sim:"
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
      Left            =   9840
      TabIndex        =   219
      Top             =   6360
      Width           =   1605
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Monster's Energy/Round:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   6840
      TabIndex        =   4
      Top             =   180
      Width           =   2355
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMonster 
      AutoSize        =   -1  'True
      Caption         =   "Choose Monster:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1440
   End
End
Attribute VB_Name = "frmMonsterAttackSim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim clsMonAtkSim As New clsMonsterAttackSim

Private Sub chkDynamicRounds_Click()
If chkDynamicRounds.Value = 1 Then
    txtNumRounds.Enabled = False
    txtNumRounds.BackColor = &H8000000F
Else
    txtNumRounds.BackColor = &H80000005
    txtNumRounds.Enabled = True
End If
End Sub

Private Sub cmbAtkType_Click(Index As Integer)
On Error GoTo error:

cmbAtkResist(Index).Enabled = True
chkAtkDmgResist(Index).Enabled = True
txtAtkHitSpellMin(Index).Enabled = True
txtAtkHitSpellMax(Index).Enabled = True

If cmbAtkType(Index).ListIndex = 2 Then
    
    txtAtkHitSpellMin(Index).Text = ""
    txtAtkHitSpellMax(Index).Text = ""
    
    txtAtkHitSpellMin(Index).Enabled = False
    txtAtkHitSpellMax(Index).Enabled = False
    
ElseIf cmbAtkType(Index).ListIndex = 0 Then
    
    txtAtkHitSpellMin(Index).Text = ""
    txtAtkHitSpellMax(Index).Text = ""
    cmbAtkResist(Index).ListIndex = 0
    chkAtkDmgResist(Index).Value = 0
    
    cmbAtkResist(Index).Enabled = False
    chkAtkDmgResist(Index).Enabled = False
    txtAtkHitSpellMin(Index).Enabled = False
    txtAtkHitSpellMax(Index).Enabled = False
    
End If

If txtAtkHitSpellMin(Index).Enabled Then
    txtAtkHitSpellMin(Index).BackColor = &H80000005
Else
    txtAtkHitSpellMin(Index).BackColor = &H8000000F
End If

If txtAtkHitSpellMax(Index).Enabled Then
    txtAtkHitSpellMax(Index).BackColor = &H80000005
Else
    txtAtkHitSpellMax(Index).BackColor = &H8000000F
End If

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("cmbAtkType_Click")
Resume out:
End Sub

Private Sub cmbMonster_Click()
If cmbMonster.ListIndex < 0 Then Exit Sub
Call PopulateAttacks(cmbMonster.ItemData(cmbMonster.ListIndex))
End Sub

Public Sub PopulateAttacks(nMonster As Long)
Dim nStatus As Integer, x As Integer, y As Integer
Dim nPercent As Integer, sTemp As String, nTest As Integer
On Error GoTo error:

nStatus = BTRCALL(BGETEQUAL, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), nMonster, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error on BGETEQUAL: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If
Call MonsterRowToStruct(Monsterdatabuf.buf)

Call ResetMonsterFields

txtMonsterEnergy.Text = Monsterrec.Energy

If Monsterrec.WeaponNumber > 0 Then
    If GetItemLimit(Monsterrec.WeaponNumber) = 0 Then
        '4 = max dam ... accy = 22, 105, 106
        If ItemHasAbility(Monsterrec.WeaponNumber, 22) > 0 _
            Or ItemHasAbility(Monsterrec.WeaponNumber, 105) > 0 _
            Or ItemHasAbility(Monsterrec.WeaponNumber, 106) > 0 _
            Or ItemHasAbility(Monsterrec.WeaponNumber, 4) > 0 Then
            
            txtWeaponNumber.Text = Monsterrec.WeaponNumber
        End If
    End If
End If

For x = 0 To 9
    If Monsterrec.ItemNumber(x) > 0 Then
        If GetItemLimit(Monsterrec.ItemNumber(x)) = 0 Then
            '4 = max dam ... accy = 22, 105, 106
            If ItemHasAbility(Monsterrec.ItemNumber(x), 22) > 0 _
                Or ItemHasAbility(Monsterrec.ItemNumber(x), 105) > 0 _
                Or ItemHasAbility(Monsterrec.ItemNumber(x), 106) > 0 _
                Or ItemHasAbility(Monsterrec.ItemNumber(x), 4) > 0 Then
                
                txtItemNumber(x).Text = Monsterrec.ItemNumber(x)
                txtItemDropPer(x).Text = Monsterrec.ItemDropPer(x)
            End If
        End If
    End If
Next x

nPercent = 0
For x = 0 To 4
    If Monsterrec.AttackType(x) > 0 And Monsterrec.AttackType(x) < 4 Then
        sTemp = GetMonsterAttackName(Monsterrec.Number, x, 20)
        If InStr(1, sTemp, " you ", vbTextCompare) Then
            sTemp = Mid(sTemp, 1, InStr(1, sTemp, " you ", vbTextCompare))
        ElseIf Right(sTemp, 4) = " you" Then
            sTemp = Left(sTemp, Len(sTemp) - 4)
        End If
        txtAtkName(x).Text = sTemp
        cmbAtkType(x).ListIndex = Monsterrec.AttackType(x)
        txtAtkEnergy(x).Text = Monsterrec.AttackEnergy(x)
        
        txtAtkChance(x).Text = Monsterrec.AttackPer(x) - nPercent
        nPercent = Monsterrec.AttackPer(x)
        
        If Monsterrec.AttackType(x) = 2 Then 'spell
            nStatus = GetSpell(Monsterrec.AttackAccuSpell(x))
            If nStatus = 0 Then
                cmbAtkResist(x).ListIndex = Spellrec.TypeOfResists
                
                txtAtkDur(x).Text = GetSpellDuration(Monsterrec.AttackAccuSpell(x), Monsterrec.AttackMaxHCastLvl(x))
                txtAtkMin(x).Text = 0
                txtAtkMax(x).Text = 0
                
                nTest = SpellHasAbility(Monsterrec.AttackAccuSpell(x), 1) '1=damage
                If nTest >= 0 Then
                    chkAtkDmgResist(x).Value = 0 'NO MR resist
                    If nTest > 0 Then
                        txtAtkMin(x).Text = nTest
                        txtAtkMax(x).Text = nTest
                    Else
                        txtAtkMin(x).Text = GetSpellMinDamage(Monsterrec.AttackAccuSpell(x), Monsterrec.AttackMaxHCastLvl(x))
                        txtAtkMax(x).Text = GetSpellMaxDamage(Monsterrec.AttackAccuSpell(x), Monsterrec.AttackMaxHCastLvl(x))
                    End If
                End If
                
                nTest = SpellHasAbility(Monsterrec.AttackAccuSpell(x), 17) '17=damage
                If nTest >= 0 Then
                    chkAtkDmgResist(x).Value = 1 'MR resist
                    If nTest > 0 Then
                        txtAtkMin(x).Text = nTest
                        txtAtkMax(x).Text = nTest
                    Else
                        txtAtkMin(x).Text = GetSpellMinDamage(Monsterrec.AttackAccuSpell(x), Monsterrec.AttackMaxHCastLvl(x))
                        txtAtkMax(x).Text = GetSpellMaxDamage(Monsterrec.AttackAccuSpell(x), Monsterrec.AttackMaxHCastLvl(x))
                    End If
                End If
                
                nTest = SpellHasAbility(Monsterrec.AttackAccuSpell(x), 8) '8=drain
                If nTest >= 0 Then
                    chkAtkDmgResist(x).Value = 0 'NO MR resist
                    If nTest > 0 Then
                        txtAtkMin(x).Text = nTest
                        txtAtkMax(x).Text = nTest
                    Else
                        txtAtkMin(x).Text = GetSpellMinDamage(Monsterrec.AttackAccuSpell(x), Monsterrec.AttackMaxHCastLvl(x))
                        txtAtkMax(x).Text = GetSpellMaxDamage(Monsterrec.AttackAccuSpell(x), Monsterrec.AttackMaxHCastLvl(x))
                    End If
                End If
                
            Else
                txtAtkMin(x).Text = "!"
                txtAtkMax(x).Text = "!"
            End If
            txtAtkSuccess(x).Text = Monsterrec.AttackMinHCastPer(x)
        Else
            txtAtkMin(x).Text = Monsterrec.AttackMinHCastPer(x)
            txtAtkMax(x).Text = Monsterrec.AttackMaxHCastLvl(x)
            txtAtkSuccess(x).Text = Monsterrec.AttackAccuSpell(x)
            If Monsterrec.AttackHitSpell(x) > 0 Then
                
                nStatus = GetSpell(Monsterrec.AttackHitSpell(x))
                If nStatus = 0 Then
                    cmbAtkResist(x).ListIndex = Spellrec.TypeOfResists
                    txtAtkDur(x).Text = GetSpellDuration(Monsterrec.AttackHitSpell(x))
                    
                    If SpellHasAbility(Monsterrec.AttackHitSpell(x), 1) >= 0 Then
                        chkAtkDmgResist(x).Value = 0
                        txtAtkHitSpellMin(x).Text = GetSpellMinDamage(Monsterrec.AttackHitSpell(x))
                        txtAtkHitSpellMax(x).Text = GetSpellMaxDamage(Monsterrec.AttackHitSpell(x))
                        
                    ElseIf SpellHasAbility(Monsterrec.AttackHitSpell(x), 17) >= 0 Then
                        chkAtkDmgResist(x).Value = 1
                        txtAtkHitSpellMin(x).Text = GetSpellMinDamage(Monsterrec.AttackHitSpell(x))
                        txtAtkHitSpellMax(x).Text = GetSpellMaxDamage(Monsterrec.AttackHitSpell(x))
                        
                    Else
                        txtAtkHitSpellMin(x).Text = 0
                        txtAtkHitSpellMax(x).Text = 0
                    End If
                End If
            End If
        End If
    End If
Next x

nPercent = 0
For x = 0 To 4
    If Monsterrec.SpellNumber(x) > 0 Then
        txtBetweenSpellNumber(x).Text = Monsterrec.SpellNumber(x)
        txtBetweenSpellCastPer(x).Text = Monsterrec.SpellCastPer(x) - nPercent
        txtBetweenSpellCastLvL(x).Text = Monsterrec.SpellCastLvl(x)
        nPercent = Monsterrec.SpellCastPer(x)
    End If
Next x

For x = 0 To 4
    If Len(txtAtkName(x).Text) > 0 Then
        For y = 0 To 4
            If y <> x And txtAtkName(x).Text = txtAtkName(y).Text Then
                txtAtkName(x).Text = txtAtkName(x).Text & "-" & (x + 1)
                txtAtkName(y).Text = txtAtkName(y).Text & "-" & (y + 1)
            End If
        Next y
    End If
Next x

Call CheckAttackPercents

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("PopulateAttacks")
Resume out:
End Sub

Private Sub CheckAttackPercents()
Dim x As Integer, nPercent As Integer
On Error GoTo error:

nPercent = 0
For x = 0 To 4
    nPercent = nPercent + Val(txtAtkChance(x).Text)
Next x
If Not nPercent = 100 And nPercent > 0 Then
    lblAtkChance.ForeColor = &HFF&
Else
    lblAtkChance.ForeColor = &H80000012
End If

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("CheckAttackPercents")
Resume out:
End Sub

Public Sub ResetMonsterFields()
Dim x As Integer
On Error GoTo error:

lblItemBonus.Caption = ""
txtWeaponNumber.Text = ""
For x = 0 To 9
    txtItemNumber(x).Text = ""
    txtItemName(x).Text = ""
    txtItemDropPer(x).Text = ""
Next x

For x = 0 To 4
    txtBetweenSpellCastLvL(x).Text = ""
    txtBetweenSpellCastPer(x).Text = ""
    txtBetweenSpellName(x).Text = ""
    txtBetweenSpellNumber(x).Text = ""

    txtStatTrueCast(x).Text = ""
    txtStatAttRound(x).Text = ""
    txtStatAvgRound(x).Text = ""
    txtStatSuccess(x).Text = ""
    txtStatDmgResist(x).Text = ""
    txtStatResistDodge(x).Text = ""
    
    txtAtkName(x).Text = ""
    cmbAtkType(x).ListIndex = 0
    cmbAtkResist(x).ListIndex = 0
    chkAtkDmgResist(x).Value = 0
    txtAtkEnergy(x).Text = ""
    txtAtkMin(x).Text = ""
    txtAtkMax(x).Text = ""
    txtAtkChance(x).Text = ""
    txtAtkSuccess(x).Text = ""
    
    txtAtkDur(x).Text = ""
    txtAtkHitSpellMin(x).Text = ""
    txtAtkHitSpellMax(x).Text = ""
Next x

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("ResetMonsterFields")
Resume out:
End Sub

Private Sub cmbRefresh_Click()
Call RefreshMonsters
End Sub

Private Sub cmdBetweenRoundSpellGoto_Click(Index As Integer)
On Error GoTo error:
If Val(txtBetweenSpellNumber(Index).Text) <= 0 Then Exit Sub
Call frmSpell.GotoSpell(Val(txtBetweenSpellNumber(Index).Text))
frmSpell.Show
frmSpell.SetFocus

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("cmdBetweenRoundSpellGoto_Click")
Resume out:
End Sub


Private Sub cmdEditItemDrop_Click(Index As Integer)
If Val(txtItemNumber(Index).Text) < 1 Then Exit Sub
Call frmItem.GotoItem(Val(txtItemNumber(Index).Text))
frmItem.Show
frmItem.SetFocus
End Sub

Private Sub cmdGotoMon_Click()
On Error GoTo error:

If cmbMonster.ListIndex < 0 Then Exit Sub
frmMonster.GotoMonster (cmbMonster.ItemData(cmbMonster.ListIndex))

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("cmdGotoMon_Click")
Resume out:
End Sub

Private Sub cmdGotoWeapon_Click()
If Val(txtWeaponNumber.Text) < 1 Then Exit Sub
Call frmItem.GotoItem(Val(txtWeaponNumber.Text))
frmItem.Show
frmItem.SetFocus
End Sub

Private Sub cmdItemNote_Click()
    MsgBox "Items that are *ON* the monster will give them bonuses to ACCY and +min/max damage... " _
        & "but only if they *will* drop it or are *actually* holding it.  " _
        & "When monsters spawn, the ""roll"" for the items they are going to drop is decided, not when they are killed.  " _
        & vbCrLf & vbCrLf & "What I'm doing is adding non-limited drops and monster weapons to the list for calculation " _
        & "(if they have +accy or +max damage).  " _
        & "The bonus is then multiplied against the drop chance (e.g. 10 +max damage * 75% drop). " _
        & "Also note that bonuses stack for the same items." _
        & vbCrLf & vbCrLf & "If you want to force-include an item, just add it to the list and set the drop chance to 100%." _
        , vbOKOnly Or vbInformation, "Note on Items"
End Sub

Private Sub cmdSim_Click()
Dim x As Integer, nPrevPercent As Integer, nTest As Currency, nStatus As Integer
Dim nItemAccyBonus As Currency, nItemDamageBonus As Currency, nPrevBetweenPercent As Integer
Dim nDamageArr As Variant, nAccyArr As Variant
On Error GoTo error:

lblResultsAvgDmg.Caption = ""
lblResultsMaxRound.Caption = ""
If Val(txtMonsterEnergy.Text) < 0 Or Val(txtMonsterEnergy.Text) > 9999 Then txtMonsterEnergy.Text = 1000
If Val(txtNumRounds.Text) < 0 Or Val(txtNumRounds.Text) > 500000 Then txtNumRounds.Text = 500000

Call clsMonAtkSim.ResetValues

If chkDynamicRounds.Value = 1 Then
    clsMonAtkSim.bDynamicCalc = True
Else
    clsMonAtkSim.bDynamicCalc = False
End If

clsMonAtkSim.bUseCPU = bUseCPU
clsMonAtkSim.nNumberOfRounds = Val(txtNumRounds.Text)
clsMonAtkSim.nEnergyPerRound = Val(txtMonsterEnergy.Text)
clsMonAtkSim.nCombatLogMaxRounds = 50

nDamageArr = Array(4) '4=max damage
nAccyArr = Array(22, 105, 106) '22, 105, 106 = accuracy

nItemDamageBonus = CalculateItemBonuses(nDamageArr)
nItemAccyBonus = CalculateItemBonuses(nAccyArr)

nPrevPercent = 0
For x = 0 To 4
    txtStatTrueCast(x).Text = ""
    txtStatAttRound(x).Text = ""
    txtStatAvgRound(x).Text = ""
    txtStatSuccess(x).Text = ""
    txtStatDmgResist(x).Text = ""
    txtStatResistDodge(x).Text = ""
    
    txtAtkName(x).Text = Trim(txtAtkName(x).Text)
    
    If Len(txtAtkName(x).Text) > 20 Then txtAtkName(x).Text = Mid(txtAtkName(x).Text, 1, 20)
    If cmbAtkType(x).ListIndex > 0 Then
        If Len(txtAtkName(x).Text) = 0 Then txtAtkName(x).Text = "Attack " & x
    End If
    If Val(txtAtkEnergy(x).Text) > 9000 Then txtAtkEnergy(x).Text = 9000
    If Val(txtAtkMin(x).Text) > 99999 Then txtAtkMin(x).Text = 99999
    If Val(txtAtkMax(x).Text) > 99999 Then txtAtkMax(x).Text = 99999
    If Val(txtAtkChance(x).Text) > 100 Then txtAtkChance(x).Text = 100
    If Val(txtAtkSuccess(x).Text) > 9999 Then txtAtkSuccess(x).Text = 9999
    
    clsMonAtkSim.sAtkName(x) = txtAtkName(x).Text
    clsMonAtkSim.nAtkType(x) = cmbAtkType(x).ListIndex
    clsMonAtkSim.nAtkEnergy(x) = Val(txtAtkEnergy(x).Text)
    clsMonAtkSim.nAtkMin(x) = Val(txtAtkMin(x).Text) + IIf(cmbAtkType(x).ListIndex = 1, nItemDamageBonus, 0)
    clsMonAtkSim.nAtkMax(x) = Val(txtAtkMax(x).Text) + IIf(cmbAtkType(x).ListIndex = 1, nItemDamageBonus, 0)
    clsMonAtkSim.nAtkChance(x) = Val(txtAtkChance(x).Text) + nPrevPercent
    clsMonAtkSim.nAtkSuccess(x) = Val(txtAtkSuccess(x).Text) + IIf(cmbAtkType(x).ListIndex = 1, nItemAccyBonus, 0)
    
    clsMonAtkSim.nAtkHitSpellMin(x) = Val(txtAtkHitSpellMin(x).Text)
    clsMonAtkSim.nAtkHitSpellMax(x) = Val(txtAtkHitSpellMax(x).Text)
    clsMonAtkSim.nAtkResist(x) = cmbAtkResist(x).ListIndex
    clsMonAtkSim.nAtkDuration(x) = Val(txtAtkDur(x).Text)
    clsMonAtkSim.nAtkMRdmgResist(x) = chkAtkDmgResist(x).Value
    
    If Val(txtBetweenSpellNumber(x).Text) > 0 And Val(txtBetweenSpellCastPer(x).Text) > 0 Then
        nStatus = GetSpell(Val(txtBetweenSpellNumber(x).Text))
        If nStatus = 0 Then
            
            clsMonAtkSim.sBetweenRoundName(x) = ClipNull(Spellrec.Name)
            clsMonAtkSim.nBetweenRoundResistType(x) = Spellrec.TypeOfResists
            clsMonAtkSim.nBetweenRoundChance(x) = Val(txtBetweenSpellCastPer(x).Text) + nPrevBetweenPercent
            clsMonAtkSim.nBetweenRoundDuration(x) = GetSpellDuration(Val(txtBetweenSpellNumber(x).Text), Val(txtBetweenSpellCastLvL(x).Text))
            
            nTest = SpellHasAbility(Val(txtBetweenSpellNumber(x).Text), 1) '1=damage
            If nTest >= 0 Then
                clsMonAtkSim.nBetweenRoundResistDmgMR(x) = 0 'NO MR resist
                If nTest > 0 Then
                    clsMonAtkSim.nBetweenRoundMin(x) = nTest
                    clsMonAtkSim.nBetweenRoundMax(x) = nTest
                Else
                    clsMonAtkSim.nBetweenRoundMin(x) = GetSpellMinDamage(Val(txtBetweenSpellNumber(x).Text), Val(txtBetweenSpellCastLvL(x).Text))
                    clsMonAtkSim.nBetweenRoundMax(x) = GetSpellMaxDamage(Val(txtBetweenSpellNumber(x).Text), Val(txtBetweenSpellCastLvL(x).Text))
                End If
            End If
            
            nTest = SpellHasAbility(Val(txtBetweenSpellNumber(x).Text), 17) '17=damage-mr
            If nTest >= 0 Then
                clsMonAtkSim.nBetweenRoundResistDmgMR(x) = 1 'MR resist
                If nTest > 0 Then
                    clsMonAtkSim.nBetweenRoundMin(x) = nTest
                    clsMonAtkSim.nBetweenRoundMax(x) = nTest
                Else
                    clsMonAtkSim.nBetweenRoundMin(x) = GetSpellMinDamage(Val(txtBetweenSpellNumber(x).Text), Val(txtBetweenSpellCastLvL(x).Text))
                    clsMonAtkSim.nBetweenRoundMax(x) = GetSpellMaxDamage(Val(txtBetweenSpellNumber(x).Text), Val(txtBetweenSpellCastLvL(x).Text))
                End If
            End If
            
            nTest = SpellHasAbility(Val(txtBetweenSpellNumber(x).Text), 8) '8=drain
            If nTest >= 0 Then
                clsMonAtkSim.nBetweenRoundResistDmgMR(x) = 0 'NO MR resist
                If nTest > 0 Then
                    clsMonAtkSim.nBetweenRoundMin(x) = nTest
                    clsMonAtkSim.nBetweenRoundMax(x) = nTest
                Else
                    clsMonAtkSim.nBetweenRoundMin(x) = GetSpellMinDamage(Val(txtBetweenSpellNumber(x).Text), Val(txtBetweenSpellCastLvL(x).Text))
                    clsMonAtkSim.nBetweenRoundMax(x) = GetSpellMaxDamage(Val(txtBetweenSpellNumber(x).Text), Val(txtBetweenSpellCastLvL(x).Text))
                End If
            End If
        End If
        nPrevBetweenPercent = clsMonAtkSim.nBetweenRoundChance(x)
    End If
    
    nPrevPercent = clsMonAtkSim.nAtkChance(x)
Next x

If Val(txtUserAC.Text) > 9999 Then txtUserAC.Text = 9999
If Val(txtUserDR.Text) > 9999 Then txtUserDR.Text = 9999
If Val(txtUserDodge.Text) > 100 Then txtUserDodge.Text = 100
If Val(txtUserMR.Text) > 9999 Then txtUserMR.Text = 9999

clsMonAtkSim.nUserAC = Val(txtUserAC.Text)
clsMonAtkSim.nUserDR = Val(txtUserDR.Text)
clsMonAtkSim.nUserDodge = Val(txtUserDodge.Text)
clsMonAtkSim.nUserMR = Val(txtUserMR.Text)
clsMonAtkSim.nUserAntiMagic = chkUserAntiMagic.Value

txtCombatLog.Text = ""

Call clsMonAtkSim.RunSim

txtCombatLog.Text = clsMonAtkSim.sCombatLog

If clsMonAtkSim.nTotalAttacks > 0 And clsMonAtkSim.nNumberOfRounds > 0 Then
    lblResultsAvgDmg.Caption = "AVG Dmg/Rnd: " & Round(clsMonAtkSim.nTotalDamage / clsMonAtkSim.nNumberOfRounds, 1)
    lblResultsMaxRound.Caption = "Max Seen: " & clsMonAtkSim.nMaxRoundDamage
    
    For x = 0 To 4
        If clsMonAtkSim.nAtkType(x) > 0 Then
            txtStatTrueCast(x).Text = Round(clsMonAtkSim.nStatAtkAttempted(x) / clsMonAtkSim.nTotalAttacks, 3) * 100
            txtStatAttRound(x).Text = Round(clsMonAtkSim.nStatAtkAttempted(x) / clsMonAtkSim.nNumberOfRounds, 2)
            
            If clsMonAtkSim.nNumberOfRounds > 0 Then
                txtStatAvgRound(x).Text = Round(clsMonAtkSim.nStatAtkTotalDamage(x) / clsMonAtkSim.nNumberOfRounds)
            Else
                txtStatAvgRound(x).Text = 0
            End If
            
            If clsMonAtkSim.nStatAtkAttempted(x) > 0 Then
                txtStatSuccess(x).Text = Round(clsMonAtkSim.nStatAtkHits(x) / clsMonAtkSim.nStatAtkAttempted(x), 3) * 100
            Else
                txtStatSuccess(x).Text = 0
            End If
            
            If clsMonAtkSim.nStatAtkDmgResisted(x) <> 0 Then
                txtStatDmgResist(x).Text = IIf(clsMonAtkSim.nStatAtkTotalDamage(x) = 0, 100, _
                    Round(clsMonAtkSim.nStatAtkDmgResisted(x) / (clsMonAtkSim.nStatAtkDmgResisted(x) + clsMonAtkSim.nStatAtkTotalDamage(x)), 3) * 100)
            Else
                txtStatDmgResist(x).Text = 0
            End If
            
            If clsMonAtkSim.nStatAtkAttempted(x) > 0 And clsMonAtkSim.nAtkType(x) = 2 Then 'spell
                txtStatResistDodge(x).Text = Round(clsMonAtkSim.nStatAtkAttemptDodgedOrResisted(x) / clsMonAtkSim.nStatAtkAttempted(x), 3) * 100
            ElseIf clsMonAtkSim.nStatAtkHits(x) > 0 Or clsMonAtkSim.nStatAtkAttemptDodgedOrResisted(x) > 0 Then
                txtStatResistDodge(x).Text = Round(clsMonAtkSim.nStatAtkAttemptDodgedOrResisted(x) / (clsMonAtkSim.nStatAtkHits(x) + clsMonAtkSim.nStatAtkAttemptDodgedOrResisted(x)), 3) * 100
            Else
                txtStatResistDodge(x).Text = 0
            End If
        End If
    Next x
End If

out:
On Error Resume Next
ProgressBar.Value = 0
Exit Sub
error:
Call HandleError("cmdSim_Click")
Resume out:
End Sub

Private Sub Form_Load()
Dim x As Integer
On Error GoTo error:

Call RefreshMonsters
Call ResetMonsterFields

For x = 0 To 4
    cmbAtkType(x).ListIndex = 0
    cmbAtkResist(x).ListIndex = 0
Next x

Set clsMonAtkSim.cProgressBar = ProgressBar

SSTab1.TabCaption(0) = "Between" & vbCrLf & "Round Spells"
SSTab1.TabCaption(1) = "Attack" & vbCrLf & "Statistics"

lblHitSpell.Caption = "Hit Spell" & vbCrLf & "Min - Max"

txtUserAC.Text = ReadINI("Options", "MonSim-UserAC")
txtUserDR.Text = ReadINI("Options", "MonSim-UserDR")
txtUserDodge.Text = ReadINI("Options", "MonSim-UserDodge")
txtUserMR.Text = ReadINI("Options", "MonSim-UserMR", , 50)
chkUserAntiMagic.Value = ReadINI("Options", "MonSim-UserAntiMagic", , 0)
chkDynamicRounds.Value = ReadINI("Options", "MonSim-DynamicRounds", , 1)

Call chkDynamicRounds_Click

Me.Left = ReadINI("Windows", "MonSim-Left")
Me.Top = ReadINI("Windows", "MonSim-Top")

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("Form_Load")
Resume out:
End Sub

Public Sub GotoMonster(ByVal nMonster As Long)
Dim x As Integer

For x = 0 To cmbMonster.ListCount - 1
    If cmbMonster.ItemData(x) = nMonster Then
        cmbMonster.ListIndex = x
        Exit For
    End If
Next x

End Sub

Public Sub RefreshMonsters()
Dim nStatus As Integer
On Error GoTo error:

nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "LoadMonsters, BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

cmbMonster.clear

Do While nStatus = 0
    MonsterRowToStruct Monsterdatabuf.buf
    cmbMonster.AddItem ClipNull(Monsterrec.Name) & " (" & Monsterrec.Number & ")"
    cmbMonster.ItemData(cmbMonster.NewIndex) = Monsterrec.Number
    nStatus = BTRCALL(BGETNEXT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
Loop

If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "RefreshMonsters, Error: " & BtrieveErrorCode(nStatus)
End If

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("RefreshMonsters")
Resume out:
End Sub

Private Sub fraStats_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Call WriteINI("Options", "MonSim-UserAC", txtUserAC.Text)
Call WriteINI("Options", "MonSim-UserDR", txtUserDR.Text)
Call WriteINI("Options", "MonSim-UserDodge", txtUserDodge.Text)
Call WriteINI("Options", "MonSim-UserMR", txtUserMR.Text)
Call WriteINI("Options", "MonSim-UserAntiMagic", chkUserAntiMagic.Value)
Call WriteINI("Options", "MonSim-DynamicRounds", chkDynamicRounds.Value)

Call WriteINI("Windows", "MonSim-Left", Me.Left)
Call WriteINI("Windows", "MonSim-Top", Me.Top)
End Sub

Private Sub lblAntiMagic_Click()
If chkUserAntiMagic.Value = 0 Then
    chkUserAntiMagic.Value = 1
Else
    chkUserAntiMagic.Value = 0
End If
End Sub


Private Sub txtAtkChance_GotFocus(Index As Integer)
Call SelectAll(txtAtkChance(Index))
End Sub

Private Sub txtAtkChance_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtAtkChance_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Call CheckAttackPercents
End Sub

Private Sub txtAtkDur_GotFocus(Index As Integer)
Call SelectAll(txtAtkDur(Index))
End Sub

Private Sub txtAtkEnergy_GotFocus(Index As Integer)
Call SelectAll(txtAtkEnergy(Index))
End Sub

Private Sub txtAtkEnergy_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtAtkHitSpellMax_GotFocus(Index As Integer)
Call SelectAll(txtAtkHitSpellMax(Index))
End Sub

Private Sub txtAtkHitSpellMax_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtAtkHitSpellMin_GotFocus(Index As Integer)
Call SelectAll(txtAtkHitSpellMin(Index))
End Sub

Private Sub txtAtkHitSpellMin_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtAtkMax_GotFocus(Index As Integer)
Call SelectAll(txtAtkMax(Index))
End Sub

Private Sub txtAtkMax_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtAtkMin_GotFocus(Index As Integer)
Call SelectAll(txtAtkMin(Index))
End Sub

Private Sub txtAtkMin_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtAtkName_GotFocus(Index As Integer)
Call SelectAll(txtAtkName(Index))
End Sub

Private Sub txtAtkSuccess_GotFocus(Index As Integer)
Call SelectAll(txtAtkSuccess(Index))
End Sub

Private Sub txtAtkSuccess_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtBetweenSpellCastLvL_GotFocus(Index As Integer)
Call SelectAll(txtBetweenSpellCastLvL(Index))
End Sub

Private Sub txtBetweenSpellCastLvL_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtBetweenSpellCastPer_GotFocus(Index As Integer)
Call SelectAll(txtBetweenSpellCastPer(Index))
End Sub

Private Sub txtBetweenSpellCastPer_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtBetweenSpellNumber_GotFocus(Index As Integer)
Call SelectAll(txtBetweenSpellNumber(Index))
End Sub

Private Sub txtBetweenSpellNumber_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtItemDropPer_Change(Index As Integer)
Call UpdateItemBonusDisplay
End Sub

Private Sub txtItemDropPer_GotFocus(Index As Integer)
Call SelectAll(txtItemDropPer(Index))
End Sub

Private Sub txtItemDropPer_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtItemNumber_Change(Index As Integer)
On Error GoTo error:
If Val(txtItemNumber(Index).Text) > 0 Then
    txtItemName(Index).Text = GetItemName(Val(txtItemNumber(Index).Text))
Else
    txtItemName(Index).Text = ""
End If
Call UpdateItemBonusDisplay
out:
On Error Resume Next
Exit Sub
error:
Call HandleError("txtItemNumber_Change")
Resume out:
End Sub

Private Sub txtItemNumber_GotFocus(Index As Integer)
Call SelectAll(txtItemNumber(Index))
End Sub

Private Sub txtItemNumber_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtMonsterEnergy_GotFocus()
Call SelectAll(txtMonsterEnergy(Index))
End Sub

Private Sub txtMonsterEnergy_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub


Private Sub txtNumRounds_GotFocus()
Call SelectAll(txtNumRounds)
End Sub

Private Sub txtNumRounds_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub UpdateItemBonusDisplay()
Dim nItemDamageBonus As Long, nItemAccyBonus As Long
Dim nDamageArr As Variant, nAccyArr As Variant
Dim sText As String
On Error GoTo error:

nDamageArr = Array(4)
nAccyArr = Array(22, 105, 106)

nItemDamageBonus = CalculateItemBonuses(nDamageArr)
nItemAccyBonus = CalculateItemBonuses(nAccyArr)

If nItemDamageBonus <> 0 Then
    If Not sText = "" Then sText = sText & ", "
    sText = sText & "Physical Min/Max Dmg: " & IIf(nItemDamageBonus > 0, "+", "") & nItemDamageBonus
End If

If nItemAccyBonus <> 0 Then
    If Not sText = "" Then sText = sText & ", "
    sText = sText & "Accuracy: " & IIf(nItemAccyBonus > 0, "+", "") & nItemAccyBonus
End If

If Not sText = "" Then sText = "Bonuses-- " & sText

lblItemBonus.Caption = sText

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("UpdateItemBonusDisplay")
Resume out:
End Sub

Private Function CalculateItemBonuses(nAbilities As Variant) As Integer
Dim x As Integer, y As Integer, nTest As Integer
On Error GoTo error:

If Not IsDimmed(nAbilities) Then Exit Function

If Val(txtWeaponNumber.Text) > 0 Then
    For y = LBound(nAbilities) To UBound(nAbilities)
        nTest = ItemHasAbility(Val(txtWeaponNumber.Text), nAbilities(y))
        If nTest > 0 Then
            CalculateItemBonuses = CalculateItemBonuses + nTest
        End If
    Next y
End If

For x = 0 To 9
    If Val(txtItemNumber(x).Text) > 0 Then
        For y = LBound(nAbilities) To UBound(nAbilities)
            nTest = ItemHasAbility(Val(txtItemNumber(x).Text), nAbilities(y))
            If nTest > 0 Then
                If Val(txtItemDropPer(x).Text) > 100 Then txtItemDropPer(x).Text = 100
                CalculateItemBonuses = CalculateItemBonuses + (nTest * (Val(txtItemDropPer(x).Text) / 100))
            End If
        Next y
    End If
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalculateItemBonuses")
Resume out:
End Function

Private Sub txtBetweenSpellNumber_Change(Index As Integer)
On Error GoTo error:

If Val(txtBetweenSpellNumber(Index).Text) > 0 Then
    txtBetweenSpellName(Index).Text = GetSpellName(Val(txtBetweenSpellNumber(Index).Text))
Else
    txtBetweenSpellName(Index).Text = ""
End If

out:
Exit Sub
error:
Call HandleError("txtSpellNumber_Change")
Resume out:
End Sub

Private Sub txtUserAC_GotFocus()
Call SelectAll(txtUserAC)
End Sub

Private Sub txtUserAC_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtUserDodge_GotFocus()
Call SelectAll(txtUserDodge)
End Sub

Private Sub txtUserDodge_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtUserDR_GotFocus()
Call SelectAll(txtUserDR)
End Sub

Private Sub txtUserDR_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtUserMR_GotFocus()
Call SelectAll(txtUserMR)
End Sub

Private Sub txtUserMR_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtWeaponNumber_Change()
On Error GoTo error:

If Val(txtWeaponNumber.Text) > 0 Then
    txtWeaponName.Text = GetItemName(Val(txtWeaponNumber.Text))
Else
    txtWeaponName.Text = ""
End If
Call UpdateItemBonusDisplay
out:
On Error Resume Next
Exit Sub
error:
Call HandleError("txtWeaponNumber_Change")
Resume out:
End Sub

Private Sub txtWeaponNumber_GotFocus()
Call SelectAll(txtWeaponNumber)
End Sub

Private Sub txtWeaponNumber_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub
