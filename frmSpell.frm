VERSION 5.00
Object = "{20D5284F-7B23-4F0A-B8B1-6C9D18B64F1C}#1.0#0"; "exlimiter.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSpell 
   Caption         =   "Spell Editor"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11205
   Icon            =   "frmSpell.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   11205
   Begin VB.Frame fraFilter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   4755
      Left            =   60
      TabIndex        =   135
      Top             =   360
      Visible         =   0   'False
      Width           =   6675
      Begin VB.Frame fraFilter2 
         Caption         =   "Filtering Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   180
         TabIndex        =   136
         Top             =   180
         Width           =   6255
         Begin VB.CheckBox chkFilterExcludeZero 
            Caption         =   "Exclude 0 value on <= search"
            Height          =   255
            Left            =   3300
            TabIndex        =   176
            Top             =   360
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.TextBox txtFilterTB 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1380
            MaxLength       =   29
            TabIndex        =   175
            Text            =   "0"
            Top             =   3780
            Width           =   1575
         End
         Begin VB.TextBox txtFilterMessage 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1380
            MaxLength       =   29
            TabIndex        =   174
            Text            =   "0"
            Top             =   3420
            Width           =   1575
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Textblock"
            Enabled         =   0   'False
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   173
            Top             =   3780
            Width           =   1095
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Message"
            Enabled         =   0   'False
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   172
            Top             =   3420
            Width           =   1095
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Level Cap"
            Enabled         =   0   'False
            Height          =   255
            Index           =   9
            Left            =   3300
            TabIndex        =   171
            Top             =   1500
            Width           =   1095
         End
         Begin VB.TextBox txtFilterAbilityValue 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   9
            Left            =   5340
            TabIndex        =   170
            Text            =   "0"
            Top             =   1500
            Width           =   675
         End
         Begin VB.ComboBox cmbFilterAbilityGL 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   9
            ItemData        =   "frmSpell.frx":08CA
            Left            =   4440
            List            =   "frmSpell.frx":08D7
            Style           =   2  'Dropdown List
            TabIndex        =   169
            Top             =   1500
            Width           =   795
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Req Level"
            Enabled         =   0   'False
            Height          =   255
            Index           =   8
            Left            =   3300
            TabIndex        =   168
            Top             =   1140
            Width           =   1095
         End
         Begin VB.TextBox txtFilterAbilityValue 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   8
            Left            =   5340
            TabIndex        =   167
            Text            =   "0"
            Top             =   1140
            Width           =   675
         End
         Begin VB.ComboBox cmbFilterAbilityGL 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   8
            ItemData        =   "frmSpell.frx":08E6
            Left            =   4440
            List            =   "frmSpell.frx":08F3
            Style           =   2  'Dropdown List
            TabIndex        =   166
            Top             =   1140
            Width           =   795
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Energy"
            Enabled         =   0   'False
            Height          =   255
            Index           =   7
            Left            =   3300
            TabIndex        =   165
            Top             =   780
            Width           =   1095
         End
         Begin VB.TextBox txtFilterAbilityValue 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   7
            Left            =   5340
            TabIndex        =   164
            Text            =   "0"
            Top             =   780
            Width           =   675
         End
         Begin VB.ComboBox cmbFilterAbilityGL 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   7
            ItemData        =   "frmSpell.frx":0902
            Left            =   4440
            List            =   "frmSpell.frx":090F
            Style           =   2  'Dropdown List
            TabIndex        =   163
            Top             =   780
            Width           =   795
         End
         Begin VB.TextBox txtFilterAbilityValue 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   3
            Left            =   3780
            TabIndex        =   162
            Text            =   "0"
            Top             =   1980
            Width           =   555
         End
         Begin VB.ComboBox cmbFilterAbilityGL 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   3
            ItemData        =   "frmSpell.frx":091E
            Left            =   3000
            List            =   "frmSpell.frx":092E
            Style           =   2  'Dropdown List
            TabIndex        =   161
            Top             =   1980
            Width           =   735
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Target"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   160
            Top             =   1140
            Width           =   1095
         End
         Begin VB.ComboBox cmbFilter 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   1
            ItemData        =   "frmSpell.frx":0942
            Left            =   1380
            List            =   "frmSpell.frx":0970
            Style           =   2  'Dropdown List
            TabIndex        =   159
            Top             =   1140
            Width           =   1575
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Magery"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   158
            Top             =   1980
            Width           =   1095
         End
         Begin VB.ComboBox cmbFilter 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   3
            ItemData        =   "frmSpell.frx":0A3E
            Left            =   1380
            List            =   "frmSpell.frx":0A54
            Style           =   2  'Dropdown List
            TabIndex        =   157
            Top             =   1980
            Width           =   1575
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Ability"
            Enabled         =   0   'False
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   156
            Top             =   2340
            Width           =   1095
         End
         Begin VB.ComboBox cmbFilter 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   4
            ItemData        =   "frmSpell.frx":0A7E
            Left            =   1380
            List            =   "frmSpell.frx":0A80
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   155
            Top             =   2340
            Width           =   1575
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Type"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   154
            Top             =   780
            Width           =   1095
         End
         Begin VB.ComboBox cmbFilter 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   0
            ItemData        =   "frmSpell.frx":0A82
            Left            =   1380
            List            =   "frmSpell.frx":0A92
            Style           =   2  'Dropdown List
            TabIndex        =   153
            Top             =   780
            Width           =   1575
         End
         Begin VB.CheckBox chkFilterNone 
            Caption         =   "No Filter"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   152
            Top             =   360
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CommandButton cmdFilterCancel 
            Caption         =   "Cancel"
            Height          =   495
            Left            =   4680
            TabIndex        =   151
            Top             =   3600
            Width           =   1335
         End
         Begin VB.CommandButton cmdFilterApply 
            Caption         =   "Apply Filter"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4680
            TabIndex        =   150
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton cmdFilterReset 
            Caption         =   "Reset"
            Height          =   495
            Left            =   4680
            TabIndex        =   149
            Top             =   3000
            Width           =   1335
         End
         Begin VB.TextBox txtFilterAbilityValue 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   4
            Left            =   3780
            TabIndex        =   148
            Text            =   "0"
            Top             =   2340
            Width           =   555
         End
         Begin VB.ComboBox cmbFilterAbilityGL 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   4
            ItemData        =   "frmSpell.frx":0ABF
            Left            =   3000
            List            =   "frmSpell.frx":0ACF
            Style           =   2  'Dropdown List
            TabIndex        =   147
            Top             =   2340
            Width           =   735
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Element"
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   146
            Top             =   1500
            Width           =   1095
         End
         Begin VB.ComboBox cmbFilter 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   2
            ItemData        =   "frmSpell.frx":0AE3
            Left            =   1380
            List            =   "frmSpell.frx":0AFC
            Style           =   2  'Dropdown List
            TabIndex        =   145
            Top             =   1500
            Width           =   1575
         End
         Begin VB.ComboBox cmbFilterAbilityGL 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   5
            ItemData        =   "frmSpell.frx":0B34
            Left            =   3000
            List            =   "frmSpell.frx":0B44
            Style           =   2  'Dropdown List
            TabIndex        =   144
            Top             =   2700
            Width           =   735
         End
         Begin VB.TextBox txtFilterAbilityValue 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   5
            Left            =   3780
            TabIndex        =   143
            Text            =   "0"
            Top             =   2700
            Width           =   555
         End
         Begin VB.ComboBox cmbFilter 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   5
            ItemData        =   "frmSpell.frx":0B58
            Left            =   1380
            List            =   "frmSpell.frx":0B5A
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   142
            Top             =   2700
            Width           =   1575
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Ability"
            Enabled         =   0   'False
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   141
            Top             =   2700
            Width           =   1095
         End
         Begin VB.ComboBox cmbFilterAbilityGL 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   6
            ItemData        =   "frmSpell.frx":0B5C
            Left            =   3000
            List            =   "frmSpell.frx":0B6C
            Style           =   2  'Dropdown List
            TabIndex        =   140
            Top             =   3060
            Width           =   735
         End
         Begin VB.TextBox txtFilterAbilityValue 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   6
            Left            =   3780
            TabIndex        =   139
            Text            =   "0"
            Top             =   3060
            Width           =   555
         End
         Begin VB.ComboBox cmbFilter 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   6
            ItemData        =   "frmSpell.frx":0B80
            Left            =   1380
            List            =   "frmSpell.frx":0B82
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   138
            Top             =   3060
            Width           =   1575
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Ability"
            Enabled         =   0   'False
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   137
            Top             =   3060
            Width           =   1095
         End
      End
   End
   Begin VB.CommandButton cmdFilter 
      Caption         =   "Filter"
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
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   2895
   End
   Begin VB.Frame framNav 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   3060
      TabIndex        =   6
      Top             =   0
      Width           =   8055
      Begin VB.CheckBox chkAutoSave 
         Caption         =   "Auto-Save"
         Height          =   195
         Left            =   4620
         TabIndex        =   10
         Top             =   60
         Value           =   1  'Checked
         Width           =   1155
      End
      Begin VB.Frame frmGeneral 
         Caption         =   "General"
         Height          =   4155
         Left            =   0
         TabIndex        =   24
         Top             =   1740
         Width           =   8055
         Begin VB.CommandButton cmdFormulaCopy 
            Caption         =   "?"
            Height          =   255
            Index           =   2
            Left            =   7680
            TabIndex        =   79
            Top             =   2520
            Width           =   255
         End
         Begin VB.CommandButton cmdFormulaCopy 
            Caption         =   "Paste"
            Height          =   255
            Index           =   1
            Left            =   6660
            TabIndex        =   78
            Top             =   2520
            Width           =   975
         End
         Begin VB.CommandButton cmdFormulaCopy 
            Caption         =   "Copy"
            Height          =   255
            Index           =   0
            Left            =   5640
            TabIndex        =   77
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox txtVSMR 
            Height          =   285
            Left            =   3300
            MaxLength       =   4
            TabIndex        =   45
            Top             =   3060
            Width           =   615
         End
         Begin VB.TextBox txtMRvsCap 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   83
            Top             =   3720
            Width           =   2295
         End
         Begin VB.TextBox txtMRvsReq 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   82
            Top             =   3420
            Width           =   2295
         End
         Begin VB.TextBox txtEnergy 
            Height          =   285
            Left            =   5040
            TabIndex        =   54
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtLevel 
            Height          =   285
            Left            =   5040
            TabIndex        =   56
            Top             =   600
            Width           =   615
         End
         Begin VB.ComboBox cmbSpellType 
            Height          =   315
            ItemData        =   "frmSpell.frx":0B84
            Left            =   1320
            List            =   "frmSpell.frx":0B94
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   240
            Width           =   1995
         End
         Begin VB.TextBox txtMin 
            Height          =   285
            Left            =   5040
            TabIndex        =   60
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox txtMax 
            Height          =   285
            Left            =   5040
            TabIndex        =   62
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox txtMana 
            Height          =   285
            Left            =   1320
            TabIndex        =   34
            Top             =   2700
            Width           =   795
         End
         Begin VB.TextBox txtDifficulty 
            Height          =   285
            Left            =   1320
            TabIndex        =   33
            Top             =   2400
            Width           =   795
         End
         Begin VB.TextBox txtDuration 
            Height          =   285
            Left            =   5040
            TabIndex        =   64
            Top             =   2040
            Width           =   615
         End
         Begin VB.ComboBox cmbTarget 
            Height          =   315
            ItemData        =   "frmSpell.frx":0BC1
            Left            =   1320
            List            =   "frmSpell.frx":0BEF
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   600
            Width           =   1995
         End
         Begin VB.TextBox txtCastMsgA 
            Height          =   285
            Left            =   1320
            TabIndex        =   39
            Top             =   3420
            Width           =   795
         End
         Begin VB.TextBox txtCastMsgB 
            Height          =   285
            Left            =   1320
            TabIndex        =   43
            Top             =   3720
            Width           =   795
         End
         Begin VB.ComboBox cmbResistAbility 
            Height          =   315
            ItemData        =   "frmSpell.frx":0CBD
            Left            =   1320
            List            =   "frmSpell.frx":0CD6
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1680
            Width           =   1695
         End
         Begin VB.ComboBox cmbMageryA 
            Height          =   315
            ItemData        =   "frmSpell.frx":0D21
            Left            =   1320
            List            =   "frmSpell.frx":0D37
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   960
            Width           =   1035
         End
         Begin VB.TextBox txtMageryB 
            Height          =   315
            Left            =   2400
            MaxLength       =   5
            TabIndex        =   29
            Top             =   960
            Width           =   615
         End
         Begin VB.ComboBox cmbTypeOfAttack 
            Height          =   315
            ItemData        =   "frmSpell.frx":0D61
            Left            =   1320
            List            =   "frmSpell.frx":0D7A
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1320
            Width           =   1695
         End
         Begin VB.ComboBox cmbTypeOfResists 
            Height          =   315
            ItemData        =   "frmSpell.frx":0DB2
            Left            =   1320
            List            =   "frmSpell.frx":0DBF
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox txtLVLSMinIncr 
            Height          =   285
            Left            =   7320
            TabIndex        =   66
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtMinIncrease 
            Height          =   285
            Left            =   7320
            TabIndex        =   68
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txtLVLSMaxIncr 
            Height          =   285
            Left            =   7320
            TabIndex        =   70
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtMaxIncrease 
            Height          =   285
            Left            =   7320
            TabIndex        =   72
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox txtDurIncrease 
            Height          =   285
            Left            =   7320
            TabIndex        =   76
            Top             =   2040
            Width           =   615
         End
         Begin VB.TextBox txtLVLSDurIncr 
            Height          =   285
            Left            =   7320
            TabIndex        =   74
            Top             =   1680
            Width           =   615
         End
         Begin VB.ComboBox cmbMsgStyle 
            Height          =   315
            ItemData        =   "frmSpell.frx":0DE5
            Left            =   1320
            List            =   "frmSpell.frx":0DEF
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   3060
            Width           =   795
         End
         Begin VB.TextBox txtLevelCap 
            Height          =   285
            Left            =   5040
            TabIndex        =   58
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtCastADisplay 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   3420
            Width           =   1755
         End
         Begin VB.TextBox txtCastBDisplay 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   3720
            Width           =   1755
         End
         Begin VB.TextBox txtAtCap 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   81
            Top             =   3120
            Width           =   2295
         End
         Begin VB.TextBox txtAtReq 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   80
            Top             =   2820
            Width           =   2295
         End
         Begin VB.CommandButton cmdEditMsgA 
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   3420
            Width           =   195
         End
         Begin VB.CommandButton cmdEditMsgB 
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   3720
            Width           =   195
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "vs. MR"
            Height          =   255
            Index           =   9
            Left            =   2460
            TabIndex        =   134
            Top             =   3060
            Width           =   735
         End
         Begin VB.Label lblVSMR 
            Alignment       =   1  'Right Justify
            Caption         =   "@Cap vs 100MR"
            Height          =   195
            Index           =   1
            Left            =   3900
            TabIndex        =   133
            Top             =   3720
            Width           =   1635
         End
         Begin VB.Label lblVSMR 
            Alignment       =   1  'Right Justify
            Caption         =   "@Req vs 100MR"
            Height          =   195
            Index           =   0
            Left            =   3900
            TabIndex        =   132
            Top             =   3420
            Width           =   1635
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Energy"
            Height          =   255
            Index           =   24
            Left            =   4260
            TabIndex        =   53
            Top             =   240
            Width           =   675
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Req. Level"
            Height          =   255
            Index           =   22
            Left            =   4020
            TabIndex        =   55
            Top             =   600
            Width           =   915
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Spell Type"
            Height          =   315
            Index           =   132
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Base Min"
            Height          =   255
            Index           =   33
            Left            =   4020
            TabIndex        =   59
            Top             =   1320
            Width           =   915
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Mana"
            Height          =   255
            Index           =   26
            Left            =   660
            TabIndex        =   52
            Top             =   2700
            Width           =   555
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Difficulty"
            Height          =   255
            Index           =   25
            Left            =   540
            TabIndex        =   51
            Top             =   2400
            Width           =   675
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Base Duration"
            Height          =   255
            Index           =   23
            Left            =   3840
            TabIndex        =   63
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Target"
            Height          =   255
            Index           =   28
            Left            =   120
            TabIndex        =   46
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Cast Msg A"
            Height          =   315
            Index           =   3
            Left            =   360
            TabIndex        =   38
            Top             =   3420
            Width           =   855
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Cast Msg B"
            Height          =   315
            Index           =   130
            Left            =   240
            TabIndex        =   42
            Top             =   3720
            Width           =   975
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Resist Ability"
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   49
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Magery"
            Height          =   315
            Index           =   29
            Left            =   120
            TabIndex        =   47
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Type of Attack"
            Height          =   315
            Index           =   30
            Left            =   120
            TabIndex        =   48
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Type of Resists"
            Height          =   315
            Index           =   5
            Left            =   120
            TabIndex        =   50
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Max increase"
            Height          =   255
            Index           =   36
            Left            =   6060
            TabIndex        =   71
            Top             =   1320
            Width           =   1155
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "LVLs/Max increase"
            Height          =   255
            Index           =   35
            Left            =   5700
            TabIndex        =   69
            Top             =   960
            Width           =   1515
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Norm Msg Style"
            Height          =   315
            Index           =   10
            Left            =   60
            TabIndex        =   35
            Top             =   3060
            Width           =   1155
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Level Cap"
            Height          =   195
            Index           =   27
            Left            =   4080
            TabIndex        =   57
            Top             =   960
            Width           =   855
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "LVLs/Min increase"
            Height          =   255
            Index           =   1
            Left            =   5700
            TabIndex        =   65
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Min increase"
            Height          =   255
            Index           =   2
            Left            =   6120
            TabIndex        =   67
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "LVLs/Dur. increase"
            Height          =   255
            Index           =   7
            Left            =   5700
            TabIndex        =   73
            Top             =   1680
            Width           =   1515
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Dur. increase"
            Height          =   255
            Index           =   8
            Left            =   6000
            TabIndex        =   75
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Base Max"
            Height          =   255
            Index           =   0
            Left            =   4020
            TabIndex        =   61
            Top             =   1680
            Width           =   915
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "@Req LVL"
            Height          =   195
            Index           =   0
            Left            =   4440
            TabIndex        =   84
            Top             =   2820
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "@LVL Cap"
            Height          =   195
            Index           =   1
            Left            =   4440
            TabIndex        =   85
            Top             =   3120
            Width           =   1095
         End
      End
      Begin VB.TextBox txtUNDEFINED02 
         Height          =   315
         Left            =   6120
         TabIndex        =   130
         Top             =   5460
         Width           =   615
      End
      Begin VB.TextBox txtUNDEFINED01 
         Height          =   315
         Left            =   5460
         TabIndex        =   129
         Top             =   5460
         Width           =   615
      End
      Begin VB.CommandButton cmdAbilsClear 
         Caption         =   "Clear"
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
         Left            =   5280
         TabIndex        =   128
         Top             =   2760
         Width           =   675
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   285
         Left            =   1020
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDiscard 
         Caption         =   "Dis&card"
         Height          =   285
         Left            =   7020
         TabIndex        =   12
         Top             =   0
         Width           =   1035
      End
      Begin VB.CommandButton cmdOther 
         Caption         =   "&Abilities"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2880
         TabIndex        =   9
         Top             =   0
         Width           =   1395
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "&Insert"
         Height          =   285
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   1035
      End
      Begin VB.Frame Frame1 
         Height          =   1395
         Left            =   0
         TabIndex        =   13
         Top             =   300
         Width           =   8055
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   120
            MaxLength       =   29
            TabIndex        =   16
            Top             =   360
            Width           =   2955
         End
         Begin VB.TextBox txtShortName 
            Height          =   315
            Left            =   3120
            TabIndex        =   17
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtDesc 
            Height          =   315
            Index           =   0
            Left            =   120
            MaxLength       =   50
            TabIndex        =   22
            Top             =   960
            Width           =   3995
         End
         Begin VB.TextBox txtDesc 
            Height          =   315
            Index           =   1
            Left            =   4140
            MaxLength       =   50
            TabIndex        =   23
            Top             =   960
            Width           =   3795
         End
         Begin VB.TextBox txtNumber 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   7080
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   240
            Width           =   795
         End
         Begin VB.Label label 
            Alignment       =   2  'Center
            Caption         =   "Short Name"
            Height          =   195
            Index           =   11
            Left            =   3120
            TabIndex        =   15
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Decription Line 1"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Description Line 2"
            Height          =   255
            Left            =   4140
            TabIndex        =   21
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Number"
            Height          =   255
            Left            =   6360
            TabIndex        =   18
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Name"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   180
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   285
         Left            =   5940
         TabIndex        =   11
         Top             =   0
         Width           =   1095
      End
      Begin VB.Frame frmAbilities 
         Caption         =   "Abilities"
         Height          =   4095
         Left            =   0
         TabIndex        =   86
         Top             =   1800
         Visible         =   0   'False
         Width           =   3555
         Begin VB.TextBox txtAbilityA 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   123
            Top             =   3600
            Width           =   495
         End
         Begin VB.TextBox txtAbilityA 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   119
            Top             =   3240
            Width           =   495
         End
         Begin VB.TextBox txtAbilityA 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   115
            Top             =   2880
            Width           =   495
         End
         Begin VB.TextBox txtAbilityA 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   111
            Top             =   2520
            Width           =   495
         End
         Begin VB.TextBox txtAbilityA 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   107
            Top             =   2160
            Width           =   495
         End
         Begin VB.TextBox txtAbilityA 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   103
            Top             =   1800
            Width           =   495
         End
         Begin VB.TextBox txtAbilityA 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   99
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox txtAbilityA 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   95
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtAbilityA 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   91
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtAbilityB 
            Height          =   315
            Index           =   5
            Left            =   2640
            TabIndex        =   109
            ToolTipText     =   "Enter the value for the ability here."
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtAbilityB 
            Height          =   315
            Index           =   6
            Left            =   2640
            TabIndex        =   113
            ToolTipText     =   "Enter the value for the ability here."
            Top             =   2520
            Width           =   615
         End
         Begin VB.TextBox txtAbilityB 
            Height          =   315
            Index           =   7
            Left            =   2640
            TabIndex        =   117
            ToolTipText     =   "Enter the value for the ability here."
            Top             =   2880
            Width           =   615
         End
         Begin VB.TextBox txtAbilityB 
            Height          =   315
            Index           =   8
            Left            =   2640
            TabIndex        =   121
            ToolTipText     =   "Enter the value for the ability here."
            Top             =   3240
            Width           =   615
         End
         Begin VB.TextBox txtAbilityB 
            Height          =   315
            Index           =   9
            Left            =   2640
            TabIndex        =   125
            ToolTipText     =   "Enter the value for the ability here."
            Top             =   3600
            Width           =   615
         End
         Begin VB.TextBox txtAbilityB 
            Height          =   285
            Index           =   0
            Left            =   2640
            TabIndex        =   89
            ToolTipText     =   "Enter the value for the ability here."
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtAbilityB 
            Height          =   315
            Index           =   1
            Left            =   2640
            TabIndex        =   93
            ToolTipText     =   "Enter the value for the ability here."
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtAbilityB 
            Height          =   315
            Index           =   2
            Left            =   2640
            TabIndex        =   97
            ToolTipText     =   "Enter the value for the ability here."
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtAbilityB 
            Height          =   315
            Index           =   3
            Left            =   2640
            TabIndex        =   101
            ToolTipText     =   "Enter the value for the ability here."
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtAbilityB 
            Height          =   315
            Index           =   4
            Left            =   2640
            TabIndex        =   105
            ToolTipText     =   "Enter the value for the ability here."
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox txtAbilityA 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   87
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox lblName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
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
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   88
            Text            =   "empty"
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox lblName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
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
            Height          =   285
            Index           =   1
            Left            =   720
            TabIndex        =   92
            Text            =   "empty"
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox lblName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
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
            Height          =   285
            Index           =   2
            Left            =   720
            TabIndex        =   96
            Text            =   "empty"
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox lblName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
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
            Height          =   285
            Index           =   3
            Left            =   720
            TabIndex        =   100
            Text            =   "empty"
            Top             =   1440
            Width           =   1815
         End
         Begin VB.TextBox lblName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
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
            Height          =   285
            Index           =   4
            Left            =   720
            TabIndex        =   104
            Text            =   "empty"
            Top             =   1800
            Width           =   1815
         End
         Begin VB.TextBox lblName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
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
            Height          =   285
            Index           =   5
            Left            =   720
            TabIndex        =   108
            Text            =   "empty"
            Top             =   2160
            Width           =   1815
         End
         Begin VB.TextBox lblName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
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
            Height          =   285
            Index           =   6
            Left            =   720
            TabIndex        =   112
            Text            =   "empty"
            Top             =   2520
            Width           =   1815
         End
         Begin VB.TextBox lblName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
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
            Height          =   285
            Index           =   7
            Left            =   720
            TabIndex        =   116
            Text            =   "empty"
            Top             =   2880
            Width           =   1815
         End
         Begin VB.TextBox lblName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
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
            Height          =   285
            Index           =   8
            Left            =   720
            TabIndex        =   120
            Text            =   "empty"
            Top             =   3240
            Width           =   1815
         End
         Begin VB.TextBox lblName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
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
            Height          =   285
            Index           =   9
            Left            =   720
            TabIndex        =   124
            Text            =   "empty"
            Top             =   3600
            Width           =   1815
         End
         Begin VB.CommandButton cmdAbilityLookup 
            Height          =   255
            Index           =   9
            Left            =   3300
            TabIndex        =   126
            Top             =   3600
            Width           =   135
         End
         Begin VB.CommandButton cmdAbilityLookup 
            Height          =   255
            Index           =   8
            Left            =   3300
            TabIndex        =   122
            Top             =   3240
            Width           =   135
         End
         Begin VB.CommandButton cmdAbilityLookup 
            Height          =   255
            Index           =   7
            Left            =   3300
            TabIndex        =   118
            Top             =   2880
            Width           =   135
         End
         Begin VB.CommandButton cmdAbilityLookup 
            Height          =   255
            Index           =   6
            Left            =   3300
            TabIndex        =   114
            Top             =   2520
            Width           =   135
         End
         Begin VB.CommandButton cmdAbilityLookup 
            Height          =   255
            Index           =   5
            Left            =   3300
            TabIndex        =   110
            Top             =   2160
            Width           =   135
         End
         Begin VB.CommandButton cmdAbilityLookup 
            Height          =   255
            Index           =   4
            Left            =   3300
            TabIndex        =   106
            Top             =   1800
            Width           =   135
         End
         Begin VB.CommandButton cmdAbilityLookup 
            Height          =   255
            Index           =   3
            Left            =   3300
            TabIndex        =   102
            Top             =   1440
            Width           =   135
         End
         Begin VB.CommandButton cmdAbilityLookup 
            Height          =   255
            Index           =   2
            Left            =   3300
            TabIndex        =   98
            Top             =   1080
            Width           =   135
         End
         Begin VB.CommandButton cmdAbilityLookup 
            Height          =   255
            Index           =   1
            Left            =   3300
            TabIndex        =   94
            Top             =   720
            Width           =   135
         End
         Begin VB.CommandButton cmdAbilityLookup 
            Height          =   255
            Index           =   0
            Left            =   3300
            TabIndex        =   90
            Top             =   360
            Width           =   135
         End
      End
      Begin VB.Label label 
         Alignment       =   1  'Right Justify
         Caption         =   "unknowns"
         Height          =   255
         Index           =   6
         Left            =   4500
         TabIndex        =   131
         Top             =   5460
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Type + and - in the '#' field to cycle though the abilities.  You can also type the name of the ability in the 'Name' field."
         Height          =   1215
         Left            =   4320
         TabIndex        =   127
         Top             =   3120
         Width           =   2655
      End
   End
   Begin VB.TextBox txtNumberSearch 
      Height          =   285
      Left            =   60
      TabIndex        =   2
      Top             =   540
      Width           =   615
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   540
      Width           =   2235
   End
   Begin MSComctlLib.ListView lvDatabase 
      Height          =   5055
      Left            =   60
      TabIndex        =   4
      Top             =   840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin exlimiter.EL EL1 
      Left            =   10440
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
   End
   Begin VB.Label lblNumberSearch 
      Caption         =   "#"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label7 
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
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmSpell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim nVSMR As Integer
Dim bDontCalc As Boolean
Dim bLoaded As Boolean
Dim nCurrentRecord As Long

Private Sub Form_Load()
Dim sCaption As String, j As Integer
On Error Resume Next

bLoaded = False

'cmbFilter.ListIndex = 0
sCaption = frmMain.Caption
frmMain.Caption = sCaption & " - Loading Spells ..."
DoEvents

With EL1
    .FormInQuestion = Me
    .MINHEIGHT = 425 + (TITLEBAR_OFFSET / 10)
    .MINWIDTH = 755
    .CenterOnLoad = False
    .EnableLimiter = True
End With

Me.Top = ReadINI("Windows", "SpellTop")
Me.Left = ReadINI("Windows", "SpellLeft")
Me.Width = ReadINI("Windows", "SpellWidth")
Me.Height = ReadINI("Windows", "SpellHeight")

nVSMR = ReadINI("Settings", "SpellVSMR")
If nVSMR < 1 Or nVSMR > 9999 Then nVSMR = 100
txtVSMR.Text = nVSMR

Call LoadAbilities

For j = 0 To 6
    cmbFilter(j).ListIndex = 0
    Call AutoSizeDropDownWidth(cmbFilter(j))
    Call ExpandCombo(cmbFilter(j), HeightOnly, TripleWidth, fraFilter2.hwnd)
Next j

For j = 3 To 9
    cmbFilterAbilityGL(j).ListIndex = 0
Next j

lvDatabase.ListItems.clear

Call LoadSpells

Me.Show
Me.SetFocus
txtSearch.SetFocus
frmMain.Caption = sCaption
If ReadINI("Windows", "SpellMaxed") = "1" Then Me.WindowState = vbMaximized

DoEvents

End Sub

Private Sub cmdFilterApply_Click()
Dim nStatus As Integer, bAdd As Boolean, x As Integer, bFiltered As Boolean
Dim z As Integer, bAbilMatch(4 To 6) As Boolean, nVal As Long

On Error GoTo error:

If bLoaded Then Call saverecord(nCurrentRecord)

nStatus = BTRCALL(BGETFIRST, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "BGETFIRST, Spell, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

Me.MousePointer = vbHourglass

bLoaded = False
lvDatabase.ListItems.clear

If chkFilterNone.Value = 1 Then
    Call LoadSpells
    Call cmdFilter_Click
    cmdFilter.Caption = "Filter"
    GoTo out:
End If

Do While nStatus = 0
    bAdd = True
    SpellRowToStruct Spelldatabuf.buf
    
    If chkFilter(0).Value = 1 And bAdd Then 'type
        If Not Spellrec.SpellType = cmbFilter(0).ListIndex Then bAdd = False
    End If
    If chkFilter(1).Value = 1 And bAdd Then 'target
        If Not Spellrec.Target = cmbFilter(1).ListIndex Then bAdd = False
    End If
    If chkFilter(2).Value = 1 And bAdd Then 'element
        If Not Spellrec.TypeOfAttack = cmbFilter(2).ListIndex Then bAdd = False
    End If
    If chkFilter(3).Value = 1 And bAdd Then 'magery
        If Spellrec.MageryA = cmbFilter(3).ListIndex Then
            If cmbFilterAbilityGL(3).ListIndex = 0 Then 'ANY
            ElseIf cmbFilterAbilityGL(3).ListIndex = 1 Then '<=
                If Spellrec.MageryB > Val(txtFilterAbilityValue(3).Text) Then bAdd = False
                If chkFilterExcludeZero.Value = 1 And Val(txtFilterAbilityValue(3).Text) = 0 Then bAdd = False
            ElseIf cmbFilterAbilityGL(3).ListIndex = 2 Then '>=
                If Spellrec.MageryB < Val(txtFilterAbilityValue(3).Text) Then bAdd = False
            Else '=
                If Not Spellrec.MageryB = Val(txtFilterAbilityValue(3).Text) Then bAdd = False
            End If
        Else
            bAdd = False
        End If
    End If

    If (chkFilter(5).Value = 1 Or chkFilter(6).Value = 1 Or chkFilter(4).Value = 1) And bAdd Then 'ability
        If chkFilter(5).Value = 1 Then bAbilMatch(5) = False Else bAbilMatch(5) = True
        If chkFilter(6).Value = 1 Then bAbilMatch(6) = False Else bAbilMatch(6) = True
        If chkFilter(4).Value = 1 Then bAbilMatch(4) = False Else bAbilMatch(4) = True
        
        For x = 0 To 9
            For z = 4 To 6
                If chkFilter(z).Value = 1 Then
                    If Spellrec.AbilityA(x) = cmbFilter(z).ItemData(cmbFilter(z).ListIndex) Then
                        nVal = Val(txtFilterAbilityValue(z).Text)
                        bAbilMatch(z) = True
                        
                        If cmbFilterAbilityGL(z).ListIndex = 0 Then  'ANY
                        ElseIf cmbFilterAbilityGL(z).ListIndex = 1 Then  '<=
                            If Spellrec.AbilityB(x) > nVal Then bAbilMatch(z) = False
                            If chkFilterExcludeZero.Value = 1 And Spellrec.AbilityB(x) = 0 Then bAbilMatch(z) = False
                        ElseIf cmbFilterAbilityGL(z).ListIndex = 2 Then  '>=
                            If Spellrec.AbilityB(x) < nVal Then bAbilMatch(z) = False
                        ElseIf cmbFilterAbilityGL(z).ListIndex = 3 Then  '=
                            If Not Spellrec.AbilityB(x) = nVal Then bAbilMatch(z) = False
                        End If
                        
                        If Not bAbilMatch(z) Then GoTo abil_out:
                    End If
                End If
            Next z
        Next x
abil_out:
        If Not (bAbilMatch(5) And bAbilMatch(6) And bAbilMatch(4)) Then bAdd = False
    End If
    
    If chkFilter(7).Value = 1 And bAdd Then 'energy
        nVal = Val(txtFilterAbilityValue(7).Text)
        If cmbFilterAbilityGL(7).ListIndex = 0 Then '<=
            If Spellrec.Energy > nVal Then bAdd = False
            If chkFilterExcludeZero.Value = 1 And Spellrec.Energy = 0 Then bAdd = False
        ElseIf cmbFilterAbilityGL(7).ListIndex = 1 Then '>=
            If Spellrec.Energy < nVal Then bAdd = False
        Else '=
            If Not Spellrec.Energy = nVal Then bAdd = False
        End If
    End If
    If chkFilter(8).Value = 1 And bAdd Then 'req level
        nVal = Val(txtFilterAbilityValue(8).Text)
        If cmbFilterAbilityGL(8).ListIndex = 0 Then '<=
            If Spellrec.Level > nVal Then bAdd = False
            If chkFilterExcludeZero.Value = 1 And Spellrec.Level = 0 Then bAdd = False
        ElseIf cmbFilterAbilityGL(8).ListIndex = 1 Then '>=
            If Spellrec.Level < nVal Then bAdd = False
        Else '=
            If Not Spellrec.Level = nVal Then bAdd = False
        End If
    End If
    If chkFilter(9).Value = 1 And bAdd Then 'cap
        nVal = Val(txtFilterAbilityValue(9).Text)
        If cmbFilterAbilityGL(9).ListIndex = 0 Then '<=
            If Spellrec.LevelCap > nVal Then bAdd = False
            If chkFilterExcludeZero.Value = 1 And Spellrec.LevelCap = 0 Then bAdd = False
        ElseIf cmbFilterAbilityGL(9).ListIndex = 1 Then '>=
            If Spellrec.LevelCap < nVal Then bAdd = False
        Else '=
            If Not Spellrec.LevelCap = nVal Then bAdd = False
        End If
    End If
    
    If chkFilter(10).Value = 1 And bAdd Then 'message
        nVal = Val(txtFilterMessage.Text)
        If Spellrec.CastMsgA = nVal Then GoTo msg_match:
        If Spellrec.CastMsgB = nVal Then GoTo msg_match:
        For x = 0 To 9
            Select Case Spellrec.AbilityA(x)
                Case 101, 115, 120:
                    If Spellrec.AbilityB(x) = nVal Then GoTo msg_match:
            End Select
        Next x
        bAdd = False
msg_match:
    End If
    
    If chkFilter(11).Value = 1 And bAdd Then 'textblock
        nVal = Val(txtFilterTB.Text)
        For x = 0 To 9
            If Spellrec.AbilityA(x) = 148 Then
                If Spellrec.AbilityB(x) = nVal Then GoTo tb_match:
            End If
        Next x
        bAdd = False
tb_match:
    End If
    
    If bAdd Then
        Call AddSpellToLV(Spellrec.Number)
    Else
        bFiltered = True
    End If
    
    nStatus = BTRCALL(BGETNEXT, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
Loop
If bFiltered Then
    cmdFilter.Caption = "*Filter*"
Else
    cmdFilter.Caption = "Filter"
End If

If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "Error: " & BtrieveErrorCode(nStatus)
    GoTo out:
End If

If lvDatabase.ListItems.Count > 0 Then
    SortListView lvDatabase, 1, ldtNumber, True
    If Not lvDatabase.SelectedItem Is Nothing Then lvDatabase.SelectedItem.Selected = False
    lvDatabase.ListItems(1).Selected = True
    lvDatabase.ListItems(1).EnsureVisible
    Call lvDatabase_ItemClick(lvDatabase.ListItems(1))
    Call cmdFilter_Click
Else
    MsgBox "No Records Matched.", vbInformation
End If

out:
Me.MousePointer = vbDefault
Exit Sub
error:
Call HandleError("cmdFilterApply_Click")
Resume out:

End Sub



Private Sub chkFilterNone_Click()
On Error Resume Next
Dim x As Integer, bAction As Boolean

If chkFilterNone.Value = 0 Then bAction = True

For x = 0 To 11
    chkFilter(x).Enabled = bAction
    If x <= 6 Then cmbFilter(x).Enabled = bAction
Next x

For x = 3 To 9
    cmbFilterAbilityGL(x).Enabled = bAction
    txtFilterAbilityValue(x).Enabled = bAction
Next x

txtFilterMessage.Enabled = bAction
txtFilterTB.Enabled = bAction

End Sub

Private Sub cmdAbilsClear_Click()
Dim x As Integer
On Error GoTo error:

For x = 0 To 9
    txtAbilityA(x).Text = 0
    txtAbilityB(x).Text = 0
Next x

out:
Exit Sub
error:
Call HandleError("cmdAbilsClear_Click")
Resume out:

End Sub

Private Sub cmdFilterCancel_Click()
    txtNumberSearch.Enabled = True
    fraFilter.Visible = False
    framNav.Enabled = True
    lvDatabase.Enabled = True
    txtSearch.Enabled = True
End Sub

Private Sub cmdFilterReset_Click()
Dim x As Integer

For x = 0 To 9
    chkFilter(x).Value = 0
    If x <= 6 Then cmbFilter(x).ListIndex = 0
Next x
For x = 3 To 9
    cmbFilterAbilityGL(x).ListIndex = 0
    txtFilterAbilityValue(x).Text = 0
Next x

txtFilterMessage.Text = 0
txtFilterTB.Text = 0

End Sub

Private Sub cmdFormulaCopy_Click(Index As Integer)
Dim bDur As Boolean, sCopy As String

If Index = 0 Then
    sSpellCopyPaste(0) = Val(txtEnergy.Text)
    sSpellCopyPaste(1) = Val(txtLevel.Text)
    sSpellCopyPaste(2) = Val(txtLevelCap.Text)
    sSpellCopyPaste(3) = Val(txtMin.Text)
    sSpellCopyPaste(4) = Val(txtMax.Text)
    sSpellCopyPaste(5) = Val(txtDuration.Text)
    sSpellCopyPaste(6) = Val(txtLVLSMinIncr.Text)
    sSpellCopyPaste(7) = Val(txtMinIncrease.Text)
    sSpellCopyPaste(8) = Val(txtLVLSMaxIncr.Text)
    sSpellCopyPaste(9) = Val(txtMaxIncrease.Text)
    sSpellCopyPaste(10) = Val(txtLVLSDurIncr.Text)
    sSpellCopyPaste(11) = Val(txtDurIncrease.Text)
    
    If (Val(sSpellCopyPaste(10)) <> 0 And Val(sSpellCopyPaste(11)) <> 0) _
        Or Val(sSpellCopyPaste(5)) <> 0 Then bDur = True
        
    sCopy = sSpellCopyPaste(1) & "/" & sSpellCopyPaste(2) & "/" _
        & sSpellCopyPaste(3) & "/" & sSpellCopyPaste(4) _
        & IIf(bDur, "/" & sSpellCopyPaste(5), "") & ", " _
        & sSpellCopyPaste(6) & "/" & sSpellCopyPaste(7) & "/" _
        & sSpellCopyPaste(8) & "/" & sSpellCopyPaste(9) _
        & IIf(bDur, "/" & sSpellCopyPaste(10) & "/" & sSpellCopyPaste(11), "")
    
    sCopy = sCopy & " -- @Req " & txtAtReq.Text & " -- @Cap " & txtAtCap.Text
    
    Clipboard.clear
    Clipboard.SetText sCopy
ElseIf Index = 1 Then
    txtEnergy.Text = sSpellCopyPaste(0)
    txtLevel.Text = sSpellCopyPaste(1)
    txtLevelCap.Text = sSpellCopyPaste(2)
    txtMin.Text = sSpellCopyPaste(3)
    txtMax.Text = sSpellCopyPaste(4)
    txtDuration.Text = sSpellCopyPaste(5)
    txtLVLSMinIncr.Text = sSpellCopyPaste(6)
    txtMinIncrease.Text = sSpellCopyPaste(7)
    txtLVLSMaxIncr.Text = sSpellCopyPaste(8)
    txtMaxIncrease.Text = sSpellCopyPaste(9)
    txtLVLSDurIncr.Text = sSpellCopyPaste(10)
    txtDurIncrease.Text = sSpellCopyPaste(11)
Else
    MsgBox "Clicking copy will copy the data in the fields above to a memory location " _
        & "within NMR as well as copy the formula data to the windows clipboard." _
        & "The paste button will only paste what's in the NMR memory location.", vbInformation
End If

End Sub

Private Sub LoadAbilities()
Dim x As Integer
On Error GoTo error:

For x = 4 To 6
    cmbFilter(x).clear
Next x
rsAbilities.MoveFirst
Do Until rsAbilities.EOF
    If Not rsAbilities.Fields("Number") = 0 Then
        For x = 4 To 6
            cmbFilter(x).AddItem rsAbilities.Fields("Name") & " (" & rsAbilities.Fields("Number") & ")"
            cmbFilter(x).ItemData(cmbFilter(x).NewIndex) = rsAbilities.Fields("Number")
        Next x
    End If
    rsAbilities.MoveNext
Loop

For x = 4 To 6
    cmbFilter(x).AddItem "None (0)", 0
    cmbFilter(x).ListIndex = 0
Next x

out:
Exit Sub
error:
Call HandleError("LoadAbilities")
Resume out:

End Sub

Private Sub cmdAbilityLookup_Click(Index As Integer)
    Call LookupAbility(Val(txtAbilityA(Index).Text), Val(txtAbilityB(Index).Text))
End Sub

Private Sub cmdFilter_Click()
If fraFilter.Visible Then
    txtNumberSearch.Enabled = True
    framNav.Enabled = True
    lvDatabase.Enabled = True
    txtSearch.Enabled = True
    fraFilter.Visible = False
Else
    txtNumberSearch.Enabled = False
    framNav.Enabled = False
    lvDatabase.Enabled = False
    txtSearch.Enabled = False
    fraFilter.Visible = True
End If
'On Error GoTo Error:
'Dim nStatus As Integer
'
'If bLoaded Then Call saverecord(nCurrentRecord)
'
'nStatus = BTRCALL(BGETFIRST, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
'If Not nStatus = 0 Then
'    MsgBox "LoadSpell, BGETFIRST, Spell, Error: " & BtrieveErrorCode(nStatus)
'    Exit Sub
'End If
'
'lvDatabase.ListItems.clear
'
'If cmbFilter.ListIndex = 0 Then
'    Call LoadSpells
'    Exit Sub
'End If
'
'Do While nStatus = 0
'    SpellRowToStruct Spelldatabuf.buf
'
'    If Spellrec.MageryA = cmbFilter.ListIndex - 1 Then Call AddSpellToLV(Spellrec.Number)
'
'    nStatus = BTRCALL(BGETNEXT, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
'Loop
'If Not nStatus = 0 And Not nStatus = 9 Then
'    MsgBox "LoadSpells, Error: " & BtrieveErrorCode(nStatus)
'End If
'
'If lvDatabase.ListItems.Count >= 1 Then Call lvDatabase_ItemClick(lvDatabase.ListItems(1))
'
'lvDatabase.refresh
'SortListView lvDatabase, 1, ldtNumber, True
'bLoaded = True
'
'Exit Sub
'Error:
'Call HandleError
End Sub

Private Sub cmdDiscard_Click()
Dim nStatus As Integer

If lvDatabase.SelectedItem Is Nothing Or nCurrentRecord = 0 Then
    MsgBox "No current record."
    Exit Sub
End If

nStatus = BTRCALL(BGETEQUAL, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), nCurrentRecord, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error on BGETEQUAL: " & BtrieveErrorCode(nStatus)
Else
    DispSpellInfo Spelldatabuf.buf
End If
End Sub

Private Sub cmdEditMsgA_Click()
Call frmMessage.GotoMSG(Val(txtCastMsgA.Text))

End Sub

Private Sub cmdEditMsgB_Click()
Call frmMessage.GotoMSG(Val(txtCastMsgB.Text))

End Sub

Private Sub cmdSave_Click()
On Error GoTo error:

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
If lvDatabase.SelectedItem Is Nothing Then Exit Sub

Call saverecord(nCurrentRecord)
'Call lvDatabase_ItemClick(lvDatabase.SelectedItem)

Dim oLI As ListItem
Set oLI = lvDatabase.FindItem(Spellrec.Number, lvwText, , 0)
If Not oLI Is Nothing Then
    oLI.ListSubItems(1).Text = ClipNull(Spellrec.Name)
    If Not bOnlyNames Then
        oLI.ListSubItems(2).Text = GetMagery(Spellrec.MageryA, Spellrec.MageryB)
        oLI.ListSubItems(3).Text = Spellrec.ShortName
        oLI.ListSubItems(4).Text = Spellrec.Level
    End If
End If
Set oLI = Nothing

out:
Exit Sub
error:
Call HandleError("cmdSave_Click")
Resume out:

End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = vbMinimized Then Exit Sub
framNav.Left = Me.Width - framNav.Width - 200
lvDatabase.Width = framNav.Left - 175
lvDatabase.Height = Me.Height - 1325 - TITLEBAR_OFFSET
End Sub

Private Sub lblName_GotFocus(Index As Integer)
Call SelectAll(lblName(Index))

End Sub

Private Sub lblName_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Call FindAbilityName(KeyCode, txtAbilityA(Index), lblName(Index))

End Sub

Private Sub lblName_LostFocus(Index As Integer)
Call txtAbilityA_Change(Index)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If bLoaded = True Then Call saverecord(nCurrentRecord)
Call WriteINI("Settings", "SpellVSMR", nVSMR)
If Me.WindowState = vbMinimized Then Exit Sub

If Me.WindowState = vbMaximized Then
    Call WriteINI("Windows", "SpellMaxed", 1)
Else
    Call WriteINI("Windows", "SpellMaxed", 0)
    Call WriteINI("Windows", "SpellTop", Me.Top)
    Call WriteINI("Windows", "SpellLeft", Me.Left)
    Call WriteINI("Windows", "SpellHeight", Me.Height)
    Call WriteINI("Windows", "SpellWidth", Me.Width)
End If
End Sub

Private Sub lvDatabase_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim nSort As ListDataType
Select Case ColumnHeader.Index
    Case 1, 5: nSort = ldtNumber
    Case Else: nSort = ldtString
End Select
SortListView lvDatabase, ColumnHeader.Index, nSort, lvDatabase.SortOrder
End Sub

Private Sub lvDatabase_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim temp As Long, nStatus As Integer

If bLoaded = True And chkAutoSave.Value = 1 Then Call saverecord(nCurrentRecord)

temp = Val(Item.Text)
nStatus = BTRCALL(BGETEQUAL, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), temp, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error on BGETEQUAL: " & BtrieveErrorCode(nStatus)
    bLoaded = False
Else
    nCurrentRecord = temp
    DispSpellInfo Spelldatabuf.buf
    bLoaded = True
End If
End Sub



Private Sub txtAbilityA_Change(Index As Integer)

If Me.ActiveControl Is lblName(Index) Then Exit Sub
Call FindAbilityNumber(txtAbilityA(Index), lblName(Index))

If Val(txtAbilityA(Index).Text) = 17 Then Call CalcValues

End Sub

Private Sub txtAbilityA_GotFocus(Index As Integer)
Call SelectAll(txtAbilityA(Index))

End Sub

Private Sub txtAbilityA_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 43
            txtAbilityA(Index).Text = Val(txtAbilityA(Index).Text) + 1
        Case 61
            txtAbilityA(Index).Text = Val(txtAbilityA(Index).Text) + 1
        Case 45
            txtAbilityA(Index).Text = Val(txtAbilityA(Index).Text) - 1
    End Select
End Sub

Private Sub LoadSpells()
On Error GoTo error:
Dim nStatus As Integer

lvDatabase.ColumnHeaders.clear
lvDatabase.ColumnHeaders.add 1, "Number", "#", 600, lvwColumnLeft
lvDatabase.ColumnHeaders.add 2, "Name", "Name", 1900, lvwColumnCenter
If Not bOnlyNames Then
    lvDatabase.ColumnHeaders.add 3, "Magery", "Magery", 1000, lvwColumnCenter
    lvDatabase.ColumnHeaders.add 4, "Short", "Short", 700, lvwColumnCenter
    lvDatabase.ColumnHeaders.add 5, "Level", "Level", 600, lvwColumnCenter
    lvDatabase.ColumnHeaders.add 6, "Diff", "Diff", 600, lvwColumnCenter
End If

nStatus = BTRCALL(BGETFIRST, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "LoadSpells, BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

Do While nStatus = 0
    SpellRowToStruct Spelldatabuf.buf
    
    Call AddSpellToLV(Spellrec.Number)
    
    nStatus = BTRCALL(BGETNEXT, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "LoadSpells, Error: " & BtrieveErrorCode(nStatus)
End If

If lvDatabase.ListItems.Count >= 1 Then
    lvDatabase.refresh
    If bOppositeListOrder Then
        SortListView lvDatabase, 1, ldtNumber, False
    Else
        SortListView lvDatabase, 1, ldtNumber, True
    End If
    Call lvDatabase_ItemClick(lvDatabase.ListItems(1))
End If

bLoaded = True

Exit Sub
error:
Call HandleError
End Sub
Private Sub AddSpellToLV(ByVal nNumber As Integer)
Dim nStatus As Integer, oLI As ListItem
On Error GoTo error:

If Not nNumber = Spellrec.Number Then
    nStatus = BTRCALL(BGETEQUAL, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), nNumber, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then MsgBox "Error getting record " & nNumber & ": " & BtrieveErrorCode(nStatus)
    bLoaded = False
    Exit Sub
End If

Set oLI = lvDatabase.ListItems.add()
oLI.Text = Spellrec.Number

oLI.ListSubItems.add (1), "Name", ClipNull(Spellrec.Name)
If Not bOnlyNames Then
    oLI.ListSubItems.add (2), "Magery", GetMagery(Spellrec.MageryA, Spellrec.MageryB)
    oLI.ListSubItems.add (3), "Short", Spellrec.ShortName
    oLI.ListSubItems.add (4), "Level", Spellrec.Level
    oLI.ListSubItems.add (5), "Diff", Spellrec.Difficulty
End If

Set oLI = Nothing
Exit Sub
error:
Call HandleError
Set oLI = Nothing
End Sub

Private Sub DispSpellInfo(row() As Byte)
On Error GoTo error:
Dim x As Integer

Call SpellRowToStruct(row())

Me.Caption = "Spell Editor -- " & ClipNull(Spellrec.Name)

bDontCalc = True

txtNumber.Text = SInt2UInt(Spellrec.Number)
txtName.Text = Spellrec.Name
txtDesc(0).Text = Spellrec.DescA
txtDesc(1).Text = Spellrec.DescB
txtCastMsgA.Text = Spellrec.CastMsgA
txtEnergy.Text = Spellrec.Energy
txtLevel.Text = Spellrec.Level
txtMin.Text = Spellrec.Min
txtMax.Text = Spellrec.Max
txtLevelCap.Text = Spellrec.LevelCap
txtCastMsgB.Text = Spellrec.CastMsgB
txtDifficulty.Text = Spellrec.Difficulty
txtMana.Text = Spellrec.Mana
txtLVLSMaxIncr.Text = Spellrec.LVLSMaxIncr
txtMaxIncrease.Text = Spellrec.MaxIncrease
txtShortName.Text = Spellrec.ShortName
txtDuration.Text = Spellrec.duration
txtMageryB.Text = Spellrec.MageryB
cmbTypeOfResists.ListIndex = Spellrec.TypeOfResists
cmbSpellType.ListIndex = Spellrec.SpellType
cmbTarget.ListIndex = Spellrec.Target
cmbResistAbility.ListIndex = Spellrec.ResistAbility
cmbTypeOfAttack.ListIndex = Spellrec.TypeOfAttack
cmbMageryA.ListIndex = Spellrec.MageryA
'txtCastADisplay.Text = GetMessages(Spellrec.CastMsgA, 1)
'txtCastBDisplay.Text = GetMessages(Spellrec.CastMsgB, 2)

If Spellrec.MsgStyle > 1 Then
    Select Case Spellrec.MsgStyle
        Case Is = 32
            Spellrec.MsgStyle = 0
        Case Is = 33
            Spellrec.MsgStyle = 1
        Case Is = 114
            Spellrec.MsgStyle = 0
        Case Is = 115
            Spellrec.MsgStyle = 1
     End Select
End If

cmbMsgStyle.ListIndex = Spellrec.MsgStyle

For x = 0 To 9
    txtAbilityA(x).Text = Spellrec.AbilityA(x)
    txtAbilityB(x).Text = Spellrec.AbilityB(x)
Next

txtUNDEFINED01.Text = Spellrec.UNDEFINED01
txtUNDEFINED02.Text = Spellrec.UNDEFINED02
txtMinIncrease.Text = Spellrec.MinIncrease
txtLVLSMinIncr.Text = Spellrec.LVLSMinIncr
txtLVLSDurIncr.Text = Spellrec.LVLSDurIncr
txtDurIncrease.Text = Spellrec.DurIncrease

bDontCalc = False
Call CalcValues

Exit Sub
error:
Call HandleError
MsgBox "Warning, record was not completely displayed." & vbCrLf _
    & "Previous records stats may still be in memory.  Select 'Disable DB Writing'" & vbCrLf _
    & "from the file menu and then reload the editor.", vbExclamation
End Sub


Private Sub saverecord(ByVal nRecord As Long)
On Error GoTo error:
Dim nStatus As Integer, x As Integer

If nRecord = 0 Then Exit Sub

nStatus = BTRCALL(BGETEQUAL, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), nRecord, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Save Error on BGETEQUAL: " & BtrieveErrorCode(nStatus)
    Exit Sub
Else
    SpellRowToStruct Spelldatabuf.buf
End If

'DoEvents
Spellrec.Name = RTrim(txtName.Text) & Chr(0)
Spellrec.DescA = Trim(txtDesc(0).Text) & Chr(0)
Spellrec.DescB = Trim(txtDesc(1).Text) & Chr(0)
Spellrec.CastMsgA = Val(txtCastMsgA.Text)
Spellrec.Energy = Val(txtEnergy.Text)
Spellrec.Level = Val(txtLevel.Text)
Spellrec.Min = Val(txtMin.Text)
Spellrec.Max = Val(txtMax.Text)
Spellrec.LevelCap = Val(txtLevelCap.Text)
Spellrec.CastMsgB = Val(txtCastMsgB.Text)
Spellrec.Difficulty = Val(txtDifficulty.Text)
Spellrec.Mana = Val(txtMana.Text)
Spellrec.LVLSMaxIncr = Val(txtLVLSMaxIncr.Text)
Spellrec.MaxIncrease = Val(txtMaxIncrease.Text)
Spellrec.ShortName = RTrim(txtShortName.Text) & Chr(0)
Spellrec.duration = Val(txtDuration.Text)
Spellrec.MageryB = Val(txtMageryB.Text)
Spellrec.TypeOfResists = cmbTypeOfResists.ListIndex
Spellrec.SpellType = cmbSpellType.ListIndex
Spellrec.Target = cmbTarget.ListIndex
Spellrec.ResistAbility = cmbResistAbility.ListIndex
Spellrec.TypeOfAttack = cmbTypeOfAttack.ListIndex
Spellrec.MageryA = cmbMageryA.ListIndex
Spellrec.MsgStyle = cmbMsgStyle.ListIndex

For x = 0 To 9
    Spellrec.AbilityA(x) = Val(txtAbilityA(x).Text)
    Spellrec.AbilityB(x) = Val(txtAbilityB(x).Text)
Next

Spellrec.UNDEFINED01 = Val(txtUNDEFINED01.Text)
Spellrec.UNDEFINED02 = Val(txtUNDEFINED02.Text)
Spellrec.MinIncrease = Val(txtMinIncrease.Text)
Spellrec.LVLSMinIncr = Val(txtLVLSMinIncr.Text)
Spellrec.LVLSDurIncr = Val(txtLVLSDurIncr.Text)
Spellrec.DurIncrease = Val(txtDurIncrease.Text)

nStatus = UpdateSpell
If Not nStatus = 0 Then
    MsgBox "SaveRecord, Error: " & BtrieveErrorCode(nStatus)
Else
    DispSpellInfo Spelldatabuf.buf
End If

Exit Sub
error:
Call HandleError
End Sub


Private Sub txtAbilityB_GotFocus(Index As Integer)
Call SelectAll(txtAbilityB(Index))

End Sub

Private Sub txtAtCap_GotFocus()
Call SelectAll(txtAtCap)

End Sub

Private Sub txtAtReq_GotFocus()
Call SelectAll(txtAtReq)

End Sub

Private Sub txtCastMsgA_Change()
txtCastADisplay.Text = GetMessages(Val(txtCastMsgA.Text), 1)
End Sub

Private Sub txtCastMsgA_GotFocus()
Call SelectAll(txtCastMsgA)

End Sub

Private Sub txtCastMsgB_Change()
txtCastBDisplay.Text = GetMessages(Val(txtCastMsgB.Text), 2)
End Sub

Private Sub txtCastMsgB_GotFocus()
Call SelectAll(txtCastMsgB)

End Sub


Private Sub txtDesc_Change(Index As Integer)
If Index = 1 Then Exit Sub
If txtDesc(Index).SelStart = txtDesc(Index).MaxLength Then
    txtDesc(Index).Text = Trim(txtDesc(Index).Text)
    txtDesc(Index + 1).SetFocus
    DoEvents
    txtDesc(Index + 1).SelStart = 0
End If
End Sub

Private Sub txtDesc_GotFocus(Index As Integer)
Call SelectAll(txtDesc(Index))

End Sub

Private Sub txtDifficulty_GotFocus()
Call SelectAll(txtDifficulty)

End Sub

Private Sub txtDuration_Change()
Call CalcValues
End Sub

Private Sub txtDuration_GotFocus()
Call SelectAll(txtDuration)

End Sub

Private Sub txtDurIncrease_Change()
Call CalcValues
End Sub

Private Sub txtDurIncrease_GotFocus()
Call SelectAll(txtDurIncrease)

End Sub

Private Sub txtEnergy_Change()
Call CalcValues
End Sub

Private Sub txtEnergy_GotFocus()
Call SelectAll(txtEnergy)

End Sub

Private Sub txtFilterAbilityValue_GotFocus(Index As Integer)
Call SelectAll(txtFilterAbilityValue(Index))
End Sub

Private Sub txtFilterAbilityValue_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtFilterMessage_GotFocus()
Call SelectAll(txtFilterMessage)
End Sub

Private Sub txtFilterMessage_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtFilterTB_GotFocus()
Call SelectAll(txtFilterTB)
End Sub

Private Sub txtFilterTB_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtLevel_Change()
Call CalcValues
End Sub

Private Sub txtLevel_GotFocus()
Call SelectAll(txtLevel)

End Sub

Private Sub txtLevelCap_Change()
Call CalcValues
End Sub

Private Sub txtLevelCap_GotFocus()
Call SelectAll(txtLevelCap)

End Sub

Private Sub txtLVLSDurIncr_Change()
Call CalcValues
End Sub

Private Sub txtLVLSDurIncr_GotFocus()
Call SelectAll(txtLVLSDurIncr)

End Sub

Private Sub txtLVLSMaxIncr_Change()
Call CalcValues
End Sub

Private Sub txtLVLSMaxIncr_GotFocus()
Call SelectAll(txtLVLSMaxIncr)

End Sub

Private Sub txtLVLSMinIncr_Change()
Call CalcValues
End Sub

Private Sub txtLVLSMinIncr_GotFocus()
Call SelectAll(txtLVLSMinIncr)

End Sub

Private Sub txtMageryB_GotFocus()
Call SelectAll(txtMageryB)

End Sub

Private Sub txtMana_GotFocus()
Call SelectAll(txtMana)

End Sub

Private Sub txtMax_Change()
Call CalcValues
End Sub

Private Sub txtMax_GotFocus()
Call SelectAll(txtMax)

End Sub

Private Sub txtMaxIncrease_Change()
Call CalcValues
End Sub

Private Sub txtMaxIncrease_GotFocus()
Call SelectAll(txtMaxIncrease)

End Sub

Private Sub txtMin_Change()
Call CalcValues
End Sub
Private Sub CalcValues()
On Error GoTo error:
Dim nMin As Currency, nMax As Currency, nDur As Currency, nEnergy As Currency, x As Integer
Dim nMRVal As Currency

If bDontCalc Then Exit Sub

If nVSMR > 0 Then
    For x = 0 To 9
        If Val(txtAbilityA(x).Text) = 17 Then
            nMRVal = ((nVSMR - 50) / 2) / 100
            Exit For
        End If
    Next x
Else
    nMRVal = -1
End If

If x = 10 Or nMRVal < 0 Then
    txtMRvsReq.Text = ""
    txtMRvsCap.Text = ""
    nMRVal = -1
End If

nEnergy = Val(txtEnergy.Text)
If nEnergy < 1000 And nEnergy > 0 Then
    nEnergy = Fix(1000 / nEnergy)
    If nEnergy < 2 Then nEnergy = 0
Else
    nEnergy = 0
End If

If Val(txtLVLSMinIncr.Text) = 0 Then
    nMin = Val(txtMin.Text)
Else
    nMin = Val(txtMin.Text) + Fix((Val(txtMinIncrease.Text) / Val(txtLVLSMinIncr.Text)) * Val(txtLevel.Text))
End If
If Val(txtLVLSMaxIncr.Text) = 0 Then
    nMax = Val(txtMax.Text)
Else
    nMax = Val(txtMax.Text) + Fix((Val(txtMaxIncrease.Text) / Val(txtLVLSMaxIncr.Text)) * Val(txtLevel.Text))
End If
If Val(txtLVLSDurIncr.Text) = 0 Then
    nDur = Val(txtDuration.Text)
Else
    nDur = Val(txtDuration.Text) + Fix((Val(txtDurIncrease.Text) / Val(txtLVLSDurIncr.Text)) * Val(txtLevel.Text))
End If

If nEnergy > 0 Then
    nMin = nMin * nEnergy
    nMax = nMax * nEnergy
End If
txtAtReq.Text = IIf(nEnergy > 0, "(x" & nEnergy & "): ", "") _
    & nMin & " to " & nMax & IIf(nDur <> 0, ", " & nDur & " rnds", "")


If nMRVal > 0 Then
    txtMRvsReq.Text = IIf(nEnergy > 0, "(x" & nEnergy & "): ", "") _
        & (nMin - Fix(nMin * nMRVal)) & " to " _
        & (nMax - Fix(nMax * nMRVal)) _
        & IIf(nDur <> 0, ", " & nDur & " rnds", "")
End If

If Val(txtLVLSMinIncr.Text) = 0 Then
    nMin = Val(txtMin.Text)
Else
    nMin = Val(txtMin.Text) + Fix((Val(txtMinIncrease.Text) / Val(txtLVLSMinIncr.Text)) * Val(txtLevelCap.Text))
End If
If Val(txtLVLSMaxIncr.Text) = 0 Then
    nMax = Val(txtMax.Text)
Else
    nMax = Val(txtMax.Text) + Fix((Val(txtMaxIncrease.Text) / Val(txtLVLSMaxIncr.Text)) * Val(txtLevelCap.Text))
End If
If Val(txtLVLSDurIncr.Text) = 0 Then
    nDur = Val(txtDuration.Text)
Else
    nDur = Val(txtDuration.Text) + Fix((Val(txtDurIncrease.Text) / Val(txtLVLSDurIncr.Text)) * Val(txtLevelCap.Text))
End If

If nEnergy > 0 Then
    nMin = nMin * nEnergy
    nMax = nMax * nEnergy
End If
txtAtCap.Text = IIf(nEnergy > 0, "(x" & nEnergy & "): ", "") _
    & nMin & " to " & nMax & IIf(nDur <> 0, ", " & nDur & " rnds", "")

If nMRVal > 0 Then
    txtMRvsCap.Text = IIf(nEnergy > 0, "(x" & nEnergy & "): ", "") _
        & (nMin - Fix(nMin * nMRVal)) & " to " _
        & (nMax - Fix(nMax * nMRVal)) _
        & IIf(nDur <> 0, ", " & nDur & " rnds", "")
End If

''''''''''''''''''''''''
'If Val(txtLVLSMinIncr.Text) = 0 Then
'    txtMinMin.Text = Val(txtMin.Text)
'    txtMinMax.Text = Val(txtMin.Text)
'Else
'    txtMinMin.Text = Val(txtMin.Text) + Round((Val(txtMinIncrease.Text) / Val(txtLVLSMinIncr.Text)) * Val(txtLevel.Text), 2)
'    txtMinMax.Text = Val(txtMin.Text) + Round((Val(txtMinIncrease.Text) / Val(txtLVLSMinIncr.Text)) * Val(txtLevelCap.Text), 2)
'End If
'
'If Val(txtLVLSMaxIncr.Text) = 0 Then
'    txtMaxMin.Text = Val(txtMax.Text)
'    txtMaxMax.Text = Val(txtMax.Text)
'Else
'    txtMaxMin.Text = Val(txtMax.Text) + Round((Val(txtMaxIncrease.Text) / Val(txtLVLSMaxIncr.Text)) * Val(txtLevel.Text), 2)
'    txtMaxMax.Text = Val(txtMax.Text) + Round((Val(txtMaxIncrease.Text) / Val(txtLVLSMaxIncr.Text)) * Val(txtLevelCap.Text), 2)
'End If
'
'If Val(txtLVLSDurIncr.Text) = 0 Then
'    txtDurMin.Text = Val(txtDuration.Text)
'    txtDurMax.Text = Val(txtDuration.Text)
'Else
'    txtDurMin.Text = Val(txtDuration.Text) + Round((Val(txtDurIncrease.Text) / Val(txtLVLSDurIncr.Text)) * Val(txtLevel.Text), 2)
'    txtDurMax.Text = Val(txtDuration.Text) + Round((Val(txtDurIncrease.Text) / Val(txtLVLSDurIncr.Text)) * Val(txtLevelCap.Text), 2)
'End If

out:
Exit Sub
error:
Call HandleError("CalcValues")
Resume out:
End Sub

Private Sub txtMin_GotFocus()
Call SelectAll(txtMin)

End Sub

Private Sub txtMinIncrease_Change()
Call CalcValues
End Sub

Private Sub txtMinIncrease_GotFocus()
Call SelectAll(txtMinIncrease)

End Sub

Private Sub txtMRvsCap_GotFocus()
Call SelectAll(txtMRvsCap)

End Sub

Private Sub txtMRvsReq_GotFocus()
Call SelectAll(txtMRvsReq)

End Sub

Private Sub txtName_GotFocus()
Call SelectAll(txtName)

End Sub

Private Sub txtNumber_GotFocus()
Call SelectAll(txtNumber)

End Sub

Private Sub txtNumberSearch_GotFocus()
Call SelectAll(txtNumberSearch)

End Sub

Private Sub txtNumberSearch_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)

End Sub

Private Sub txtNumberSearch_KeyUp(KeyCode As Integer, Shift As Integer)
Dim x As Long, SearchStart As Long

If txtNumberSearch.Text = "" Then Exit Sub
If lvDatabase.ListItems.Count < 1 Then Exit Sub

SearchStart = 1

If KeyCode = vbKeyUp Then Exit Sub
If KeyCode = vbKeyDown Then lvDatabase.SetFocus
If KeyCode = vbKeyLeft Then Exit Sub
If KeyCode = vbKeyRight Then SearchStart = lvDatabase.SelectedItem.Index + 1
If KeyCode = vbKeyControl Then Exit Sub 'control
If KeyCode = 18 Then Exit Sub 'alt
If KeyCode = vbKeyTab Then Exit Sub 'tab
If KeyCode = vbKeyShift Then Exit Sub

For x = SearchStart To lvDatabase.ListItems.Count
    If Val(lvDatabase.ListItems(x).Text) = Val(txtNumberSearch.Text) Then
        Set lvDatabase.SelectedItem = lvDatabase.ListItems(x)
        lvDatabase.SelectedItem.EnsureVisible
        Call lvDatabase_ItemClick(lvDatabase.SelectedItem)
        Exit For
    End If
Next x

End Sub

Private Sub txtSearch_GotFocus()
Call SelectAll(txtSearch)

End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
Dim x As Long, SearchStart As Long

If txtSearch.Text = "" Then Exit Sub
If lvDatabase.ListItems.Count < 1 Then Exit Sub

SearchStart = 1

If KeyCode = vbKeyUp Then Exit Sub
If KeyCode = vbKeyDown Then lvDatabase.SetFocus
If KeyCode = vbKeyLeft Then Exit Sub
If KeyCode = vbKeyRight Then SearchStart = lvDatabase.SelectedItem.Index + 1
If KeyCode = vbKeyControl Then Exit Sub 'control
If KeyCode = 18 Then Exit Sub 'alt
If KeyCode = vbKeyTab Then Exit Sub 'tab
If KeyCode = vbKeyShift Then Exit Sub

For x = SearchStart To lvDatabase.ListItems.Count
    If Not InStr(1, LCase(lvDatabase.ListItems(x).ListSubItems(1)), LCase(txtSearch.Text)) = 0 Then
        Set lvDatabase.SelectedItem = lvDatabase.ListItems(x)
        lvDatabase.SelectedItem.EnsureVisible
        Call lvDatabase_ItemClick(lvDatabase.SelectedItem)
        Exit For
    End If
Next x
End Sub

Private Sub cmdDelete_Click()
On Error GoTo error:
Dim nStatus As Integer
Dim nDelete As Integer, temp As Long

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

nDelete = MsgBox("Delete this record from database?", vbYesNo, "Delete Record?")

If bLoaded Then Call saverecord(nCurrentRecord)

If Not nDelete = vbYes Then Exit Sub
    
nCurrentRecord = Val(lvDatabase.SelectedItem.Text)
temp = lvDatabase.SelectedItem.Index

nStatus = BTRCALL(BGETEQUAL, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), nCurrentRecord, KEY_BUF_LEN, 0)
If nStatus = 0 Then
    nStatus = BTRCALL(BDELETE, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "cmdDelete, BDELETE, Error: " & BtrieveErrorCode(nStatus)
    Else
        lvDatabase.ListItems.Remove temp
        nCurrentRecord = 0
        bLoaded = False
        If lvDatabase.ListItems.Count >= 1 Then
            If temp > 1 Then temp = temp - 1 Else temp = 1
            Set lvDatabase.SelectedItem = lvDatabase.ListItems(temp)
            lvDatabase.SelectedItem.EnsureVisible
            Call lvDatabase_ItemClick(lvDatabase.SelectedItem)
        Else
            Call Form_Unload(1)
            Call Form_Load
        End If
    End If
Else
    MsgBox "Couldn't get record, Error: " & BtrieveErrorCode(nStatus)
End If

Exit Sub
error:
Call HandleError
End Sub

Private Sub cmdInsert_Click()
On Error GoTo error:
Dim nStatus As Integer
Dim nNewSpellNumber As String, oLI As ListItem

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

If bLoaded = True Then Call saverecord(nCurrentRecord)

nNewSpellNumber = InputBox("New Spell Number:" & vbCrLf & vbCrLf & "Enter 0 for the next highest number.", "Insert", "0")
If nNewSpellNumber = "" Then Exit Sub

Spellrec.Number = Val(nNewSpellNumber)
'Spellrec.Name = "New Spell" & Chr(0)
Call SpellStructToRow(Spelldatabuf.buf)

nStatus = BTRCALL(BINSERT, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdInsert, BINSERT, Error: " & BtrieveErrorCode(nStatus)
    bLoaded = False
Else
    SpellRowToStruct Spelldatabuf.buf
    
    Call AddSpellToLV(Spellrec.Number)
    
    nCurrentRecord = Spellrec.Number
    DispSpellInfo Spelldatabuf.buf
    
    SortListView lvDatabase, 1, ldtNumber, True
    
    Set oLI = lvDatabase.FindItem(Spellrec.Number, lvwText, , 0)
    If Not oLI Is Nothing Then
        Set lvDatabase.SelectedItem = oLI
        lvDatabase.SelectedItem.EnsureVisible
        Call lvDatabase_ItemClick(lvDatabase.SelectedItem)
        Set oLI = Nothing
    Else
        Set lvDatabase.SelectedItem = lvDatabase.ListItems(lvDatabase.ListItems.Count)
        lvDatabase.SelectedItem.EnsureVisible
        Call lvDatabase_ItemClick(lvDatabase.SelectedItem)
    End If
End If

Set oLI = Nothing
Exit Sub
error:
Call HandleError
Set oLI = Nothing
End Sub


Private Sub cmdOther_Click()
If frmAbilities.Visible = False Then
    frmAbilities.Visible = True
    frmGeneral.Visible = False
    cmdOther.Caption = "&General"
Else
    frmAbilities.Visible = False
    frmGeneral.Visible = True
    cmdOther.Caption = "&Abilities"
End If
End Sub
Public Sub GotoSpell(ByVal nRecnum As Integer)
Dim oLI As ListItem

Set oLI = lvDatabase.FindItem(nRecnum, lvwText, 1, 0)

If Not oLI Is Nothing Then
    Set lvDatabase.SelectedItem = oLI
    lvDatabase.SelectedItem.EnsureVisible
    Call lvDatabase_ItemClick(oLI)
End If

Set oLI = Nothing
Me.Show
Me.SetFocus
End Sub

Private Sub txtShortName_GotFocus()
Call SelectAll(txtShortName)

End Sub

Private Sub txtVSMR_Change()
On Error GoTo error:

If Val(txtVSMR.Text) < 0 Or Val(txtVSMR.Text) > 9999 Then
    txtVSMR.Text = 0
    Call SelectAll(txtVSMR)
    Exit Sub
End If

nVSMR = Val(txtVSMR.Text)
lblVSMR(0).Caption = "@Req vs " & nVSMR & "MR"
lblVSMR(1).Caption = "@Cap vs " & nVSMR & "MR"
Call CalcValues

Exit Sub
error:
Call HandleError("txtVSMR_Change")

End Sub

Private Sub txtVSMR_GotFocus()
Call SelectAll(txtVSMR)
End Sub

Private Sub txtVSMR_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub
