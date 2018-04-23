VERSION 5.00
Object = "{20D5284F-7B23-4F0A-B8B1-6C9D18B64F1C}#1.0#0"; "exlimiter.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMonster 
   Caption         =   "Monster Editor"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8655
   Icon            =   "frmMonster.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   8655
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
      Left            =   120
      TabIndex        =   447
      Top             =   0
      Width           =   2775
   End
   Begin VB.Frame fraFilter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   5475
      Left            =   60
      TabIndex        =   420
      Top             =   360
      Visible         =   0   'False
      Width           =   4935
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
         Height          =   5115
         Left            =   180
         TabIndex        =   421
         Top             =   180
         Width           =   4575
         Begin VB.CheckBox chkFilter 
            Caption         =   "Undead"
            Enabled         =   0   'False
            Height          =   255
            Index           =   13
            Left            =   3300
            TabIndex        =   464
            Top             =   1500
            Width           =   1095
         End
         Begin VB.CheckBox chkFilterExcludeZero 
            Caption         =   "Exclude 0 value on <= search"
            Height          =   255
            Left            =   1920
            TabIndex        =   463
            Top             =   360
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.ComboBox cmbFilterAbilityGL 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   6
            ItemData        =   "frmMonster.frx":08CA
            Left            =   1380
            List            =   "frmMonster.frx":08D7
            Style           =   2  'Dropdown List
            TabIndex        =   462
            Top             =   3300
            Width           =   795
         End
         Begin VB.TextBox txtFilterAbilityValue 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   6
            Left            =   2280
            TabIndex        =   461
            Text            =   "0"
            Top             =   3300
            Width           =   1455
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Exp Multi"
            Enabled         =   0   'False
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   460
            Top             =   3300
            Width           =   1095
         End
         Begin VB.ComboBox cmbFilterAbilityGL 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   5
            ItemData        =   "frmMonster.frx":08E6
            Left            =   1380
            List            =   "frmMonster.frx":08F3
            Style           =   2  'Dropdown List
            TabIndex        =   459
            Top             =   2940
            Width           =   795
         End
         Begin VB.TextBox txtFilterAbilityValue 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   5
            Left            =   2280
            TabIndex        =   458
            Text            =   "0"
            Top             =   2940
            Width           =   1455
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Exp"
            Enabled         =   0   'False
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   457
            Top             =   2940
            Width           =   1095
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Spell"
            Enabled         =   0   'False
            Height          =   255
            Index           =   10
            Left            =   2640
            TabIndex        =   456
            Top             =   4080
            Width           =   795
         End
         Begin VB.TextBox txtFilterSpell 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3480
            MaxLength       =   29
            TabIndex        =   455
            Text            =   "0"
            Top             =   4080
            Width           =   855
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Item"
            Enabled         =   0   'False
            Height          =   255
            Index           =   9
            Left            =   2640
            TabIndex        =   454
            Top             =   3720
            Width           =   795
         End
         Begin VB.TextBox txtFilterItem 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3480
            MaxLength       =   29
            TabIndex        =   453
            Text            =   "0"
            Top             =   3720
            Width           =   855
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Alignment"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   452
            Top             =   1500
            Width           =   1095
         End
         Begin VB.ComboBox cmbFilter 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   1
            ItemData        =   "frmMonster.frx":0902
            Left            =   1380
            List            =   "frmMonster.frx":091B
            Style           =   2  'Dropdown List
            TabIndex        =   451
            Top             =   1500
            Width           =   1575
         End
         Begin VB.TextBox txtFilterIndex 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   2280
            MaxLength       =   29
            TabIndex        =   449
            Text            =   "0"
            Top             =   1140
            Width           =   675
         End
         Begin VB.TextBox txtFilterIndex 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   1380
            MaxLength       =   29
            TabIndex        =   448
            Text            =   "0"
            Top             =   1140
            Width           =   675
         End
         Begin VB.TextBox txtFilterTB 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1380
            MaxLength       =   29
            TabIndex        =   446
            Text            =   "0"
            Top             =   4080
            Width           =   855
         End
         Begin VB.TextBox txtFilterMessage 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1380
            MaxLength       =   29
            TabIndex        =   445
            Text            =   "0"
            Top             =   3720
            Width           =   855
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Textblock"
            Enabled         =   0   'False
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   444
            Top             =   4080
            Width           =   1095
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Message"
            Enabled         =   0   'False
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   443
            Top             =   3720
            Width           =   1095
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Limit > 0"
            Enabled         =   0   'False
            Height          =   255
            Index           =   12
            Left            =   3300
            TabIndex        =   442
            Top             =   1140
            Width           =   1095
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Regen > 0"
            Enabled         =   0   'False
            Height          =   255
            Index           =   11
            Left            =   3300
            TabIndex        =   441
            Top             =   780
            Width           =   1095
         End
         Begin VB.CheckBox chkFilterIndex 
            Caption         =   "Index"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   440
            Top             =   1140
            Width           =   1095
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Ability"
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   439
            Top             =   1860
            Width           =   1095
         End
         Begin VB.ComboBox cmbFilter 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   2
            ItemData        =   "frmMonster.frx":096A
            Left            =   1380
            List            =   "frmMonster.frx":096C
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   438
            Top             =   1860
            Width           =   1575
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Group"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   437
            Top             =   780
            Width           =   1095
         End
         Begin VB.ComboBox cmbFilter 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   0
            ItemData        =   "frmMonster.frx":096E
            Left            =   1380
            List            =   "frmMonster.frx":09EA
            Style           =   2  'Dropdown List
            TabIndex        =   436
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
            TabIndex        =   435
            Top             =   360
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CommandButton cmdFilterCancel 
            Caption         =   "Cancel"
            Height          =   435
            Left            =   3000
            TabIndex        =   434
            Top             =   4500
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
            Height          =   435
            Left            =   240
            TabIndex        =   433
            Top             =   4500
            Width           =   1335
         End
         Begin VB.CommandButton cmdFilterReset 
            Caption         =   "Reset"
            Height          =   435
            Left            =   1740
            TabIndex        =   432
            Top             =   4500
            Width           =   1095
         End
         Begin VB.TextBox txtFilterAbilityValue 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   2
            Left            =   3780
            TabIndex        =   431
            Text            =   "0"
            Top             =   1860
            Width           =   555
         End
         Begin VB.ComboBox cmbFilterAbilityGL 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   2
            ItemData        =   "frmMonster.frx":0B59
            Left            =   3000
            List            =   "frmMonster.frx":0B69
            Style           =   2  'Dropdown List
            TabIndex        =   430
            Top             =   1860
            Width           =   735
         End
         Begin VB.ComboBox cmbFilterAbilityGL 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   3
            ItemData        =   "frmMonster.frx":0B7D
            Left            =   3000
            List            =   "frmMonster.frx":0B8D
            Style           =   2  'Dropdown List
            TabIndex        =   429
            Top             =   2220
            Width           =   735
         End
         Begin VB.TextBox txtFilterAbilityValue 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   3
            Left            =   3780
            TabIndex        =   428
            Text            =   "0"
            Top             =   2220
            Width           =   555
         End
         Begin VB.ComboBox cmbFilter 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   3
            ItemData        =   "frmMonster.frx":0BA1
            Left            =   1380
            List            =   "frmMonster.frx":0BA3
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   427
            Top             =   2220
            Width           =   1575
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Ability"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   426
            Top             =   2220
            Width           =   1095
         End
         Begin VB.ComboBox cmbFilterAbilityGL 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   4
            ItemData        =   "frmMonster.frx":0BA5
            Left            =   3000
            List            =   "frmMonster.frx":0BB5
            Style           =   2  'Dropdown List
            TabIndex        =   425
            Top             =   2580
            Width           =   735
         End
         Begin VB.TextBox txtFilterAbilityValue 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   4
            Left            =   3780
            TabIndex        =   424
            Text            =   "0"
            Top             =   2580
            Width           =   555
         End
         Begin VB.ComboBox cmbFilter 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   4
            ItemData        =   "frmMonster.frx":0BC9
            Left            =   1380
            List            =   "frmMonster.frx":0BCB
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   423
            Top             =   2580
            Width           =   1575
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Ability"
            Enabled         =   0   'False
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   422
            Top             =   2580
            Width           =   1095
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "to"
            Height          =   195
            Left            =   2100
            TabIndex        =   450
            Top             =   1200
            Width           =   135
         End
      End
   End
   Begin VB.Frame framNav 
      BorderStyle     =   0  'None
      Height          =   6675
      Left            =   3000
      TabIndex        =   5
      Top             =   0
      Width           =   5595
      Begin VB.CheckBox chkAutoSave 
         Caption         =   "Auto-Save"
         Height          =   195
         Left            =   2280
         TabIndex        =   395
         Top             =   60
         Value           =   1  'Checked
         Width           =   1155
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   285
         Left            =   900
         TabIndex        =   7
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdDiscard 
         Caption         =   "Dis&card"
         Height          =   285
         Left            =   4500
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "&Insert"
         Height          =   285
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   915
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   285
         Left            =   3480
         TabIndex        =   8
         Top             =   0
         Width           =   1035
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6375
         Left            =   0
         TabIndex        =   10
         Top             =   300
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   11245
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "General "
         TabPicture(0)   =   "frmMonster.frx":0BCD
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "label(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "label(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Line1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "label(54)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "label(15)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "label(14)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "label(13)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "label(12)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "label(11)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "label(9)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "label(8)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "label(7)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "label(6)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "label(5)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "label(4)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "label(2)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "label(16)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "label(74)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Label14"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Label15"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Label16"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "lblBase"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "lblMulti"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "label(3)"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "label(19)"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "label(24)"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "label(30)"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "txtNumber"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "txtName"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "txtGameLimit"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "chkUndead"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "txtEnergy"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "txtRegenTime"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "txtFollow"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "txtDR"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "txtAC"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "txtCharmlvl"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).Control(37)=   "txtMR"
         Tab(0).Control(37).Enabled=   0   'False
         Tab(0).Control(38)=   "txtHpRegen"
         Tab(0).Control(38).Enabled=   0   'False
         Tab(0).Control(39)=   "txtHitPoints"
         Tab(0).Control(39).Enabled=   0   'False
         Tab(0).Control(40)=   "txtExperience"
         Tab(0).Control(40).Enabled=   0   'False
         Tab(0).Control(41)=   "txtIndex"
         Tab(0).Control(41).Enabled=   0   'False
         Tab(0).Control(42)=   "txtGender"
         Tab(0).Control(42).Enabled=   0   'False
         Tab(0).Control(43)=   "cmbType"
         Tab(0).Control(43).Enabled=   0   'False
         Tab(0).Control(44)=   "txtAlignment"
         Tab(0).Control(44).Enabled=   0   'False
         Tab(0).Control(45)=   "cmbGroup"
         Tab(0).Control(45).Enabled=   0   'False
         Tab(0).Control(46)=   "Frame2"
         Tab(0).Control(46).Enabled=   0   'False
         Tab(0).Control(47)=   "txtActive"
         Tab(0).Control(47).Enabled=   0   'False
         Tab(0).Control(48)=   "txtDateKilled"
         Tab(0).Control(48).Enabled=   0   'False
         Tab(0).Control(49)=   "txtTimeKilled"
         Tab(0).Control(49).Enabled=   0   'False
         Tab(0).Control(50)=   "txtMulti"
         Tab(0).Control(50).Enabled=   0   'False
         Tab(0).Control(51)=   "txtBase"
         Tab(0).Control(51).Enabled=   0   'False
         Tab(0).Control(52)=   "txtBSDefense"
         Tab(0).Control(52).Enabled=   0   'False
         Tab(0).Control(53)=   "txtCharmRes"
         Tab(0).Control(53).Enabled=   0   'False
         Tab(0).Control(54)=   "cmdResetKill"
         Tab(0).Control(54).Enabled=   0   'False
         Tab(0).ControlCount=   55
         TabCaption(1)   =   " Drop "
         TabPicture(1)   =   "frmMonster.frx":0BE9
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   " Spells/Attacks "
         TabPicture(2)   =   "frmMonster.frx":0C05
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtDeathSpellName"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "txtDeathSpellNumber"
         Tab(2).Control(2)=   "txtCreateSpellNumber"
         Tab(2).Control(3)=   "txtCreateSpellName"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "cmdEditCreateSpell"
         Tab(2).Control(5)=   "cmdEditDeathSpell"
         Tab(2).Control(6)=   "Frame3"
         Tab(2).Control(7)=   "SSTab2"
         Tab(2).Control(8)=   "label(41)"
         Tab(2).Control(9)=   "label(10)"
         Tab(2).ControlCount=   10
         TabCaption(3)   =   " Weapon/Txt/Msg "
         TabPicture(3)   =   "frmMonster.frx":0C21
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtDeathMsgDisplay"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "txtDeathMsg"
         Tab(3).Control(2)=   "txtMoveMsgDisplay"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).Control(3)=   "txtMoveMsg"
         Tab(3).Control(4)=   "txtTalkTxt"
         Tab(3).Control(5)=   "txtDescTxt"
         Tab(3).Control(6)=   "txtGreetTxt"
         Tab(3).Control(7)=   "txtGreetTxtDisplay"
         Tab(3).Control(7).Enabled=   0   'False
         Tab(3).Control(8)=   "txtTalkTxtDisplay"
         Tab(3).Control(8).Enabled=   0   'False
         Tab(3).Control(9)=   "txtDescTxtDisplay"
         Tab(3).Control(9).Enabled=   0   'False
         Tab(3).Control(10)=   "cmdEditMoveMsg"
         Tab(3).Control(11)=   "cmdEditDeathMsg"
         Tab(3).Control(12)=   "cmdEditGreetTxt"
         Tab(3).Control(13)=   "cmdEditTalkText"
         Tab(3).Control(14)=   "cmdEditDescText"
         Tab(3).Control(15)=   "txtWeaponNumber"
         Tab(3).Control(16)=   "txtWeaponName"
         Tab(3).Control(16).Enabled=   0   'False
         Tab(3).Control(17)=   "cmdEditWeapon"
         Tab(3).Control(18)=   "Label13(1)"
         Tab(3).Control(19)=   "Label13(0)"
         Tab(3).Control(20)=   "label(69)"
         Tab(3).Control(21)=   "label(70)"
         Tab(3).Control(22)=   "label(71)"
         Tab(3).Control(23)=   "label(72)"
         Tab(3).Control(24)=   "label(73)"
         Tab(3).Control(25)=   "label(17)"
         Tab(3).ControlCount=   26
         TabCaption(4)   =   " Abilities "
         TabPicture(4)   =   "frmMonster.frx":0C3D
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "cmdAbilsClear"
         Tab(4).Control(1)=   "frmAbilities"
         Tab(4).Control(2)=   "Label8"
         Tab(4).ControlCount=   3
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
            Left            =   -74760
            TabIndex        =   349
            Top             =   600
            Width           =   675
         End
         Begin VB.Frame frmAbilities 
            Caption         =   "Abilities"
            Height          =   4095
            Left            =   -73920
            TabIndex        =   348
            Top             =   480
            Width           =   3615
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   4
               Left            =   180
               TabIndex        =   366
               Top             =   1860
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   3
               Left            =   180
               TabIndex        =   362
               Top             =   1500
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   2
               Left            =   180
               TabIndex        =   358
               Top             =   1140
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   1
               Left            =   180
               TabIndex        =   354
               Top             =   780
               Width           =   495
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   0
               Left            =   2700
               TabIndex        =   352
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   420
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   1
               Left            =   2700
               TabIndex        =   356
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   780
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   2
               Left            =   2700
               TabIndex        =   360
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   1140
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   3
               Left            =   2700
               TabIndex        =   364
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   1500
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   4
               Left            =   2700
               TabIndex        =   368
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   1860
               Width           =   615
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   0
               Left            =   180
               TabIndex        =   350
               Top             =   420
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
               Left            =   780
               TabIndex        =   351
               Text            =   "empty"
               Top             =   420
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
               Left            =   780
               TabIndex        =   355
               Text            =   "empty"
               Top             =   780
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
               Left            =   780
               TabIndex        =   359
               Text            =   "empty"
               Top             =   1140
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
               Left            =   780
               TabIndex        =   363
               Text            =   "empty"
               Top             =   1500
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
               Left            =   780
               TabIndex        =   367
               Text            =   "empty"
               Top             =   1860
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
               Left            =   780
               TabIndex        =   371
               Text            =   "empty"
               Top             =   2220
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
               Left            =   780
               TabIndex        =   375
               Text            =   "empty"
               Top             =   2580
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
               Left            =   780
               TabIndex        =   379
               Text            =   "empty"
               Top             =   2940
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
               Left            =   780
               TabIndex        =   383
               Text            =   "empty"
               Top             =   3300
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
               Left            =   780
               TabIndex        =   387
               Text            =   "empty"
               Top             =   3660
               Width           =   1815
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   5
               Left            =   180
               TabIndex        =   370
               Top             =   2220
               Width           =   495
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   5
               Left            =   2700
               TabIndex        =   372
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   2220
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   6
               Left            =   2700
               TabIndex        =   376
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   2580
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   7
               Left            =   2700
               TabIndex        =   380
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   2940
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   8
               Left            =   2700
               TabIndex        =   384
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   3300
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   9
               Left            =   2700
               TabIndex        =   388
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   3660
               Width           =   615
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   6
               Left            =   180
               TabIndex        =   374
               Top             =   2580
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   7
               Left            =   180
               TabIndex        =   378
               Top             =   2940
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   8
               Left            =   180
               TabIndex        =   382
               Top             =   3300
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   9
               Left            =   180
               TabIndex        =   386
               Top             =   3660
               Width           =   495
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   9
               Left            =   3360
               TabIndex        =   389
               Top             =   3660
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   8
               Left            =   3360
               TabIndex        =   385
               Top             =   3300
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   7
               Left            =   3360
               TabIndex        =   381
               Top             =   2940
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   6
               Left            =   3360
               TabIndex        =   377
               Top             =   2580
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   5
               Left            =   3360
               TabIndex        =   373
               Top             =   2220
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   4
               Left            =   3360
               TabIndex        =   369
               Top             =   1860
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   3
               Left            =   3360
               TabIndex        =   365
               Top             =   1500
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   2
               Left            =   3360
               TabIndex        =   361
               Top             =   1140
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   1
               Left            =   3360
               TabIndex        =   357
               Top             =   780
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   0
               Left            =   3360
               TabIndex        =   353
               Top             =   420
               Width           =   135
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Caption         =   "#"
               Height          =   255
               Index           =   1
               Left            =   180
               TabIndex        =   392
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Name"
               Height          =   255
               Index           =   1
               Left            =   780
               TabIndex        =   391
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Value"
               Height          =   255
               Index           =   1
               Left            =   2700
               TabIndex        =   390
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.TextBox txtDeathMsgDisplay 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   -72720
            Locked          =   -1  'True
            TabIndex        =   334
            TabStop         =   0   'False
            Top             =   1380
            Width           =   3015
         End
         Begin VB.TextBox txtDeathMsg 
            Height          =   285
            Left            =   -73380
            TabIndex        =   333
            Top             =   1380
            Width           =   615
         End
         Begin VB.TextBox txtMoveMsgDisplay 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   -72720
            Locked          =   -1  'True
            TabIndex        =   330
            TabStop         =   0   'False
            Top             =   1020
            Width           =   3015
         End
         Begin VB.TextBox txtMoveMsg 
            Height          =   285
            Left            =   -73380
            TabIndex        =   329
            Top             =   1020
            Width           =   615
         End
         Begin VB.TextBox txtTalkTxt 
            Height          =   285
            Left            =   -73380
            TabIndex        =   341
            Top             =   2220
            Width           =   615
         End
         Begin VB.TextBox txtDescTxt 
            Height          =   285
            Left            =   -73380
            TabIndex        =   345
            Top             =   2580
            Width           =   615
         End
         Begin VB.TextBox txtGreetTxt 
            Height          =   285
            Left            =   -73380
            TabIndex        =   337
            Top             =   1860
            Width           =   615
         End
         Begin VB.TextBox txtGreetTxtDisplay 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   -72720
            Locked          =   -1  'True
            TabIndex        =   338
            TabStop         =   0   'False
            Top             =   1860
            Width           =   3015
         End
         Begin VB.TextBox txtTalkTxtDisplay 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   -72720
            Locked          =   -1  'True
            TabIndex        =   342
            TabStop         =   0   'False
            Top             =   2220
            Width           =   3015
         End
         Begin VB.TextBox txtDescTxtDisplay 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   -72720
            Locked          =   -1  'True
            TabIndex        =   346
            TabStop         =   0   'False
            Top             =   2580
            Width           =   3015
         End
         Begin VB.CommandButton cmdEditMoveMsg 
            Height          =   195
            Left            =   -74580
            TabIndex        =   327
            Top             =   1020
            Width           =   195
         End
         Begin VB.CommandButton cmdEditDeathMsg 
            Height          =   195
            Left            =   -74580
            TabIndex        =   331
            Top             =   1380
            Width           =   195
         End
         Begin VB.CommandButton cmdEditGreetTxt 
            Height          =   195
            Left            =   -74580
            TabIndex        =   335
            Top             =   1860
            Width           =   195
         End
         Begin VB.CommandButton cmdEditTalkText 
            Height          =   195
            Left            =   -74580
            TabIndex        =   339
            Top             =   2220
            Width           =   195
         End
         Begin VB.CommandButton cmdEditDescText 
            Height          =   195
            Left            =   -74580
            TabIndex        =   343
            Top             =   2580
            Width           =   195
         End
         Begin VB.TextBox txtDeathSpellName 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   -72360
            Locked          =   -1  'True
            MaxLength       =   28
            TabIndex        =   115
            TabStop         =   0   'False
            Top             =   720
            Width           =   2415
         End
         Begin VB.TextBox txtDeathSpellNumber 
            Height          =   285
            Left            =   -73080
            TabIndex        =   114
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtCreateSpellNumber 
            Height          =   285
            Left            =   -73080
            TabIndex        =   110
            Top             =   420
            Width           =   615
         End
         Begin VB.TextBox txtCreateSpellName 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   -72360
            Locked          =   -1  'True
            MaxLength       =   28
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   420
            Width           =   2415
         End
         Begin VB.CommandButton cmdEditCreateSpell 
            Height          =   195
            Left            =   -74340
            TabIndex        =   108
            Top             =   420
            Width           =   195
         End
         Begin VB.CommandButton cmdEditDeathSpell 
            Height          =   195
            Left            =   -74340
            TabIndex        =   112
            Top             =   720
            Width           =   195
         End
         Begin VB.TextBox txtWeaponNumber 
            Height          =   285
            Left            =   -73380
            TabIndex        =   325
            Top             =   540
            Width           =   615
         End
         Begin VB.TextBox txtWeaponName 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   -72720
            Locked          =   -1  'True
            MaxLength       =   28
            TabIndex        =   326
            TabStop         =   0   'False
            Top             =   540
            Width           =   3015
         End
         Begin VB.CommandButton cmdEditWeapon 
            Height          =   195
            Left            =   -74580
            TabIndex        =   323
            Top             =   540
            Width           =   195
         End
         Begin VB.CommandButton cmdResetKill 
            Caption         =   "< Reset"
            Height          =   255
            Left            =   3480
            TabIndex        =   26
            Top             =   4455
            Width           =   795
         End
         Begin VB.TextBox txtCharmRes 
            Height          =   285
            Left            =   4620
            TabIndex        =   34
            Top             =   2880
            Width           =   735
         End
         Begin VB.TextBox txtBSDefense 
            Height          =   285
            Left            =   4620
            TabIndex        =   35
            Top             =   3180
            Width           =   735
         End
         Begin VB.Frame Frame3 
            Caption         =   "Betwen Round Spells:"
            Height          =   1755
            Left            =   -74580
            TabIndex        =   116
            Top             =   1080
            Width           =   4755
            Begin VB.TextBox txtSpellCastLvL 
               Height          =   285
               Index           =   4
               Left            =   3960
               TabIndex        =   145
               Top             =   1380
               Width           =   615
            End
            Begin VB.TextBox txtSpellCastPer 
               Height          =   285
               Index           =   4
               Left            =   3360
               TabIndex        =   144
               Top             =   1380
               Width           =   615
            End
            Begin VB.TextBox txtSpellName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   4
               Left            =   1080
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   143
               TabStop         =   0   'False
               Top             =   1380
               Width           =   2295
            End
            Begin VB.TextBox txtSpellNumber 
               Height          =   285
               Index           =   4
               Left            =   480
               TabIndex        =   142
               Top             =   1380
               Width           =   615
            End
            Begin VB.TextBox txtSpellCastLvL 
               Height          =   285
               Index           =   3
               Left            =   3960
               TabIndex        =   140
               Top             =   1140
               Width           =   615
            End
            Begin VB.TextBox txtSpellCastPer 
               Height          =   285
               Index           =   3
               Left            =   3360
               TabIndex        =   139
               Top             =   1140
               Width           =   615
            End
            Begin VB.TextBox txtSpellName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   3
               Left            =   1080
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   138
               TabStop         =   0   'False
               Top             =   1140
               Width           =   2295
            End
            Begin VB.TextBox txtSpellNumber 
               Height          =   285
               Index           =   3
               Left            =   480
               TabIndex        =   137
               Top             =   1140
               Width           =   615
            End
            Begin VB.TextBox txtSpellCastLvL 
               Height          =   285
               Index           =   2
               Left            =   3960
               TabIndex        =   135
               Top             =   900
               Width           =   615
            End
            Begin VB.TextBox txtSpellCastPer 
               Height          =   285
               Index           =   2
               Left            =   3360
               TabIndex        =   134
               Top             =   900
               Width           =   615
            End
            Begin VB.TextBox txtSpellName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   2
               Left            =   1080
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   133
               TabStop         =   0   'False
               Top             =   900
               Width           =   2295
            End
            Begin VB.TextBox txtSpellNumber 
               Height          =   285
               Index           =   2
               Left            =   480
               TabIndex        =   132
               Top             =   900
               Width           =   615
            End
            Begin VB.TextBox txtSpellCastLvL 
               Height          =   285
               Index           =   1
               Left            =   3960
               TabIndex        =   130
               Top             =   660
               Width           =   615
            End
            Begin VB.TextBox txtSpellCastPer 
               Height          =   285
               Index           =   1
               Left            =   3360
               TabIndex        =   129
               Top             =   660
               Width           =   615
            End
            Begin VB.TextBox txtSpellName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   1
               Left            =   1080
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   128
               TabStop         =   0   'False
               Top             =   660
               Width           =   2295
            End
            Begin VB.TextBox txtSpellNumber 
               Height          =   285
               Index           =   1
               Left            =   480
               TabIndex        =   127
               Top             =   660
               Width           =   615
            End
            Begin VB.TextBox txtSpellCastLvL 
               Height          =   285
               Index           =   0
               Left            =   3960
               TabIndex        =   125
               Top             =   420
               Width           =   615
            End
            Begin VB.TextBox txtSpellCastPer 
               Height          =   285
               Index           =   0
               Left            =   3360
               TabIndex        =   124
               Top             =   420
               Width           =   615
            End
            Begin VB.TextBox txtSpellName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   0
               Left            =   1080
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   123
               TabStop         =   0   'False
               Top             =   420
               Width           =   2295
            End
            Begin VB.TextBox txtSpellNumber 
               Height          =   285
               Index           =   0
               Left            =   480
               TabIndex        =   122
               Top             =   420
               Width           =   615
            End
            Begin VB.CommandButton cmdEditSpell 
               Height          =   195
               Index           =   0
               Left            =   180
               TabIndex        =   121
               Top             =   465
               Width           =   195
            End
            Begin VB.CommandButton cmdEditSpell 
               Height          =   195
               Index           =   1
               Left            =   180
               TabIndex        =   126
               Top             =   705
               Width           =   195
            End
            Begin VB.CommandButton cmdEditSpell 
               Height          =   195
               Index           =   2
               Left            =   180
               TabIndex        =   131
               Top             =   945
               Width           =   195
            End
            Begin VB.CommandButton cmdEditSpell 
               Height          =   195
               Index           =   3
               Left            =   180
               TabIndex        =   136
               Top             =   1185
               Width           =   195
            End
            Begin VB.CommandButton cmdEditSpell 
               Height          =   195
               Index           =   4
               Left            =   180
               TabIndex        =   141
               Top             =   1425
               Width           =   195
            End
            Begin VB.Label label 
               Alignment       =   2  'Center
               Caption         =   "#"
               Height          =   255
               Index           =   34
               Left            =   480
               TabIndex        =   117
               Top             =   240
               Width           =   615
            End
            Begin VB.Label label 
               Alignment       =   2  'Center
               Caption         =   "Name"
               Height          =   255
               Index           =   33
               Left            =   1080
               TabIndex        =   118
               Top             =   240
               Width           =   2295
            End
            Begin VB.Label label 
               Alignment       =   2  'Center
               Caption         =   "%"
               Height          =   255
               Index           =   32
               Left            =   3360
               TabIndex        =   119
               Top             =   240
               Width           =   615
            End
            Begin VB.Label label 
               Alignment       =   2  'Center
               Caption         =   "LvL"
               Height          =   255
               Index           =   31
               Left            =   3960
               TabIndex        =   120
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.TextBox txtBase 
            Height          =   285
            Left            =   1140
            TabIndex        =   13
            Top             =   1020
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtMulti 
            Height          =   285
            Left            =   2160
            TabIndex        =   14
            Top             =   1020
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.TextBox txtTimeKilled 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "M/dd/yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Left            =   1140
            TabIndex        =   24
            Top             =   4440
            Width           =   1095
         End
         Begin VB.TextBox txtDateKilled 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "M/dd/yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Left            =   2280
            TabIndex        =   25
            Top             =   4440
            Width           =   1095
         End
         Begin VB.TextBox txtActive 
            Height          =   285
            Left            =   1140
            TabIndex        =   19
            Top             =   2580
            Width           =   735
         End
         Begin VB.Frame Frame1 
            Caption         =   "Items Dropped"
            Height          =   4935
            Left            =   -74760
            TabIndex        =   43
            Top             =   540
            Width           =   5115
            Begin VB.CommandButton cmdEditItemDrop 
               Height          =   195
               Index           =   9
               Left            =   300
               TabIndex        =   93
               Top             =   3120
               Width           =   195
            End
            Begin VB.CommandButton cmdEditItemDrop 
               Height          =   195
               Index           =   8
               Left            =   300
               TabIndex        =   88
               Top             =   2835
               Width           =   195
            End
            Begin VB.CommandButton cmdEditItemDrop 
               Height          =   195
               Index           =   7
               Left            =   300
               TabIndex        =   83
               Top             =   2550
               Width           =   195
            End
            Begin VB.CommandButton cmdEditItemDrop 
               Height          =   195
               Index           =   6
               Left            =   300
               TabIndex        =   78
               Top             =   2265
               Width           =   195
            End
            Begin VB.CommandButton cmdEditItemDrop 
               Height          =   195
               Index           =   5
               Left            =   300
               TabIndex        =   73
               Top             =   1980
               Width           =   195
            End
            Begin VB.CommandButton cmdEditItemDrop 
               Height          =   195
               Index           =   4
               Left            =   300
               TabIndex        =   68
               Top             =   1680
               Width           =   195
            End
            Begin VB.CommandButton cmdEditItemDrop 
               Height          =   195
               Index           =   3
               Left            =   300
               TabIndex        =   63
               Top             =   1395
               Width           =   195
            End
            Begin VB.CommandButton cmdEditItemDrop 
               Height          =   195
               Index           =   2
               Left            =   300
               TabIndex        =   58
               Top             =   1110
               Width           =   195
            End
            Begin VB.CommandButton cmdEditItemDrop 
               Height          =   195
               Index           =   1
               Left            =   300
               TabIndex        =   53
               Top             =   825
               Width           =   195
            End
            Begin VB.CommandButton cmdEditItemDrop 
               Height          =   195
               Index           =   0
               Left            =   300
               TabIndex        =   48
               Top             =   540
               Width           =   195
            End
            Begin VB.TextBox txtItemNumber 
               Height          =   285
               Index           =   5
               Left            =   540
               TabIndex        =   74
               Top             =   1980
               Width           =   615
            End
            Begin VB.TextBox txtItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   6
               Left            =   1155
               Locked          =   -1  'True
               TabIndex        =   80
               TabStop         =   0   'False
               Top             =   2265
               Width           =   2295
            End
            Begin VB.TextBox txtItemNumber 
               Height          =   285
               Index           =   6
               Left            =   540
               TabIndex        =   79
               Top             =   2265
               Width           =   615
            End
            Begin VB.TextBox txtItemUses 
               Height          =   285
               Index           =   6
               Left            =   4065
               TabIndex        =   82
               Top             =   2265
               Width           =   615
            End
            Begin VB.TextBox txtItemDropPer 
               Height          =   285
               Index           =   6
               Left            =   3450
               TabIndex        =   81
               Top             =   2265
               Width           =   615
            End
            Begin VB.TextBox txtItemNumber 
               Height          =   285
               Index           =   1
               Left            =   540
               TabIndex        =   54
               Top             =   825
               Width           =   615
            End
            Begin VB.TextBox txtItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   1
               Left            =   1155
               Locked          =   -1  'True
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   825
               Width           =   2295
            End
            Begin VB.TextBox txtItemDropPer 
               Height          =   285
               Index           =   1
               Left            =   3450
               TabIndex        =   56
               Top             =   825
               Width           =   615
            End
            Begin VB.TextBox txtItemUses 
               Height          =   285
               Index           =   1
               Left            =   4065
               TabIndex        =   57
               Top             =   825
               Width           =   615
            End
            Begin VB.TextBox txtItemNumber 
               Height          =   285
               Index           =   0
               Left            =   540
               TabIndex        =   49
               Top             =   540
               Width           =   615
            End
            Begin VB.TextBox txtItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   0
               Left            =   1155
               Locked          =   -1  'True
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   540
               Width           =   2295
            End
            Begin VB.TextBox txtItemDropPer 
               Height          =   285
               Index           =   0
               Left            =   3450
               TabIndex        =   51
               Top             =   540
               Width           =   615
            End
            Begin VB.TextBox txtItemUses 
               Height          =   285
               Index           =   0
               Left            =   4065
               TabIndex        =   52
               Top             =   540
               Width           =   615
            End
            Begin VB.TextBox txtItemNumber 
               Height          =   285
               Index           =   2
               Left            =   540
               TabIndex        =   59
               Top             =   1110
               Width           =   615
            End
            Begin VB.TextBox txtItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   2
               Left            =   1155
               Locked          =   -1  'True
               TabIndex        =   60
               TabStop         =   0   'False
               Top             =   1110
               Width           =   2295
            End
            Begin VB.TextBox txtItemNumber 
               Height          =   285
               Index           =   3
               Left            =   540
               TabIndex        =   64
               Top             =   1395
               Width           =   615
            End
            Begin VB.TextBox txtItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   3
               Left            =   1155
               Locked          =   -1  'True
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   1395
               Width           =   2295
            End
            Begin VB.TextBox txtItemNumber 
               Height          =   285
               Index           =   4
               Left            =   540
               TabIndex        =   69
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   4
               Left            =   1155
               Locked          =   -1  'True
               TabIndex        =   70
               TabStop         =   0   'False
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox txtItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   5
               Left            =   1155
               Locked          =   -1  'True
               TabIndex        =   75
               TabStop         =   0   'False
               Top             =   1980
               Width           =   2295
            End
            Begin VB.TextBox txtItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   7
               Left            =   1155
               Locked          =   -1  'True
               TabIndex        =   85
               TabStop         =   0   'False
               Top             =   2550
               Width           =   2295
            End
            Begin VB.TextBox txtItemNumber 
               Height          =   285
               Index           =   7
               Left            =   540
               TabIndex        =   84
               Top             =   2550
               Width           =   615
            End
            Begin VB.TextBox txtItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   8
               Left            =   1155
               Locked          =   -1  'True
               TabIndex        =   90
               TabStop         =   0   'False
               Top             =   2835
               Width           =   2295
            End
            Begin VB.TextBox txtItemNumber 
               Height          =   285
               Index           =   8
               Left            =   540
               TabIndex        =   89
               Top             =   2835
               Width           =   615
            End
            Begin VB.TextBox txtItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   9
               Left            =   1155
               Locked          =   -1  'True
               TabIndex        =   95
               TabStop         =   0   'False
               Top             =   3120
               Width           =   2295
            End
            Begin VB.TextBox txtItemNumber 
               Height          =   285
               Index           =   9
               Left            =   540
               TabIndex        =   94
               Top             =   3120
               Width           =   615
            End
            Begin VB.TextBox txtItemDropPer 
               Height          =   285
               Index           =   9
               Left            =   3450
               TabIndex        =   96
               Top             =   3120
               Width           =   615
            End
            Begin VB.TextBox txtItemUses 
               Height          =   285
               Index           =   9
               Left            =   4065
               TabIndex        =   97
               Top             =   3120
               Width           =   615
            End
            Begin VB.TextBox txtItemDropPer 
               Height          =   285
               Index           =   8
               Left            =   3450
               TabIndex        =   91
               Top             =   2835
               Width           =   615
            End
            Begin VB.TextBox txtItemUses 
               Height          =   285
               Index           =   8
               Left            =   4065
               TabIndex        =   92
               Top             =   2835
               Width           =   615
            End
            Begin VB.TextBox txtItemDropPer 
               Height          =   285
               Index           =   7
               Left            =   3450
               TabIndex        =   86
               Top             =   2550
               Width           =   615
            End
            Begin VB.TextBox txtItemUses 
               Height          =   285
               Index           =   7
               Left            =   4065
               TabIndex        =   87
               Top             =   2550
               Width           =   615
            End
            Begin VB.TextBox txtItemDropPer 
               Height          =   285
               Index           =   5
               Left            =   3450
               TabIndex        =   76
               Top             =   1980
               Width           =   615
            End
            Begin VB.TextBox txtItemUses 
               Height          =   285
               Index           =   5
               Left            =   4065
               TabIndex        =   77
               Top             =   1980
               Width           =   615
            End
            Begin VB.TextBox txtCopper 
               Height          =   315
               Left            =   3420
               TabIndex        =   107
               Top             =   4020
               Width           =   615
            End
            Begin VB.TextBox txtSilver 
               Height          =   315
               Left            =   3420
               TabIndex        =   105
               Top             =   3660
               Width           =   615
            End
            Begin VB.TextBox txtGold 
               Height          =   315
               Left            =   1860
               TabIndex        =   103
               Top             =   4380
               Width           =   615
            End
            Begin VB.TextBox txtPlatinum 
               Height          =   315
               Left            =   1860
               TabIndex        =   101
               Top             =   4020
               Width           =   615
            End
            Begin VB.TextBox txtRunic 
               Height          =   315
               Left            =   1860
               TabIndex        =   99
               Top             =   3660
               Width           =   615
            End
            Begin VB.TextBox txtItemUses 
               Height          =   285
               Index           =   4
               Left            =   4065
               TabIndex        =   72
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtItemDropPer 
               Height          =   285
               Index           =   4
               Left            =   3450
               TabIndex        =   71
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtItemUses 
               Height          =   285
               Index           =   3
               Left            =   4065
               TabIndex        =   67
               Top             =   1395
               Width           =   615
            End
            Begin VB.TextBox txtItemDropPer 
               Height          =   285
               Index           =   3
               Left            =   3450
               TabIndex        =   66
               Top             =   1395
               Width           =   615
            End
            Begin VB.TextBox txtItemUses 
               Height          =   285
               Index           =   2
               Left            =   4065
               TabIndex        =   62
               Top             =   1110
               Width           =   615
            End
            Begin VB.TextBox txtItemDropPer 
               Height          =   285
               Index           =   2
               Left            =   3450
               TabIndex        =   61
               Top             =   1110
               Width           =   615
            End
            Begin VB.Label label 
               Caption         =   "Copper"
               Height          =   255
               Index           =   29
               Left            =   2700
               TabIndex        =   106
               Top             =   4020
               Width           =   735
            End
            Begin VB.Label label 
               Caption         =   "Silver"
               Height          =   255
               Index           =   28
               Left            =   2700
               TabIndex        =   104
               Top             =   3660
               Width           =   735
            End
            Begin VB.Label label 
               Caption         =   "Gold"
               Height          =   255
               Index           =   27
               Left            =   1140
               TabIndex        =   102
               Top             =   4380
               Width           =   735
            End
            Begin VB.Label label 
               Caption         =   "Platinum"
               Height          =   255
               Index           =   26
               Left            =   1140
               TabIndex        =   100
               Top             =   4020
               Width           =   735
            End
            Begin VB.Label label 
               Caption         =   "Runic"
               Height          =   255
               Index           =   25
               Left            =   1140
               TabIndex        =   98
               Top             =   3660
               Width           =   735
            End
            Begin VB.Label label 
               Alignment       =   2  'Center
               Caption         =   "Uses"
               Height          =   255
               Index           =   23
               Left            =   4020
               TabIndex        =   47
               Top             =   300
               Width           =   615
            End
            Begin VB.Label label 
               Alignment       =   2  'Center
               Caption         =   "%"
               Height          =   255
               Index           =   22
               Left            =   3420
               TabIndex        =   46
               Top             =   300
               Width           =   615
            End
            Begin VB.Label label 
               Alignment       =   2  'Center
               Caption         =   "Name"
               Height          =   255
               Index           =   21
               Left            =   1140
               TabIndex        =   45
               Top             =   300
               Width           =   2295
            End
            Begin VB.Label label 
               Alignment       =   2  'Center
               Caption         =   "#"
               Height          =   255
               Index           =   20
               Left            =   540
               TabIndex        =   44
               Top             =   300
               Width           =   615
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Description"
            Height          =   1515
            Left            =   120
            TabIndex        =   38
            Top             =   4740
            Width           =   5295
            Begin VB.TextBox txtDesc 
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   1
               Left            =   120
               MaxLength       =   70
               TabIndex        =   40
               Top             =   540
               Width           =   5055
            End
            Begin VB.TextBox txtDesc 
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   0
               Left            =   120
               MaxLength       =   70
               TabIndex        =   39
               Top             =   240
               Width           =   5055
            End
            Begin VB.TextBox txtDesc 
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   2
               Left            =   120
               MaxLength       =   70
               TabIndex        =   41
               Top             =   840
               Width           =   5055
            End
            Begin VB.TextBox txtDesc 
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   3
               Left            =   120
               MaxLength       =   70
               TabIndex        =   42
               Top             =   1140
               Width           =   5055
            End
         End
         Begin VB.ComboBox cmbGroup 
            Height          =   315
            ItemData        =   "frmMonster.frx":0C59
            Left            =   1140
            List            =   "frmMonster.frx":0CD5
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1920
            Width           =   1455
         End
         Begin VB.ComboBox txtAlignment 
            Height          =   315
            ItemData        =   "frmMonster.frx":0E44
            Left            =   1140
            List            =   "frmMonster.frx":0E5D
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   3540
            Width           =   1455
         End
         Begin VB.ComboBox cmbType 
            Height          =   315
            ItemData        =   "frmMonster.frx":0EAC
            Left            =   1140
            List            =   "frmMonster.frx":0EBC
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   3180
            Width           =   1455
         End
         Begin VB.ComboBox txtGender 
            Height          =   315
            ItemData        =   "frmMonster.frx":0EE4
            Left            =   1140
            List            =   "frmMonster.frx":0EF1
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   3900
            Width           =   1455
         End
         Begin VB.TextBox txtIndex 
            Height          =   285
            Left            =   1140
            TabIndex        =   16
            Top             =   1620
            Width           =   735
         End
         Begin VB.TextBox txtExperience 
            Height          =   285
            Left            =   1140
            TabIndex        =   15
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox txtHitPoints 
            Height          =   285
            Left            =   4620
            TabIndex        =   29
            Top             =   1380
            Width           =   735
         End
         Begin VB.TextBox txtHpRegen 
            Height          =   285
            Left            =   4620
            TabIndex        =   30
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtMR 
            Height          =   285
            Left            =   4620
            TabIndex        =   31
            Top             =   1980
            Width           =   735
         End
         Begin VB.TextBox txtCharmlvl 
            Height          =   285
            Left            =   4620
            TabIndex        =   33
            Top             =   2580
            Width           =   735
         End
         Begin VB.TextBox txtAC 
            Height          =   285
            Left            =   3840
            TabIndex        =   27
            Top             =   1020
            Width           =   735
         End
         Begin VB.TextBox txtDR 
            Height          =   285
            Left            =   4620
            TabIndex        =   28
            Top             =   1020
            Width           =   735
         End
         Begin VB.TextBox txtFollow 
            Height          =   285
            Left            =   4620
            TabIndex        =   32
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txtRegenTime 
            Height          =   285
            Left            =   1140
            TabIndex        =   20
            Top             =   2880
            Width           =   735
         End
         Begin VB.TextBox txtEnergy 
            Height          =   285
            Left            =   4620
            TabIndex        =   36
            Top             =   3480
            Width           =   735
         End
         Begin VB.CheckBox chkUndead 
            Alignment       =   1  'Right Justify
            Caption         =   "Undead"
            Height          =   255
            Left            =   3900
            TabIndex        =   37
            Top             =   3810
            Width           =   930
         End
         Begin VB.TextBox txtGameLimit 
            Height          =   285
            Left            =   1140
            TabIndex        =   18
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   1080
            MaxLength       =   29
            TabIndex        =   12
            Top             =   420
            Width           =   2535
         End
         Begin VB.TextBox txtNumber 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   4620
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   420
            Width           =   735
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   3375
            Left            =   -74880
            TabIndex        =   146
            Top             =   2880
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   5953
            _Version        =   393216
            Style           =   1
            Tabs            =   6
            TabsPerRow      =   6
            TabHeight       =   520
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Attack 1"
            TabPicture(0)   =   "frmMonster.frx":0F02
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lblAttackHitSpell(52)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "lblAttackDodgeMsg(51)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "lblAttackMissMsg(50)"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "lblAttackHitMsg(49)"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "lblAttackEnergy(48)"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "lblAttackPercent(47)"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "lblAttackMaxHCastLvl(0)"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "lblAttackMinHCastPer(0)"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "lblAttackAccuSpell(0)"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "lblAttackType(43)"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "lblAttackSpellRange(0)"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "txtAttackHitSpellName(0)"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "txtAttackMissMsgDisplay(0)"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "txtAttackDodgeMsgDisplay(0)"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "txtAttackHitMsgDisplay(0)"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "txtAttackHitSpell(0)"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "txtAttackDodgeMsg(0)"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "txtAttackMissMsg(0)"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).Control(18)=   "txtAttackHitMsg(0)"
            Tab(0).Control(18).Enabled=   0   'False
            Tab(0).Control(19)=   "txtAttackEnergy(0)"
            Tab(0).Control(19).Enabled=   0   'False
            Tab(0).Control(20)=   "txtAttackPer(0)"
            Tab(0).Control(20).Enabled=   0   'False
            Tab(0).Control(21)=   "txtAttackMaxHCastLvL(0)"
            Tab(0).Control(21).Enabled=   0   'False
            Tab(0).Control(22)=   "txtAttackMinHCastPer(0)"
            Tab(0).Control(22).Enabled=   0   'False
            Tab(0).Control(23)=   "txtAttackAccuSpellName(0)"
            Tab(0).Control(23).Enabled=   0   'False
            Tab(0).Control(24)=   "txtAttackAccuSpell(0)"
            Tab(0).Control(24).Enabled=   0   'False
            Tab(0).Control(25)=   "cmbAttackType(0)"
            Tab(0).Control(25).Enabled=   0   'False
            Tab(0).Control(26)=   "cmdEditAttackSpell(0)"
            Tab(0).Control(26).Enabled=   0   'False
            Tab(0).Control(27)=   "cmdEditHitMsg(0)"
            Tab(0).Control(27).Enabled=   0   'False
            Tab(0).Control(28)=   "cmdEditMissMsg(0)"
            Tab(0).Control(28).Enabled=   0   'False
            Tab(0).Control(29)=   "cmdEditDodgeMsg(0)"
            Tab(0).Control(29).Enabled=   0   'False
            Tab(0).Control(30)=   "cmdEditHitSpell(0)"
            Tab(0).Control(30).Enabled=   0   'False
            Tab(0).Control(31)=   "txtAttackSpellDamage(0)"
            Tab(0).Control(31).Enabled=   0   'False
            Tab(0).Control(32)=   "cmdAttackSim(0)"
            Tab(0).Control(32).Enabled=   0   'False
            Tab(0).ControlCount=   33
            TabCaption(1)   =   "Attack 2"
            TabPicture(1)   =   "frmMonster.frx":0F1E
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lblAttackType(0)"
            Tab(1).Control(1)=   "lblAttackAccuSpell(1)"
            Tab(1).Control(2)=   "lblAttackMinHCastPer(1)"
            Tab(1).Control(3)=   "lblAttackMaxHCastLvl(1)"
            Tab(1).Control(4)=   "lblAttackPercent(0)"
            Tab(1).Control(5)=   "lblAttackEnergy(0)"
            Tab(1).Control(6)=   "lblAttackHitMsg(0)"
            Tab(1).Control(7)=   "lblAttackMissMsg(0)"
            Tab(1).Control(8)=   "lblAttackDodgeMsg(0)"
            Tab(1).Control(9)=   "lblAttackHitSpell(0)"
            Tab(1).Control(10)=   "lblAttackSpellRange(1)"
            Tab(1).Control(11)=   "txtAttackAccuSpell(1)"
            Tab(1).Control(12)=   "txtAttackAccuSpellName(1)"
            Tab(1).Control(12).Enabled=   0   'False
            Tab(1).Control(13)=   "txtAttackMinHCastPer(1)"
            Tab(1).Control(14)=   "txtAttackMaxHCastLvL(1)"
            Tab(1).Control(15)=   "txtAttackPer(1)"
            Tab(1).Control(16)=   "txtAttackEnergy(1)"
            Tab(1).Control(17)=   "txtAttackHitMsg(1)"
            Tab(1).Control(18)=   "txtAttackMissMsg(1)"
            Tab(1).Control(19)=   "txtAttackDodgeMsg(1)"
            Tab(1).Control(20)=   "txtAttackHitSpell(1)"
            Tab(1).Control(21)=   "txtAttackHitMsgDisplay(1)"
            Tab(1).Control(21).Enabled=   0   'False
            Tab(1).Control(22)=   "txtAttackDodgeMsgDisplay(1)"
            Tab(1).Control(22).Enabled=   0   'False
            Tab(1).Control(23)=   "txtAttackMissMsgDisplay(1)"
            Tab(1).Control(23).Enabled=   0   'False
            Tab(1).Control(24)=   "txtAttackHitSpellName(1)"
            Tab(1).Control(24).Enabled=   0   'False
            Tab(1).Control(25)=   "cmbAttackType(1)"
            Tab(1).Control(26)=   "cmdEditAttackSpell(1)"
            Tab(1).Control(27)=   "cmdEditHitMsg(1)"
            Tab(1).Control(28)=   "cmdEditMissMsg(1)"
            Tab(1).Control(29)=   "cmdEditDodgeMsg(1)"
            Tab(1).Control(30)=   "cmdEditHitSpell(1)"
            Tab(1).Control(31)=   "txtAttackSpellDamage(1)"
            Tab(1).Control(32)=   "cmdAttackSim(1)"
            Tab(1).ControlCount=   33
            TabCaption(2)   =   "Attack 3"
            TabPicture(2)   =   "frmMonster.frx":0F3A
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "lblAttackType(1)"
            Tab(2).Control(1)=   "lblAttackAccuSpell(2)"
            Tab(2).Control(2)=   "lblAttackMinHCastPer(2)"
            Tab(2).Control(3)=   "lblAttackMaxHCastLvl(2)"
            Tab(2).Control(4)=   "lblAttackPercent(1)"
            Tab(2).Control(5)=   "lblAttackEnergy(1)"
            Tab(2).Control(6)=   "lblAttackHitMsg(1)"
            Tab(2).Control(7)=   "lblAttackMissMsg(1)"
            Tab(2).Control(8)=   "lblAttackDodgeMsg(1)"
            Tab(2).Control(9)=   "lblAttackHitSpell(1)"
            Tab(2).Control(10)=   "lblAttackSpellRange(2)"
            Tab(2).Control(11)=   "txtAttackAccuSpell(2)"
            Tab(2).Control(12)=   "txtAttackAccuSpellName(2)"
            Tab(2).Control(12).Enabled=   0   'False
            Tab(2).Control(13)=   "txtAttackMinHCastPer(2)"
            Tab(2).Control(14)=   "txtAttackMaxHCastLvL(2)"
            Tab(2).Control(15)=   "txtAttackPer(2)"
            Tab(2).Control(16)=   "txtAttackEnergy(2)"
            Tab(2).Control(17)=   "txtAttackHitMsg(2)"
            Tab(2).Control(18)=   "txtAttackMissMsg(2)"
            Tab(2).Control(19)=   "txtAttackDodgeMsg(2)"
            Tab(2).Control(20)=   "txtAttackHitSpell(2)"
            Tab(2).Control(21)=   "txtAttackHitMsgDisplay(2)"
            Tab(2).Control(21).Enabled=   0   'False
            Tab(2).Control(22)=   "txtAttackDodgeMsgDisplay(2)"
            Tab(2).Control(22).Enabled=   0   'False
            Tab(2).Control(23)=   "txtAttackMissMsgDisplay(2)"
            Tab(2).Control(23).Enabled=   0   'False
            Tab(2).Control(24)=   "txtAttackHitSpellName(2)"
            Tab(2).Control(24).Enabled=   0   'False
            Tab(2).Control(25)=   "cmbAttackType(2)"
            Tab(2).Control(26)=   "cmdEditAttackSpell(2)"
            Tab(2).Control(27)=   "cmdEditHitMsg(2)"
            Tab(2).Control(28)=   "cmdEditMissMsg(2)"
            Tab(2).Control(29)=   "cmdEditDodgeMsg(2)"
            Tab(2).Control(30)=   "cmdEditHitSpell(2)"
            Tab(2).Control(31)=   "txtAttackSpellDamage(2)"
            Tab(2).Control(32)=   "cmdAttackSim(2)"
            Tab(2).ControlCount=   33
            TabCaption(3)   =   "Attack 4"
            TabPicture(3)   =   "frmMonster.frx":0F56
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "lblAttackType(2)"
            Tab(3).Control(1)=   "lblAttackAccuSpell(3)"
            Tab(3).Control(2)=   "lblAttackMinHCastPer(3)"
            Tab(3).Control(3)=   "lblAttackMaxHCastLvl(3)"
            Tab(3).Control(4)=   "lblAttackPercent(2)"
            Tab(3).Control(5)=   "lblAttackEnergy(2)"
            Tab(3).Control(6)=   "lblAttackHitMsg(2)"
            Tab(3).Control(7)=   "lblAttackMissMsg(2)"
            Tab(3).Control(8)=   "lblAttackDodgeMsg(2)"
            Tab(3).Control(9)=   "lblAttackHitSpell(2)"
            Tab(3).Control(10)=   "lblAttackSpellRange(3)"
            Tab(3).Control(11)=   "txtAttackAccuSpell(3)"
            Tab(3).Control(12)=   "txtAttackAccuSpellName(3)"
            Tab(3).Control(12).Enabled=   0   'False
            Tab(3).Control(13)=   "txtAttackMinHCastPer(3)"
            Tab(3).Control(14)=   "txtAttackMaxHCastLvL(3)"
            Tab(3).Control(15)=   "txtAttackPer(3)"
            Tab(3).Control(16)=   "txtAttackEnergy(3)"
            Tab(3).Control(17)=   "txtAttackHitMsg(3)"
            Tab(3).Control(18)=   "txtAttackMissMsg(3)"
            Tab(3).Control(19)=   "txtAttackDodgeMsg(3)"
            Tab(3).Control(20)=   "txtAttackHitSpell(3)"
            Tab(3).Control(21)=   "txtAttackHitMsgDisplay(3)"
            Tab(3).Control(21).Enabled=   0   'False
            Tab(3).Control(22)=   "txtAttackDodgeMsgDisplay(3)"
            Tab(3).Control(22).Enabled=   0   'False
            Tab(3).Control(23)=   "txtAttackMissMsgDisplay(3)"
            Tab(3).Control(23).Enabled=   0   'False
            Tab(3).Control(24)=   "txtAttackHitSpellName(3)"
            Tab(3).Control(24).Enabled=   0   'False
            Tab(3).Control(25)=   "cmbAttackType(3)"
            Tab(3).Control(26)=   "cmdEditAttackSpell(3)"
            Tab(3).Control(27)=   "cmdEditHitMsg(3)"
            Tab(3).Control(28)=   "cmdEditMissMsg(3)"
            Tab(3).Control(29)=   "cmdEditDodgeMsg(3)"
            Tab(3).Control(30)=   "cmdEditHitSpell(3)"
            Tab(3).Control(31)=   "txtAttackSpellDamage(3)"
            Tab(3).Control(32)=   "cmdAttackSim(3)"
            Tab(3).ControlCount=   33
            TabCaption(4)   =   "Attack 5"
            TabPicture(4)   =   "frmMonster.frx":0F72
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "lblAttackType(3)"
            Tab(4).Control(1)=   "lblAttackAccuSpell(4)"
            Tab(4).Control(2)=   "lblAttackMinHCastPer(4)"
            Tab(4).Control(3)=   "lblAttackMaxHCastLvl(4)"
            Tab(4).Control(4)=   "lblAttackPercent(3)"
            Tab(4).Control(5)=   "lblAttackEnergy(3)"
            Tab(4).Control(6)=   "lblAttackHitMsg(3)"
            Tab(4).Control(7)=   "lblAttackMissMsg(3)"
            Tab(4).Control(8)=   "lblAttackDodgeMsg(3)"
            Tab(4).Control(9)=   "lblAttackHitSpell(3)"
            Tab(4).Control(10)=   "lblAttackSpellRange(4)"
            Tab(4).Control(11)=   "txtAttackAccuSpell(4)"
            Tab(4).Control(12)=   "txtAttackAccuSpellName(4)"
            Tab(4).Control(12).Enabled=   0   'False
            Tab(4).Control(13)=   "txtAttackMinHCastPer(4)"
            Tab(4).Control(14)=   "txtAttackMaxHCastLvL(4)"
            Tab(4).Control(15)=   "txtAttackPer(4)"
            Tab(4).Control(16)=   "txtAttackEnergy(4)"
            Tab(4).Control(17)=   "txtAttackHitMsg(4)"
            Tab(4).Control(18)=   "txtAttackMissMsg(4)"
            Tab(4).Control(19)=   "txtAttackDodgeMsg(4)"
            Tab(4).Control(20)=   "txtAttackHitSpell(4)"
            Tab(4).Control(21)=   "txtAttackHitMsgDisplay(4)"
            Tab(4).Control(21).Enabled=   0   'False
            Tab(4).Control(22)=   "txtAttackDodgeMsgDisplay(4)"
            Tab(4).Control(22).Enabled=   0   'False
            Tab(4).Control(23)=   "txtAttackMissMsgDisplay(4)"
            Tab(4).Control(23).Enabled=   0   'False
            Tab(4).Control(24)=   "txtAttackHitSpellName(4)"
            Tab(4).Control(24).Enabled=   0   'False
            Tab(4).Control(25)=   "cmbAttackType(4)"
            Tab(4).Control(26)=   "cmdEditAttackSpell(4)"
            Tab(4).Control(27)=   "cmdEditHitMsg(4)"
            Tab(4).Control(28)=   "cmdEditMissMsg(4)"
            Tab(4).Control(29)=   "cmdEditDodgeMsg(4)"
            Tab(4).Control(30)=   "cmdEditHitSpell(4)"
            Tab(4).Control(31)=   "txtAttackSpellDamage(4)"
            Tab(4).Control(32)=   "cmdAttackSim(4)"
            Tab(4).ControlCount=   33
            TabCaption(5)   =   "Copy/Paste"
            TabPicture(5)   =   "frmMonster.frx":0F8E
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "Label4"
            Tab(5).Control(1)=   "cmdAttackCopyAll(0)"
            Tab(5).Control(2)=   "cmdAttackCopyAll(1)"
            Tab(5).Control(3)=   "cmdAttackCopySingle(0)"
            Tab(5).Control(4)=   "cmdAttackCopySingle(1)"
            Tab(5).Control(5)=   "cmdAttackCopySingle(2)"
            Tab(5).Control(6)=   "cmdAttackCopySingle(3)"
            Tab(5).Control(7)=   "cmdAttackCopySingle(4)"
            Tab(5).Control(8)=   "cmdAttackCopySingle(5)"
            Tab(5).Control(9)=   "cmdAttackCopySingle(6)"
            Tab(5).Control(10)=   "cmdAttackCopySingle(7)"
            Tab(5).Control(11)=   "cmdAttackCopySingle(8)"
            Tab(5).Control(12)=   "cmdAttackCopySingle(9)"
            Tab(5).Control(13)=   "cmdAttackClear"
            Tab(5).ControlCount=   14
            Begin VB.CommandButton cmdAttackSim 
               Caption         =   "Open Combat Sim."
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
               Index           =   4
               Left            =   -71820
               TabIndex        =   469
               Top             =   420
               Width           =   1995
            End
            Begin VB.CommandButton cmdAttackSim 
               Caption         =   "Open Combat Sim."
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
               Index           =   3
               Left            =   -71820
               TabIndex        =   468
               Top             =   420
               Width           =   1995
            End
            Begin VB.CommandButton cmdAttackSim 
               Caption         =   "Open Combat Sim."
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
               Index           =   2
               Left            =   -71820
               TabIndex        =   467
               Top             =   420
               Width           =   1995
            End
            Begin VB.CommandButton cmdAttackSim 
               Caption         =   "Open Combat Sim."
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
               Left            =   -71820
               TabIndex        =   466
               Top             =   420
               Width           =   1995
            End
            Begin VB.CommandButton cmdAttackSim 
               Caption         =   "Open Combat Sim."
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
               Left            =   3180
               TabIndex        =   465
               Top             =   420
               Width           =   1995
            End
            Begin VB.TextBox txtAttackSpellDamage 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   4
               Left            =   -71880
               Locked          =   -1  'True
               TabIndex        =   414
               Top             =   1500
               Width           =   2055
            End
            Begin VB.TextBox txtAttackSpellDamage 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   3
               Left            =   -71880
               Locked          =   -1  'True
               TabIndex        =   413
               Top             =   1500
               Width           =   2055
            End
            Begin VB.TextBox txtAttackSpellDamage 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   2
               Left            =   -71880
               Locked          =   -1  'True
               TabIndex        =   412
               Top             =   1500
               Width           =   2055
            End
            Begin VB.TextBox txtAttackSpellDamage 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   1
               Left            =   -71880
               Locked          =   -1  'True
               TabIndex        =   411
               Top             =   1500
               Width           =   2055
            End
            Begin VB.TextBox txtAttackSpellDamage 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   0
               Left            =   3120
               Locked          =   -1  'True
               TabIndex        =   410
               Top             =   1500
               Width           =   2055
            End
            Begin VB.CommandButton cmdAttackClear 
               Caption         =   "Clear Attacks"
               Height          =   375
               Left            =   -71340
               TabIndex        =   409
               Top             =   2760
               Width           =   1215
            End
            Begin VB.CommandButton cmdAttackCopySingle 
               Caption         =   "Paste to Attack 5"
               Height          =   375
               Index           =   9
               Left            =   -73260
               TabIndex        =   407
               Top             =   2760
               Width           =   1515
            End
            Begin VB.CommandButton cmdAttackCopySingle 
               Caption         =   "Paste to Attack 4"
               Height          =   375
               Index           =   8
               Left            =   -73260
               TabIndex        =   406
               Top             =   2340
               Width           =   1515
            End
            Begin VB.CommandButton cmdAttackCopySingle 
               Caption         =   "Paste to Attack 3"
               Height          =   375
               Index           =   7
               Left            =   -73260
               TabIndex        =   405
               Top             =   1920
               Width           =   1515
            End
            Begin VB.CommandButton cmdAttackCopySingle 
               Caption         =   "Paste to Attack 2"
               Height          =   375
               Index           =   6
               Left            =   -73260
               TabIndex        =   404
               Top             =   1500
               Width           =   1515
            End
            Begin VB.CommandButton cmdAttackCopySingle 
               Caption         =   "Paste to Attack 1"
               Height          =   375
               Index           =   5
               Left            =   -73260
               TabIndex        =   403
               Top             =   1080
               Width           =   1515
            End
            Begin VB.CommandButton cmdAttackCopySingle 
               Caption         =   "Copy Attack 5"
               Height          =   375
               Index           =   4
               Left            =   -74820
               TabIndex        =   402
               Top             =   2760
               Width           =   1395
            End
            Begin VB.CommandButton cmdAttackCopySingle 
               Caption         =   "Copy Attack 4"
               Height          =   375
               Index           =   3
               Left            =   -74820
               TabIndex        =   401
               Top             =   2340
               Width           =   1395
            End
            Begin VB.CommandButton cmdAttackCopySingle 
               Caption         =   "Copy Attack 3"
               Height          =   375
               Index           =   2
               Left            =   -74820
               TabIndex        =   400
               Top             =   1920
               Width           =   1395
            End
            Begin VB.CommandButton cmdAttackCopySingle 
               Caption         =   "Copy Attack 2"
               Height          =   375
               Index           =   1
               Left            =   -74820
               TabIndex        =   399
               Top             =   1500
               Width           =   1395
            End
            Begin VB.CommandButton cmdAttackCopySingle 
               Caption         =   "Copy Attack 1"
               Height          =   375
               Index           =   0
               Left            =   -74820
               TabIndex        =   398
               Top             =   1080
               Width           =   1395
            End
            Begin VB.CommandButton cmdAttackCopyAll 
               Caption         =   "Paste All"
               Height          =   375
               Index           =   1
               Left            =   -73260
               TabIndex        =   397
               Top             =   480
               Width           =   1515
            End
            Begin VB.CommandButton cmdAttackCopyAll 
               Caption         =   "Copy All"
               Height          =   375
               Index           =   0
               Left            =   -74820
               TabIndex        =   396
               Top             =   480
               Width           =   1395
            End
            Begin VB.CommandButton cmdEditHitSpell 
               Height          =   195
               Index           =   4
               Left            =   -73980
               TabIndex        =   244
               Top             =   3000
               Width           =   135
            End
            Begin VB.CommandButton cmdEditHitSpell 
               Height          =   195
               Index           =   3
               Left            =   -73980
               TabIndex        =   224
               Top             =   3000
               Width           =   135
            End
            Begin VB.CommandButton cmdEditHitSpell 
               Height          =   195
               Index           =   2
               Left            =   -73980
               TabIndex        =   204
               Top             =   3000
               Width           =   135
            End
            Begin VB.CommandButton cmdEditHitSpell 
               Height          =   195
               Index           =   1
               Left            =   -73980
               TabIndex        =   184
               Top             =   3000
               Width           =   135
            End
            Begin VB.CommandButton cmdEditHitSpell 
               Height          =   195
               Index           =   0
               Left            =   1020
               TabIndex        =   164
               Top             =   3000
               Width           =   135
            End
            Begin VB.CommandButton cmdEditDodgeMsg 
               Height          =   195
               Index           =   4
               Left            =   -73980
               TabIndex        =   241
               Top             =   2640
               Width           =   135
            End
            Begin VB.CommandButton cmdEditDodgeMsg 
               Height          =   195
               Index           =   3
               Left            =   -73980
               TabIndex        =   221
               Top             =   2640
               Width           =   135
            End
            Begin VB.CommandButton cmdEditDodgeMsg 
               Height          =   195
               Index           =   2
               Left            =   -73980
               TabIndex        =   201
               Top             =   2640
               Width           =   135
            End
            Begin VB.CommandButton cmdEditDodgeMsg 
               Height          =   195
               Index           =   1
               Left            =   -73980
               TabIndex        =   181
               Top             =   2640
               Width           =   135
            End
            Begin VB.CommandButton cmdEditDodgeMsg 
               Height          =   195
               Index           =   0
               Left            =   1020
               TabIndex        =   161
               Top             =   2640
               Width           =   135
            End
            Begin VB.CommandButton cmdEditMissMsg 
               Height          =   195
               Index           =   4
               Left            =   -73980
               TabIndex        =   238
               Top             =   2280
               Width           =   135
            End
            Begin VB.CommandButton cmdEditMissMsg 
               Height          =   195
               Index           =   3
               Left            =   -73980
               TabIndex        =   218
               Top             =   2280
               Width           =   135
            End
            Begin VB.CommandButton cmdEditMissMsg 
               Height          =   195
               Index           =   2
               Left            =   -73980
               TabIndex        =   198
               Top             =   2280
               Width           =   135
            End
            Begin VB.CommandButton cmdEditMissMsg 
               Height          =   195
               Index           =   1
               Left            =   -73980
               TabIndex        =   178
               Top             =   2280
               Width           =   135
            End
            Begin VB.CommandButton cmdEditMissMsg 
               Height          =   195
               Index           =   0
               Left            =   1020
               TabIndex        =   158
               Top             =   2280
               Width           =   135
            End
            Begin VB.CommandButton cmdEditHitMsg 
               Height          =   195
               Index           =   4
               Left            =   -73980
               TabIndex        =   235
               Top             =   1920
               Width           =   135
            End
            Begin VB.CommandButton cmdEditHitMsg 
               Height          =   195
               Index           =   3
               Left            =   -73980
               TabIndex        =   215
               Top             =   1920
               Width           =   135
            End
            Begin VB.CommandButton cmdEditHitMsg 
               Height          =   195
               Index           =   2
               Left            =   -73980
               TabIndex        =   195
               Top             =   1920
               Width           =   135
            End
            Begin VB.CommandButton cmdEditHitMsg 
               Height          =   195
               Index           =   1
               Left            =   -73980
               TabIndex        =   175
               Top             =   1920
               Width           =   135
            End
            Begin VB.CommandButton cmdEditHitMsg 
               Height          =   195
               Index           =   0
               Left            =   1020
               TabIndex        =   155
               Top             =   1920
               Width           =   135
            End
            Begin VB.CommandButton cmdEditAttackSpell 
               Height          =   195
               Index           =   4
               Left            =   -73980
               TabIndex        =   228
               Top             =   840
               Width           =   135
            End
            Begin VB.CommandButton cmdEditAttackSpell 
               Height          =   195
               Index           =   3
               Left            =   -73980
               TabIndex        =   208
               Top             =   840
               Width           =   135
            End
            Begin VB.CommandButton cmdEditAttackSpell 
               Height          =   195
               Index           =   2
               Left            =   -73980
               TabIndex        =   188
               Top             =   840
               Width           =   135
            End
            Begin VB.CommandButton cmdEditAttackSpell 
               Height          =   195
               Index           =   1
               Left            =   -73980
               TabIndex        =   168
               Top             =   840
               Width           =   135
            End
            Begin VB.CommandButton cmdEditAttackSpell 
               Height          =   195
               Index           =   0
               Left            =   1020
               TabIndex        =   148
               Top             =   840
               Width           =   135
            End
            Begin VB.ComboBox cmbAttackType 
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
               ItemData        =   "frmMonster.frx":0FAA
               Left            =   -73800
               List            =   "frmMonster.frx":0FBA
               Style           =   2  'Dropdown List
               TabIndex        =   227
               Top             =   420
               Width           =   1455
            End
            Begin VB.ComboBox cmbAttackType 
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
               ItemData        =   "frmMonster.frx":0FD8
               Left            =   -73800
               List            =   "frmMonster.frx":0FE8
               Style           =   2  'Dropdown List
               TabIndex        =   207
               Top             =   420
               Width           =   1455
            End
            Begin VB.ComboBox cmbAttackType 
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
               ItemData        =   "frmMonster.frx":1006
               Left            =   -73800
               List            =   "frmMonster.frx":1016
               Style           =   2  'Dropdown List
               TabIndex        =   187
               Top             =   420
               Width           =   1455
            End
            Begin VB.ComboBox cmbAttackType 
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
               ItemData        =   "frmMonster.frx":1034
               Left            =   -73800
               List            =   "frmMonster.frx":1044
               Style           =   2  'Dropdown List
               TabIndex        =   167
               Top             =   420
               Width           =   1455
            End
            Begin VB.TextBox txtAttackHitSpellName 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   -73080
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   246
               TabStop         =   0   'False
               Top             =   2940
               Width           =   3255
            End
            Begin VB.TextBox txtAttackMissMsgDisplay 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   -73080
               Locked          =   -1  'True
               MaxLength       =   75
               TabIndex        =   240
               TabStop         =   0   'False
               Top             =   2220
               Width           =   3255
            End
            Begin VB.TextBox txtAttackDodgeMsgDisplay 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   -73080
               Locked          =   -1  'True
               MaxLength       =   75
               TabIndex        =   243
               TabStop         =   0   'False
               Top             =   2580
               Width           =   3255
            End
            Begin VB.TextBox txtAttackHitMsgDisplay 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   -73080
               Locked          =   -1  'True
               MaxLength       =   75
               TabIndex        =   237
               TabStop         =   0   'False
               Top             =   1860
               Width           =   3255
            End
            Begin VB.TextBox txtAttackHitSpell 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   -73800
               TabIndex        =   245
               Top             =   2940
               Width           =   615
            End
            Begin VB.TextBox txtAttackDodgeMsg 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   -73800
               TabIndex        =   242
               Top             =   2580
               Width           =   615
            End
            Begin VB.TextBox txtAttackMissMsg 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   -73800
               TabIndex        =   239
               Top             =   2220
               Width           =   615
            End
            Begin VB.TextBox txtAttackHitMsg 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   -73800
               TabIndex        =   236
               Top             =   1860
               Width           =   615
            End
            Begin VB.TextBox txtAttackEnergy 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   -70440
               TabIndex        =   234
               Top             =   1140
               Width           =   615
            End
            Begin VB.TextBox txtAttackPer 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   -73800
               TabIndex        =   233
               Top             =   1500
               Width           =   615
            End
            Begin VB.TextBox txtAttackMaxHCastLvL 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   -71880
               TabIndex        =   232
               Top             =   1140
               Width           =   615
            End
            Begin VB.TextBox txtAttackMinHCastPer 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   -73800
               TabIndex        =   231
               Top             =   1140
               Width           =   615
            End
            Begin VB.TextBox txtAttackAccuSpellName 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   -73080
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   230
               TabStop         =   0   'False
               Top             =   780
               Width           =   3255
            End
            Begin VB.TextBox txtAttackAccuSpell 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   -73800
               TabIndex        =   229
               Top             =   780
               Width           =   615
            End
            Begin VB.TextBox txtAttackHitSpellName 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   -73080
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   226
               TabStop         =   0   'False
               Top             =   2940
               Width           =   3255
            End
            Begin VB.TextBox txtAttackMissMsgDisplay 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   -73080
               Locked          =   -1  'True
               MaxLength       =   75
               TabIndex        =   220
               TabStop         =   0   'False
               Top             =   2220
               Width           =   3255
            End
            Begin VB.TextBox txtAttackDodgeMsgDisplay 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   -73080
               Locked          =   -1  'True
               MaxLength       =   75
               TabIndex        =   223
               TabStop         =   0   'False
               Top             =   2580
               Width           =   3255
            End
            Begin VB.TextBox txtAttackHitMsgDisplay 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   -73080
               Locked          =   -1  'True
               MaxLength       =   75
               TabIndex        =   217
               TabStop         =   0   'False
               Top             =   1860
               Width           =   3255
            End
            Begin VB.TextBox txtAttackHitSpell 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   -73800
               TabIndex        =   225
               Top             =   2940
               Width           =   615
            End
            Begin VB.TextBox txtAttackDodgeMsg 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   -73800
               TabIndex        =   222
               Top             =   2580
               Width           =   615
            End
            Begin VB.TextBox txtAttackMissMsg 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   -73800
               TabIndex        =   219
               Top             =   2220
               Width           =   615
            End
            Begin VB.TextBox txtAttackHitMsg 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   -73800
               TabIndex        =   216
               Top             =   1860
               Width           =   615
            End
            Begin VB.TextBox txtAttackEnergy 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   -70440
               TabIndex        =   214
               Top             =   1140
               Width           =   615
            End
            Begin VB.TextBox txtAttackPer 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   -73800
               TabIndex        =   213
               Top             =   1500
               Width           =   615
            End
            Begin VB.TextBox txtAttackMaxHCastLvL 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   -71880
               TabIndex        =   212
               Top             =   1140
               Width           =   615
            End
            Begin VB.TextBox txtAttackMinHCastPer 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   -73800
               TabIndex        =   211
               Top             =   1140
               Width           =   615
            End
            Begin VB.TextBox txtAttackAccuSpellName 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   -73080
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   210
               TabStop         =   0   'False
               Top             =   780
               Width           =   3255
            End
            Begin VB.TextBox txtAttackAccuSpell 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   -73800
               TabIndex        =   209
               Top             =   780
               Width           =   615
            End
            Begin VB.TextBox txtAttackHitSpellName 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   -73080
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   206
               TabStop         =   0   'False
               Top             =   2940
               Width           =   3255
            End
            Begin VB.TextBox txtAttackMissMsgDisplay 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   -73080
               Locked          =   -1  'True
               MaxLength       =   75
               TabIndex        =   200
               TabStop         =   0   'False
               Top             =   2220
               Width           =   3255
            End
            Begin VB.TextBox txtAttackDodgeMsgDisplay 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   -73080
               Locked          =   -1  'True
               MaxLength       =   75
               TabIndex        =   203
               TabStop         =   0   'False
               Top             =   2580
               Width           =   3255
            End
            Begin VB.TextBox txtAttackHitMsgDisplay 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   -73080
               Locked          =   -1  'True
               MaxLength       =   75
               TabIndex        =   197
               TabStop         =   0   'False
               Top             =   1860
               Width           =   3255
            End
            Begin VB.TextBox txtAttackHitSpell 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   -73800
               TabIndex        =   205
               Top             =   2940
               Width           =   615
            End
            Begin VB.TextBox txtAttackDodgeMsg 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   -73800
               TabIndex        =   202
               Top             =   2580
               Width           =   615
            End
            Begin VB.TextBox txtAttackMissMsg 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   -73800
               TabIndex        =   199
               Top             =   2220
               Width           =   615
            End
            Begin VB.TextBox txtAttackHitMsg 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   -73800
               TabIndex        =   196
               Top             =   1860
               Width           =   615
            End
            Begin VB.TextBox txtAttackEnergy 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   -70440
               TabIndex        =   194
               Top             =   1140
               Width           =   615
            End
            Begin VB.TextBox txtAttackPer 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   -73800
               TabIndex        =   193
               Top             =   1500
               Width           =   615
            End
            Begin VB.TextBox txtAttackMaxHCastLvL 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   -71880
               TabIndex        =   192
               Top             =   1140
               Width           =   615
            End
            Begin VB.TextBox txtAttackMinHCastPer 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   -73800
               TabIndex        =   191
               Top             =   1140
               Width           =   615
            End
            Begin VB.TextBox txtAttackAccuSpellName 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   -73080
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   190
               TabStop         =   0   'False
               Top             =   780
               Width           =   3255
            End
            Begin VB.TextBox txtAttackAccuSpell 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   -73800
               TabIndex        =   189
               Top             =   780
               Width           =   615
            End
            Begin VB.TextBox txtAttackHitSpellName 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   -73080
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   186
               TabStop         =   0   'False
               Top             =   2940
               Width           =   3255
            End
            Begin VB.TextBox txtAttackMissMsgDisplay 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   -73080
               Locked          =   -1  'True
               MaxLength       =   75
               TabIndex        =   180
               TabStop         =   0   'False
               Top             =   2220
               Width           =   3255
            End
            Begin VB.TextBox txtAttackDodgeMsgDisplay 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   -73080
               Locked          =   -1  'True
               MaxLength       =   75
               TabIndex        =   183
               TabStop         =   0   'False
               Top             =   2580
               Width           =   3255
            End
            Begin VB.TextBox txtAttackHitMsgDisplay 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   -73080
               Locked          =   -1  'True
               MaxLength       =   75
               TabIndex        =   177
               TabStop         =   0   'False
               Top             =   1860
               Width           =   3255
            End
            Begin VB.TextBox txtAttackHitSpell 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   -73800
               TabIndex        =   185
               Top             =   2940
               Width           =   615
            End
            Begin VB.TextBox txtAttackDodgeMsg 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   -73800
               TabIndex        =   182
               Top             =   2580
               Width           =   615
            End
            Begin VB.TextBox txtAttackMissMsg 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   -73800
               TabIndex        =   179
               Top             =   2220
               Width           =   615
            End
            Begin VB.TextBox txtAttackHitMsg 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   -73800
               TabIndex        =   176
               Top             =   1860
               Width           =   615
            End
            Begin VB.TextBox txtAttackEnergy 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   -70440
               TabIndex        =   174
               Top             =   1140
               Width           =   615
            End
            Begin VB.TextBox txtAttackPer 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   -73800
               TabIndex        =   173
               Top             =   1500
               Width           =   615
            End
            Begin VB.TextBox txtAttackMaxHCastLvL 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   -71880
               TabIndex        =   172
               Top             =   1140
               Width           =   615
            End
            Begin VB.TextBox txtAttackMinHCastPer 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   -73800
               TabIndex        =   171
               Top             =   1140
               Width           =   615
            End
            Begin VB.TextBox txtAttackAccuSpellName 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   -73080
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   170
               TabStop         =   0   'False
               Top             =   780
               Width           =   3255
            End
            Begin VB.TextBox txtAttackAccuSpell 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   -73800
               TabIndex        =   169
               Top             =   780
               Width           =   615
            End
            Begin VB.ComboBox cmbAttackType 
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
               ItemData        =   "frmMonster.frx":1062
               Left            =   1200
               List            =   "frmMonster.frx":1072
               Style           =   2  'Dropdown List
               TabIndex        =   147
               Top             =   420
               Width           =   1455
            End
            Begin VB.TextBox txtAttackAccuSpell 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   1200
               TabIndex        =   149
               Top             =   780
               Width           =   615
            End
            Begin VB.TextBox txtAttackAccuSpellName 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   150
               TabStop         =   0   'False
               Top             =   780
               Width           =   3255
            End
            Begin VB.TextBox txtAttackMinHCastPer 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   1200
               TabIndex        =   151
               Top             =   1140
               Width           =   615
            End
            Begin VB.TextBox txtAttackMaxHCastLvL 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   3120
               TabIndex        =   152
               Top             =   1140
               Width           =   615
            End
            Begin VB.TextBox txtAttackPer 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   1200
               TabIndex        =   153
               Top             =   1500
               Width           =   615
            End
            Begin VB.TextBox txtAttackEnergy 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   4560
               TabIndex        =   154
               Top             =   1140
               Width           =   615
            End
            Begin VB.TextBox txtAttackHitMsg 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   1200
               TabIndex        =   156
               Top             =   1860
               Width           =   615
            End
            Begin VB.TextBox txtAttackMissMsg 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   1200
               TabIndex        =   159
               Top             =   2220
               Width           =   615
            End
            Begin VB.TextBox txtAttackDodgeMsg 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   1200
               TabIndex        =   162
               Top             =   2580
               Width           =   615
            End
            Begin VB.TextBox txtAttackHitSpell 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   1200
               TabIndex        =   165
               Top             =   2940
               Width           =   615
            End
            Begin VB.TextBox txtAttackHitMsgDisplay 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   75
               TabIndex        =   157
               TabStop         =   0   'False
               Top             =   1860
               Width           =   3255
            End
            Begin VB.TextBox txtAttackDodgeMsgDisplay 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   75
               TabIndex        =   163
               TabStop         =   0   'False
               Top             =   2580
               Width           =   3255
            End
            Begin VB.TextBox txtAttackMissMsgDisplay 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   75
               TabIndex        =   160
               TabStop         =   0   'False
               Top             =   2220
               Width           =   3255
            End
            Begin VB.TextBox txtAttackHitSpellName 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   166
               TabStop         =   0   'False
               Top             =   2940
               Width           =   3255
            End
            Begin VB.Label lblAttackSpellRange 
               Alignment       =   1  'Right Justify
               Caption         =   "Spell Range:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   -73140
               TabIndex        =   419
               Top             =   1500
               Width           =   1155
            End
            Begin VB.Label lblAttackSpellRange 
               Alignment       =   1  'Right Justify
               Caption         =   "Spell Range:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   -73140
               TabIndex        =   418
               Top             =   1500
               Width           =   1155
            End
            Begin VB.Label lblAttackSpellRange 
               Alignment       =   1  'Right Justify
               Caption         =   "Spell Range:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   -73140
               TabIndex        =   417
               Top             =   1500
               Width           =   1155
            End
            Begin VB.Label lblAttackSpellRange 
               Alignment       =   1  'Right Justify
               Caption         =   "Spell Range:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   -73140
               TabIndex        =   416
               Top             =   1500
               Width           =   1155
            End
            Begin VB.Label lblAttackSpellRange 
               Alignment       =   1  'Right Justify
               Caption         =   "Spell Range:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   1860
               TabIndex        =   415
               Top             =   1500
               Width           =   1155
            End
            Begin VB.Label Label4 
               Caption         =   $"frmMonster.frx":1090
               Enabled         =   0   'False
               Height          =   2235
               Left            =   -71460
               TabIndex        =   408
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label lblAttackHitSpell 
               Caption         =   "Hit Spell"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   -74880
               TabIndex        =   296
               Top             =   3000
               Width           =   855
            End
            Begin VB.Label lblAttackDodgeMsg 
               Caption         =   "Dodge Msg"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   -74880
               TabIndex        =   295
               Top             =   2640
               Width           =   855
            End
            Begin VB.Label lblAttackMissMsg 
               Caption         =   "Miss Msg"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   -74880
               TabIndex        =   294
               Top             =   2280
               Width           =   855
            End
            Begin VB.Label lblAttackHitMsg 
               Caption         =   "Hit Msg"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   -74880
               TabIndex        =   293
               Top             =   1920
               Width           =   855
            End
            Begin VB.Label lblAttackEnergy 
               Alignment       =   1  'Right Justify
               Caption         =   "Energy"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   -71640
               TabIndex        =   292
               Top             =   1200
               Width           =   1095
            End
            Begin VB.Label lblAttackPercent 
               Caption         =   "Attack%"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   -74880
               TabIndex        =   291
               Top             =   1560
               Width           =   1095
            End
            Begin VB.Label lblAttackMaxHCastLvl 
               Alignment       =   1  'Right Justify
               Caption         =   "MaxH/CastLvL"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   -73080
               TabIndex        =   290
               Top             =   1200
               Width           =   1095
            End
            Begin VB.Label lblAttackMinHCastPer 
               Caption         =   "MinH/Cast%"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   -74880
               TabIndex        =   289
               Top             =   1200
               Width           =   1095
            End
            Begin VB.Label lblAttackAccuSpell 
               Caption         =   "Accu/Spell"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   -74880
               TabIndex        =   288
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label lblAttackType 
               Caption         =   "Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   -74880
               TabIndex        =   287
               Top             =   480
               Width           =   735
            End
            Begin VB.Label lblAttackHitSpell 
               Caption         =   "Hit Spell"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   -74880
               TabIndex        =   286
               Top             =   3000
               Width           =   855
            End
            Begin VB.Label lblAttackDodgeMsg 
               Caption         =   "Dodge Msg"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   -74880
               TabIndex        =   285
               Top             =   2640
               Width           =   855
            End
            Begin VB.Label lblAttackMissMsg 
               Caption         =   "Miss Msg"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   -74880
               TabIndex        =   284
               Top             =   2280
               Width           =   855
            End
            Begin VB.Label lblAttackHitMsg 
               Caption         =   "Hit Msg"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   -74880
               TabIndex        =   283
               Top             =   1920
               Width           =   855
            End
            Begin VB.Label lblAttackEnergy 
               Alignment       =   1  'Right Justify
               Caption         =   "Energy"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   -71640
               TabIndex        =   282
               Top             =   1200
               Width           =   1095
            End
            Begin VB.Label lblAttackPercent 
               Caption         =   "Attack%"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   -74880
               TabIndex        =   281
               Top             =   1560
               Width           =   1095
            End
            Begin VB.Label lblAttackMaxHCastLvl 
               Alignment       =   1  'Right Justify
               Caption         =   "MaxH/CastLvL"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   -73080
               TabIndex        =   280
               Top             =   1200
               Width           =   1095
            End
            Begin VB.Label lblAttackMinHCastPer 
               Caption         =   "MinH/Cast%"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   -74880
               TabIndex        =   279
               Top             =   1200
               Width           =   1095
            End
            Begin VB.Label lblAttackAccuSpell 
               Caption         =   "Accu/Spell"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   -74880
               TabIndex        =   278
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label lblAttackType 
               Caption         =   "Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   -74880
               TabIndex        =   277
               Top             =   480
               Width           =   735
            End
            Begin VB.Label lblAttackHitSpell 
               Caption         =   "Hit Spell"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   -74880
               TabIndex        =   276
               Top             =   3000
               Width           =   855
            End
            Begin VB.Label lblAttackDodgeMsg 
               Caption         =   "Dodge Msg"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   -74880
               TabIndex        =   275
               Top             =   2640
               Width           =   855
            End
            Begin VB.Label lblAttackMissMsg 
               Caption         =   "Miss Msg"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   -74880
               TabIndex        =   274
               Top             =   2280
               Width           =   855
            End
            Begin VB.Label lblAttackHitMsg 
               Caption         =   "Hit Msg"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   -74880
               TabIndex        =   273
               Top             =   1920
               Width           =   855
            End
            Begin VB.Label lblAttackEnergy 
               Alignment       =   1  'Right Justify
               Caption         =   "Energy"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   -71640
               TabIndex        =   272
               Top             =   1200
               Width           =   1095
            End
            Begin VB.Label lblAttackPercent 
               Caption         =   "Attack%"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   -74880
               TabIndex        =   271
               Top             =   1560
               Width           =   1095
            End
            Begin VB.Label lblAttackMaxHCastLvl 
               Alignment       =   1  'Right Justify
               Caption         =   "MaxH/CastLvL"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   -73080
               TabIndex        =   270
               Top             =   1200
               Width           =   1095
            End
            Begin VB.Label lblAttackMinHCastPer 
               Caption         =   "MinH/Cast%"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   -74880
               TabIndex        =   269
               Top             =   1200
               Width           =   1095
            End
            Begin VB.Label lblAttackAccuSpell 
               Caption         =   "Accu/Spell"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   -74880
               TabIndex        =   268
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label lblAttackType 
               Caption         =   "Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   -74880
               TabIndex        =   267
               Top             =   480
               Width           =   735
            End
            Begin VB.Label lblAttackHitSpell 
               Caption         =   "Hit Spell"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   -74880
               TabIndex        =   266
               Top             =   3000
               Width           =   855
            End
            Begin VB.Label lblAttackDodgeMsg 
               Caption         =   "Dodge Msg"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   -74880
               TabIndex        =   265
               Top             =   2640
               Width           =   855
            End
            Begin VB.Label lblAttackMissMsg 
               Caption         =   "Miss Msg"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   -74880
               TabIndex        =   264
               Top             =   2280
               Width           =   855
            End
            Begin VB.Label lblAttackHitMsg 
               Caption         =   "Hit Msg"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   -74880
               TabIndex        =   263
               Top             =   1920
               Width           =   855
            End
            Begin VB.Label lblAttackEnergy 
               Alignment       =   1  'Right Justify
               Caption         =   "Energy"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   -71640
               TabIndex        =   262
               Top             =   1200
               Width           =   1095
            End
            Begin VB.Label lblAttackPercent 
               Caption         =   "Attack%"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   -74880
               TabIndex        =   261
               Top             =   1560
               Width           =   1095
            End
            Begin VB.Label lblAttackMaxHCastLvl 
               Alignment       =   1  'Right Justify
               Caption         =   "MaxH/CastLvL"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   -73080
               TabIndex        =   260
               Top             =   1200
               Width           =   1095
            End
            Begin VB.Label lblAttackMinHCastPer 
               Caption         =   "MinH/Cast%"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   -74880
               TabIndex        =   259
               Top             =   1200
               Width           =   1095
            End
            Begin VB.Label lblAttackAccuSpell 
               Caption         =   "Accu/Spell"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   -74880
               TabIndex        =   258
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label lblAttackType 
               Caption         =   "Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   -74880
               TabIndex        =   257
               Top             =   480
               Width           =   735
            End
            Begin VB.Label lblAttackType 
               Caption         =   "Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   43
               Left            =   120
               TabIndex        =   256
               Top             =   480
               Width           =   735
            End
            Begin VB.Label lblAttackAccuSpell 
               Caption         =   "Accu/Spell"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   255
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label lblAttackMinHCastPer 
               Caption         =   "MinH/Cast%"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   254
               Top             =   1200
               Width           =   975
            End
            Begin VB.Label lblAttackMaxHCastLvl 
               Alignment       =   1  'Right Justify
               Caption         =   "MaxH/CastLvL"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   1860
               TabIndex        =   253
               Top             =   1200
               Width           =   1155
            End
            Begin VB.Label lblAttackPercent 
               Caption         =   "Attack%"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   47
               Left            =   120
               TabIndex        =   252
               Top             =   1560
               Width           =   1095
            End
            Begin VB.Label lblAttackEnergy 
               Alignment       =   1  'Right Justify
               Caption         =   "Energy"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   48
               Left            =   3840
               TabIndex        =   251
               Top             =   1200
               Width           =   615
            End
            Begin VB.Label lblAttackHitMsg 
               Caption         =   "Hit Msg"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   49
               Left            =   120
               TabIndex        =   250
               Top             =   1920
               Width           =   855
            End
            Begin VB.Label lblAttackMissMsg 
               Caption         =   "Miss Msg"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   50
               Left            =   120
               TabIndex        =   249
               Top             =   2280
               Width           =   855
            End
            Begin VB.Label lblAttackDodgeMsg 
               Caption         =   "Dodge Msg"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   51
               Left            =   120
               TabIndex        =   248
               Top             =   2640
               Width           =   855
            End
            Begin VB.Label lblAttackHitSpell 
               Caption         =   "Hit Spell"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   52
               Left            =   120
               TabIndex        =   247
               Top             =   3000
               Width           =   855
            End
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Caption         =   $"frmMonster.frx":1162
            Height          =   1515
            Index           =   1
            Left            =   -73860
            TabIndex        =   394
            Top             =   4200
            Width           =   3675
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "Type + and - in the '#' field to cycle though the abilities.  You can also type the name of the ability in the 'Name' field."
            Height          =   675
            Left            =   -73980
            TabIndex        =   393
            Top             =   4680
            Width           =   3615
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Caption         =   "NOTE: Monster Name + Max ""Desc Txt"" Length may not exceed 28 chars."
            Height          =   555
            Index           =   0
            Left            =   -73920
            TabIndex        =   347
            Top             =   3480
            Width           =   3675
         End
         Begin VB.Label label 
            Caption         =   "Talk Txt"
            Height          =   255
            Index           =   69
            Left            =   -74280
            TabIndex        =   340
            Top             =   2220
            Width           =   735
         End
         Begin VB.Label label 
            Caption         =   "Desc Txt"
            Height          =   255
            Index           =   70
            Left            =   -74280
            TabIndex        =   344
            Top             =   2580
            Width           =   735
         End
         Begin VB.Label label 
            Caption         =   "Greet Txt"
            Height          =   255
            Index           =   71
            Left            =   -74280
            TabIndex        =   336
            Top             =   1860
            Width           =   735
         End
         Begin VB.Label label 
            Caption         =   "Death Msg"
            Height          =   255
            Index           =   72
            Left            =   -74280
            TabIndex        =   332
            Top             =   1380
            Width           =   855
         End
         Begin VB.Label label 
            Caption         =   "Move Msg"
            Height          =   255
            Index           =   73
            Left            =   -74280
            TabIndex        =   328
            Top             =   1020
            Width           =   855
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Death Spell"
            Height          =   255
            Index           =   41
            Left            =   -74040
            TabIndex        =   113
            Top             =   720
            Width           =   855
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Create Spell"
            Height          =   255
            Index           =   10
            Left            =   -74100
            TabIndex        =   109
            Top             =   420
            Width           =   915
         End
         Begin VB.Label label 
            Caption         =   "Weapon"
            Height          =   255
            Index           =   17
            Left            =   -74280
            TabIndex        =   324
            Top             =   540
            Width           =   675
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Charm Resist"
            Height          =   255
            Index           =   30
            Left            =   3300
            TabIndex        =   322
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "BS Defense"
            Height          =   255
            Index           =   24
            Left            =   3480
            TabIndex        =   321
            Top             =   3180
            Width           =   1035
         End
         Begin VB.Label label 
            Alignment       =   2  'Center
            Caption         =   "DR"
            Height          =   195
            Index           =   19
            Left            =   4620
            TabIndex        =   320
            Top             =   840
            Width           =   735
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Index"
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   319
            Top             =   1620
            Width           =   435
         End
         Begin VB.Label lblMulti 
            Alignment       =   2  'Center
            Caption         =   "Multiplier"
            Height          =   195
            Left            =   2160
            TabIndex        =   318
            Top             =   840
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label lblBase 
            Alignment       =   2  'Center
            Caption         =   "Base"
            Height          =   195
            Left            =   1140
            TabIndex        =   317
            Top             =   840
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label16 
            Caption         =   "HH:MM:SS"
            Height          =   195
            Left            =   1140
            TabIndex        =   316
            Top             =   4260
            Width           =   1095
         End
         Begin VB.Label Label15 
            Caption         =   "MM/DD/YYYY"
            Height          =   195
            Left            =   2280
            TabIndex        =   315
            Top             =   4260
            Width           =   1095
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Last Killed"
            Height          =   195
            Left            =   240
            TabIndex        =   314
            Top             =   4380
            Width           =   795
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Active"
            Height          =   255
            Index           =   74
            Left            =   480
            TabIndex        =   313
            Top             =   2580
            Width           =   555
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Gender"
            Height          =   255
            Index           =   16
            Left            =   420
            TabIndex        =   312
            Top             =   3900
            Width           =   615
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Group"
            Height          =   255
            Index           =   2
            Left            =   540
            TabIndex        =   311
            Top             =   1920
            Width           =   495
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Experience"
            Height          =   255
            Index           =   4
            Left            =   180
            TabIndex        =   310
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Hitpoints"
            Height          =   255
            Index           =   5
            Left            =   3840
            TabIndex        =   309
            Top             =   1380
            Width           =   675
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "HP Regen"
            Height          =   255
            Index           =   6
            Left            =   3600
            TabIndex        =   308
            Top             =   1680
            Width           =   915
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "M.R."
            Height          =   255
            Index           =   7
            Left            =   4080
            TabIndex        =   307
            Top             =   1980
            Width           =   435
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Charm Lvl"
            Height          =   255
            Index           =   8
            Left            =   3780
            TabIndex        =   306
            Top             =   2580
            Width           =   735
         End
         Begin VB.Label label 
            Alignment       =   2  'Center
            Caption         =   "AC"
            Height          =   195
            Index           =   9
            Left            =   3840
            TabIndex        =   305
            Top             =   840
            Width           =   735
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Follow %"
            Height          =   255
            Index           =   11
            Left            =   3780
            TabIndex        =   304
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Regen Time"
            Height          =   255
            Index           =   12
            Left            =   60
            TabIndex        =   303
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Type"
            Height          =   255
            Index           =   13
            Left            =   600
            TabIndex        =   302
            Top             =   3180
            Width           =   435
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Energy"
            Height          =   255
            Index           =   14
            Left            =   3900
            TabIndex        =   301
            Top             =   3480
            Width           =   615
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Alignment"
            Height          =   255
            Index           =   15
            Left            =   300
            TabIndex        =   300
            Top             =   3540
            Width           =   735
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Game Limit"
            Height          =   255
            Index           =   54
            Left            =   180
            TabIndex        =   299
            Top             =   2280
            Width           =   855
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   5400
            Y1              =   780
            Y2              =   780
         End
         Begin VB.Label label 
            Caption         =   "Name"
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
            Index           =   1
            Left            =   240
            TabIndex        =   298
            Top             =   420
            Width           =   855
         End
         Begin VB.Label label 
            Caption         =   "Number"
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
            Index           =   0
            Left            =   3840
            TabIndex        =   297
            Top             =   420
            Width           =   855
         End
      End
      Begin exlimiter.EL EL1 
         Left            =   4800
         Top             =   60
         _ExtentX        =   1270
         _ExtentY        =   1270
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
      Width           =   2175
   End
   Begin MSComctlLib.ListView lvDatabase 
      Height          =   5835
      Left            =   60
      TabIndex        =   4
      Top             =   840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   10292
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
   Begin VB.Label lblNumberSearch 
      Caption         =   "#"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label17 
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
      Width           =   1875
   End
End
Attribute VB_Name = "frmMonster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim bNoRefresh As Boolean
Dim bLoaded As Boolean
Dim nCurrentRecord As Long


Private Sub cmdAttackSim_Click(Index As Integer)

If lvDatabase.ListItems.Count < 1 Then Exit Sub

If FormIsLoaded("frmMonsterAttackSim") Then
    Call frmMonsterAttackSim.RefreshMonsters
Else
    Load frmMonsterAttackSim
End If

frmMonsterAttackSim.GotoMonster (Val(lvDatabase.SelectedItem.Text))
frmMonsterAttackSim.Show
frmMonsterAttackSim.SetFocus

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
End Sub

Private Sub cmdFilterApply_Click()
Dim nStatus As Integer, bAdd As Boolean, x As Integer, bFiltered As Boolean
Dim z As Integer, bAbilMatch(2 To 4) As Boolean, nVal As Long, nVal2 As Double

On Error GoTo error:

If bLoaded Then Call saverecord(nCurrentRecord)

nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "BGETFIRST, Monster, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

Me.MousePointer = vbHourglass

bLoaded = False
lvDatabase.ListItems.clear

If chkFilterNone.Value = 1 Then
    Call LoadMonsters
    Call cmdFilter_Click
    cmdFilter.Caption = "Filter"
    GoTo out:
End If

Do While nStatus = 0
    bAdd = True
    MonsterRowToStruct Monsterdatabuf.buf
    
    If chkFilter(0).Value = 1 And bAdd Then 'group
        If Not Monsterrec.Group = cmbFilter(0).ListIndex Then bAdd = False
    End If
    If chkFilter(1).Value = 1 And bAdd Then 'align
        If Not Monsterrec.Alignment = cmbFilter(1).ListIndex Then bAdd = False
    End If
    If chkFilterIndex.Value = 1 And bAdd Then 'index
        If Not Monsterrec.Index >= Val(txtFilterIndex(0).Text) Or _
            Not Monsterrec.Index <= Val(txtFilterIndex(1).Text) Then bAdd = False
    End If

    If (chkFilter(2).Value = 1 Or chkFilter(3).Value = 1 Or chkFilter(4).Value = 1) And bAdd Then 'ability
        If chkFilter(2).Value = 1 Then bAbilMatch(2) = False Else bAbilMatch(2) = True
        If chkFilter(3).Value = 1 Then bAbilMatch(3) = False Else bAbilMatch(3) = True
        If chkFilter(4).Value = 1 Then bAbilMatch(4) = False Else bAbilMatch(4) = True
        
        For x = 0 To 9
            For z = 2 To 4
                If chkFilter(z).Value = 1 Then
                    If Monsterrec.AbilityA(x) = cmbFilter(z).ItemData(cmbFilter(z).ListIndex) Then
                        nVal = Val(txtFilterAbilityValue(z).Text)
                        bAbilMatch(z) = True
                        
                        If cmbFilterAbilityGL(z).ListIndex = 0 Then  'ANY
                        ElseIf cmbFilterAbilityGL(z).ListIndex = 1 Then  '<=
                            If Monsterrec.AbilityB(x) > nVal Then bAbilMatch(z) = False
                            If chkFilterExcludeZero.Value = 1 And Monsterrec.AbilityB(x) = 0 Then bAbilMatch(z) = False
                        ElseIf cmbFilterAbilityGL(z).ListIndex = 2 Then  '>=
                            If Monsterrec.AbilityB(x) < nVal Then bAbilMatch(z) = False
                        ElseIf cmbFilterAbilityGL(z).ListIndex = 3 Then  '=
                            If Not Monsterrec.AbilityB(x) = nVal Then bAbilMatch(z) = False
                        End If
                        
                        If Not bAbilMatch(z) Then GoTo abil_out:
                    End If
                End If
            Next z
        Next x
abil_out:
        If Not (bAbilMatch(2) And bAbilMatch(3) And bAbilMatch(4)) Then bAdd = False
    End If
    
    If chkFilter(5).Value = 1 And bAdd Then 'exp
        nVal = Val(txtFilterAbilityValue(5).Text)
        nVal2 = SLong2ULong(Monsterrec.Experience)
        If cmbFilterAbilityGL(5).ListIndex = 0 Then  '<=
            If nVal2 > nVal Then bAdd = False
            If chkFilterExcludeZero.Value = 1 And nVal2 = 0 Then bAdd = False
        ElseIf cmbFilterAbilityGL(5).ListIndex = 1 Then  '>=
            If nVal2 < nVal Then bAdd = False
        ElseIf cmbFilterAbilityGL(5).ListIndex = 2 Then  '=
            If Not nVal2 = nVal Then bAdd = False
        End If
    End If
    
    If chkFilter(6).Value = 1 And bAdd Then 'exp multi
        nVal = Val(txtFilterAbilityValue(6).Text)
        nVal2 = SLong2ULong(Monsterrec.ExpMulti)
        If cmbFilterAbilityGL(6).ListIndex = 0 Then  '<=
            If nVal2 > nVal Then bAdd = False
            If chkFilterExcludeZero.Value = 1 And nVal2 = 0 Then bAdd = False
        ElseIf cmbFilterAbilityGL(6).ListIndex = 1 Then  '>=
            If nVal2 < nVal Then bAdd = False
        ElseIf cmbFilterAbilityGL(6).ListIndex = 2 Then  '=
            If Not nVal2 = nVal Then bAdd = False
        End If
    End If
    
    If chkFilter(7).Value = 1 And bAdd Then 'message
        nVal = Val(txtFilterMessage.Text)
        If Monsterrec.MoveMsg = nVal Then GoTo msg_match:
        If Monsterrec.DeathMsg = nVal Then GoTo msg_match:
        For x = 0 To 4
            If Monsterrec.AttackHitMsg(x) = nVal Then GoTo msg_match:
            If Monsterrec.AttackMissMsg(x) = nVal Then GoTo msg_match:
            If Monsterrec.AttackDodgeMsg(x) = nVal Then GoTo msg_match:
        Next x
        bAdd = False
msg_match:
    End If
    If chkFilter(8).Value = 1 And bAdd Then 'textblock
        nVal = Val(txtFilterTB.Text)
        If Monsterrec.GreetTxt = nVal Then GoTo tb_match:
        If Monsterrec.TalkTxt = nVal Then GoTo tb_match:
        If Monsterrec.DescTxt = nVal Then GoTo tb_match:
        For x = 0 To 4
            If Monsterrec.AttackHitMsg(x) = nVal Then GoTo tb_match:
            If Monsterrec.AttackMissMsg(x) = nVal Then GoTo tb_match:
            If Monsterrec.AttackDodgeMsg(x) = nVal Then GoTo tb_match:
        Next x
        bAdd = False
tb_match:
    End If
    If chkFilter(9).Value = 1 And bAdd Then 'item
        nVal = Val(txtFilterItem.Text)
        If Monsterrec.WeaponNumber = nVal Then GoTo item_match:
        For x = 0 To 9
            If Monsterrec.ItemNumber(x) = nVal Then GoTo item_match:
        Next x
        bAdd = False
item_match:
    End If
    If chkFilter(10).Value = 1 And bAdd Then 'spell
        nVal = Val(txtFilterSpell.Text)
        If Monsterrec.CreateSpellNumber = nVal Then GoTo spell_match:
        If Monsterrec.DeathSpellNumber = nVal Then GoTo spell_match:
        For x = 0 To 4
            If Monsterrec.AttackHitSpell(x) = nVal Then GoTo spell_match:
            If Monsterrec.AttackType(x) = 2 Then 'spell
                If Monsterrec.AttackAccuSpell(x) = nVal Then GoTo spell_match:
            End If
            If Monsterrec.SpellNumber(x) = nVal Then GoTo spell_match:
        Next x
        For x = 0 To 9
            If Monsterrec.ItemNumber(x) = nVal Then GoTo spell_match:
        Next x
        bAdd = False
spell_match:
    End If
    
    If chkFilter(11).Value = 1 And bAdd Then 'regen > 0
        If Not Monsterrec.RegenTime > 0 Then bAdd = False
    End If
    If chkFilter(12).Value = 1 And bAdd Then 'limit > 0
        If Not Monsterrec.GameLimit > 0 Then bAdd = False
    End If
    If chkFilter(13).Value = 1 And bAdd Then 'undead
        If Not Monsterrec.Undead = 1 Then bAdd = False
    End If
    
    If bAdd Then
        Call AddMonsterToLV(Monsterrec.Number)
    Else
        bFiltered = True
    End If
    
    nStatus = BTRCALL(BGETNEXT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
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

For x = 0 To 13
    chkFilter(x).Enabled = bAction
    If x <= 4 Then cmbFilter(x).Enabled = bAction
Next x

For x = 2 To 6
    cmbFilterAbilityGL(x).Enabled = bAction
    txtFilterAbilityValue(x).Enabled = bAction
    If x = 6 And eDatFileVersion >= v111j Then 'exp multi
        cmbFilterAbilityGL(x).Enabled = bAction
        txtFilterAbilityValue(x).Enabled = bAction
    End If
Next x

txtFilterIndex(0).Enabled = bAction
txtFilterIndex(1).Enabled = bAction
chkFilterIndex.Enabled = bAction

txtFilterMessage.Enabled = bAction
txtFilterTB.Enabled = bAction
txtFilterItem.Enabled = bAction
txtFilterSpell.Enabled = bAction

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

For x = 0 To 13
    chkFilter(x).Value = 0
    If x <= 4 Then cmbFilter(x).ListIndex = 0
Next x
For x = 2 To 6
    cmbFilterAbilityGL(x).ListIndex = 0
    txtFilterAbilityValue(x).Text = 0
Next x

txtFilterIndex(0).Text = 0
txtFilterIndex(1).Text = 0
chkFilterIndex.Value = 0

txtFilterMessage.Text = 0
txtFilterTB.Text = 0
txtFilterItem.Text = 0
txtFilterSpell.Text = 0

End Sub

Private Sub txtFilterAbilityValue_GotFocus(Index As Integer)
Call SelectAll(txtFilterAbilityValue(Index))
End Sub

Private Sub txtFilterAbilityValue_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtFilterIndex_GotFocus(Index As Integer)
Call SelectAll(txtFilterIndex(Index))
End Sub

Private Sub txtFilterIndex_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtFilterItem_GotFocus()
Call SelectAll(txtFilterItem)
End Sub

Private Sub txtFilterItem_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtFilterMessage_GotFocus()
Call SelectAll(txtFilterMessage)
End Sub

Private Sub txtFilterMessage_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtFilterSpell_GotFocus()
Call SelectAll(txtFilterSpell)
End Sub

Private Sub txtFilterSpell_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtFilterTB_GotFocus()
Call SelectAll(txtFilterTB)
End Sub

Private Sub txtFilterTB_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
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

Private Sub cmdAbilityLookup_Click(Index As Integer)
    Call LookupAbility(Val(txtAbilityA(Index).Text), Val(txtAbilityB(Index).Text))
End Sub


Private Sub cmdAttackClear_Click()
Dim x As Integer
On Error GoTo error:

For x = 0 To 4
    txtAttackAccuSpell(x).Text = 0
    txtAttackPer(x).Text = 0
    txtAttackMinHCastPer(x).Text = 0
    txtAttackMaxHCastLvL(x).Text = 0
    txtAttackHitMsg(x).Text = 0
    txtAttackDodgeMsg(x).Text = 0
    txtAttackMissMsg(x).Text = 0
    txtAttackEnergy(x).Text = 0
    txtAttackHitSpell(x).Text = 0
    cmbAttackType(x).ListIndex = 0
Next x

Exit Sub
error:
Call HandleError("cmdAttackClear_Click")
End Sub

Private Sub cmdAttackCopyAll_Click(Index As Integer)
Dim x As Integer, y As Integer
On Error GoTo error:

If Index = 0 Then
    For x = 0 To 4
        y = x * 10
        nMonsterAllAttackCopy(y) = txtAttackAccuSpell(x).Text
        nMonsterAllAttackCopy(y + 1) = txtAttackPer(x).Text
        nMonsterAllAttackCopy(y + 2) = txtAttackMinHCastPer(x).Text
        nMonsterAllAttackCopy(y + 3) = txtAttackMaxHCastLvL(x).Text
        nMonsterAllAttackCopy(y + 4) = txtAttackHitMsg(x).Text
        nMonsterAllAttackCopy(y + 5) = txtAttackDodgeMsg(x).Text
        nMonsterAllAttackCopy(y + 6) = txtAttackMissMsg(x).Text
        nMonsterAllAttackCopy(y + 7) = txtAttackEnergy(x).Text
        nMonsterAllAttackCopy(y + 8) = txtAttackHitSpell(x).Text
        nMonsterAllAttackCopy(y + 9) = cmbAttackType(x).ListIndex
    Next
Else
    For x = 0 To 4
        y = x * 10
        txtAttackAccuSpell(x).Text = nMonsterAllAttackCopy(y)
        txtAttackPer(x).Text = nMonsterAllAttackCopy(y + 1)
        txtAttackMinHCastPer(x).Text = nMonsterAllAttackCopy(y + 2)
        txtAttackMaxHCastLvL(x).Text = nMonsterAllAttackCopy(y + 3)
        txtAttackHitMsg(x).Text = nMonsterAllAttackCopy(y + 4)
        txtAttackDodgeMsg(x).Text = nMonsterAllAttackCopy(y + 5)
        txtAttackMissMsg(x).Text = nMonsterAllAttackCopy(y + 6)
        txtAttackEnergy(x).Text = nMonsterAllAttackCopy(y + 7)
        txtAttackHitSpell(x).Text = nMonsterAllAttackCopy(y + 8)
        cmbAttackType(x).ListIndex = nMonsterAllAttackCopy(y + 9)
    Next
End If

Exit Sub
error:
Call HandleError("cmdAttackCopyAll_Click")
End Sub

Private Sub cmdAttackCopySingle_Click(Index As Integer)
On Error GoTo error:

If Index >= 0 And Index <= 4 Then
    nMonsterSingleAttackCopy(0) = txtAttackAccuSpell(Index).Text
    nMonsterSingleAttackCopy(1) = txtAttackPer(Index).Text
    nMonsterSingleAttackCopy(2) = txtAttackMinHCastPer(Index).Text
    nMonsterSingleAttackCopy(3) = txtAttackMaxHCastLvL(Index).Text
    nMonsterSingleAttackCopy(4) = txtAttackHitMsg(Index).Text
    nMonsterSingleAttackCopy(5) = txtAttackDodgeMsg(Index).Text
    nMonsterSingleAttackCopy(6) = txtAttackMissMsg(Index).Text
    nMonsterSingleAttackCopy(7) = txtAttackEnergy(Index).Text
    nMonsterSingleAttackCopy(8) = txtAttackHitSpell(Index).Text
    nMonsterSingleAttackCopy(9) = cmbAttackType(Index).ListIndex
Else
    txtAttackAccuSpell(Index - 5).Text = nMonsterSingleAttackCopy(0)
    txtAttackPer(Index - 5).Text = nMonsterSingleAttackCopy(1)
    txtAttackMinHCastPer(Index - 5).Text = nMonsterSingleAttackCopy(2)
    txtAttackMaxHCastLvL(Index - 5).Text = nMonsterSingleAttackCopy(3)
    txtAttackHitMsg(Index - 5).Text = nMonsterSingleAttackCopy(4)
    txtAttackDodgeMsg(Index - 5).Text = nMonsterSingleAttackCopy(5)
    txtAttackMissMsg(Index - 5).Text = nMonsterSingleAttackCopy(6)
    txtAttackEnergy(Index - 5).Text = nMonsterSingleAttackCopy(7)
    txtAttackHitSpell(Index - 5).Text = nMonsterSingleAttackCopy(8)
    cmbAttackType(Index - 5).ListIndex = nMonsterSingleAttackCopy(9)
End If

Exit Sub
error:
Call HandleError("cmdAttackCopySingle_Click")
End Sub

Private Sub cmdResetKill_Click()
txtTimeKilled.Text = "00:00:00"
txtDateKilled.Text = "00/00/0000"
End Sub

Private Sub Form_Load()
Dim sCaption As String, j As Integer
On Error Resume Next

sCaption = frmMain.Caption
frmMain.Caption = sCaption & " - Loading Monsters ..."
DoEvents

With EL1
    .FormInQuestion = Me
    .MINHEIGHT = 475 + (TITLEBAR_OFFSET / 10)
    .MINWIDTH = 590
    .CenterOnLoad = False
    .EnableLimiter = True
End With

Me.Top = ReadINI("Windows", "MonsterTop")
Me.Left = ReadINI("Windows", "MonsterLeft")
Me.Width = ReadINI("Windows", "MonsterWidth")
Me.Height = ReadINI("Windows", "MonsterHeight")

lvDatabase.ListItems.clear

If eDatFileVersion >= v111j Then
    lblBase.Visible = True
    lblMulti.Visible = True
    txtBase.Visible = True
    txtMulti.Visible = True
    txtExperience.Locked = True
    txtExperience.BackColor = &H8000000F
Else
    lblBase.Visible = False
    lblMulti.Visible = False
    txtBase.Visible = False
    txtMulti.Visible = False
    txtExperience.Locked = False
    txtExperience.BackColor = &H80000005
End If

Call LoadAbilities

For j = 0 To 4
    cmbFilter(j).ListIndex = 0
    Call AutoSizeDropDownWidth(cmbFilter(j))
    Call ExpandCombo(cmbFilter(j), HeightOnly, TripleWidth, fraFilter2.hwnd)
Next j

For j = 2 To 6
    cmbFilterAbilityGL(j).ListIndex = 0
Next j

bLoaded = False
Call LoadMonsters

Me.Show
Me.SetFocus
txtSearch.SetFocus
If ReadINI("Windows", "MonsterMaxed") = "1" Then Me.WindowState = vbMaximized
frmMain.Caption = sCaption

End Sub

Private Sub LoadAbilities()
Dim x As Integer
On Error GoTo error:

For x = 2 To 4
    cmbFilter(x).clear
Next x
rsAbilities.MoveFirst
Do Until rsAbilities.EOF
    If Not rsAbilities.Fields("Number") = 0 Then
        For x = 2 To 4
            cmbFilter(x).AddItem rsAbilities.Fields("Name") & " (" & rsAbilities.Fields("Number") & ")"
            cmbFilter(x).ItemData(cmbFilter(x).NewIndex) = rsAbilities.Fields("Number")
        Next x
    End If
    rsAbilities.MoveNext
Loop

For x = 2 To 4
    cmbFilter(x).AddItem "None (0)", 0
    cmbFilter(x).ListIndex = 0
Next x

out:
Exit Sub
error:
Call HandleError("LoadAbilities")
Resume out:
End Sub

Private Sub cmbAttackType_Click(Index As Integer)

Call RefreshAttackWindow(Index)

End Sub

Private Sub cmdDiscard_Click()
Dim nStatus As Integer

If lvDatabase.SelectedItem Is Nothing Or nCurrentRecord = 0 Then
    MsgBox "No current record."
    Exit Sub
End If

nStatus = BTRCALL(BGETEQUAL, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), nCurrentRecord, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error on BGETEQUAL: " & BtrieveErrorCode(nStatus)
Else
    DispMonsterInfo Monsterdatabuf.buf
End If
End Sub

Private Sub cmdEditAttackSpell_Click(Index As Integer)
Call frmSpell.GotoSpell(Val(txtAttackAccuSpell(Index).Text))
frmSpell.Show
frmSpell.SetFocus
End Sub

Private Sub cmdEditCreateSpell_Click()
Call frmSpell.GotoSpell(Val(txtCreateSpellNumber.Text))
frmSpell.Show
frmSpell.SetFocus
End Sub

Private Sub cmdEditDeathMsg_Click()
Call frmMessage.GotoMSG(Val(txtDeathMsg.Text))
frmMessage.Show
frmMessage.SetFocus
End Sub

Private Sub cmdEditDeathSpell_Click()
Call frmSpell.GotoSpell(Val(txtDeathSpellNumber.Text))
frmSpell.Show
frmSpell.SetFocus
End Sub

Private Sub cmdEditDescText_Click()
Call frmTextblock.GotoTB(Val(txtDescTxt.Text))
frmTextblock.Show
frmTextblock.SetFocus
End Sub

Private Sub cmdEditDodgeMsg_Click(Index As Integer)
Call frmMessage.GotoMSG(Val(txtAttackDodgeMsg(Index).Text))
frmMessage.Show
frmMessage.SetFocus
End Sub

Private Sub cmdEditGreetTxt_Click()
Call frmTextblock.GotoTB(Val(txtGreetTxt.Text))
frmTextblock.Show
frmTextblock.SetFocus
End Sub

Private Sub cmdEditHitMsg_Click(Index As Integer)
Call frmMessage.GotoMSG(Val(txtAttackHitMsg(Index).Text))
frmMessage.Show
frmMessage.SetFocus
End Sub

Private Sub cmdEditHitSpell_Click(Index As Integer)
Call frmSpell.GotoSpell(Val(txtAttackHitSpell(Index).Text))
frmSpell.Show
frmSpell.SetFocus
End Sub

Private Sub cmdEditItemDrop_Click(Index As Integer)
Call frmItem.GotoItem(Val(txtItemNumber(Index).Text))
frmItem.Show
frmItem.SetFocus
End Sub

Private Sub cmdEditMissMsg_Click(Index As Integer)
Call frmMessage.GotoMSG(Val(txtAttackMissMsg(Index).Text))
frmMessage.Show
frmMessage.SetFocus
End Sub

Private Sub cmdEditMoveMsg_Click()
Call frmMessage.GotoMSG(Val(txtMoveMsg.Text))
frmMessage.Show
frmMessage.SetFocus
End Sub

Private Sub cmdEditSpell_Click(Index As Integer)
Call frmSpell.GotoSpell(Val(txtSpellNumber(Index).Text))
frmSpell.Show
frmSpell.SetFocus
End Sub

Private Sub cmdEditTalkText_Click()
Call frmTextblock.GotoTB(Val(txtTalkTxt.Text))
frmTextblock.Show
frmTextblock.SetFocus
End Sub

Private Sub cmdEditWeapon_Click()
Call frmItem.GotoItem(Val(txtWeaponNumber.Text))
frmItem.Show
frmItem.SetFocus
End Sub

Private Sub cmdSave_Click()
On Error GoTo error:

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
If lvDatabase.SelectedItem Is Nothing Then Exit Sub

Call saverecord(nCurrentRecord)
'Call lvDatabase_ItemClick(lvDatabase.SelectedItem)

Dim oLI As ListItem
Set oLI = lvDatabase.FindItem(Monsterrec.Number, lvwText, , 0)
If Not oLI Is Nothing Then
    oLI.ListSubItems(1).Text = ClipNull(Monsterrec.Name)
    If Not bOnlyNames Then
        If eDatFileVersion >= v111j Then
            oLI.ListSubItems(2).Text = (Monsterrec.Experience * Monsterrec.ExpMulti)
        Else
            oLI.ListSubItems(2).Text = Monsterrec.Experience
        End If
        oLI.ListSubItems(3).Text = GetMonGroupName(Monsterrec.Group)
        oLI.ListSubItems(4).Text = Monsterrec.RegenTime
        oLI.ListSubItems(5).Text = Monsterrec.GameLimit
    End If
End If
Set oLI = Nothing

out:
Exit Sub
error:
Call HandleError("cmdSave_Click")
Resume out:

End Sub
Public Sub GotoMonster(ByVal nRecnum As Long)
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

Private Sub lvDatabase_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim nSort As ListDataType
Select Case ColumnHeader.Index
    Case 1, 3, 5, 6: nSort = ldtNumber
    Case Else: nSort = ldtString
End Select
SortListView lvDatabase, ColumnHeader.Index, nSort, lvDatabase.SortOrder
End Sub

Private Sub lvDatabase_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim temp As Long, nStatus As Integer

If bLoaded = True And chkAutoSave.Value = 1 Then Call saverecord(nCurrentRecord)

temp = Val(Item.Text)
nStatus = BTRCALL(BGETEQUAL, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), temp, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error on BGETEQUAL: " & BtrieveErrorCode(nStatus)
    bLoaded = False
Else
    nCurrentRecord = temp
    DispMonsterInfo Monsterdatabuf.buf
    bLoaded = True
End If
End Sub

Private Sub txtAbilityA_Change(Index As Integer)

If Me.ActiveControl Is lblName(Index) Then Exit Sub
Call FindAbilityNumber(txtAbilityA(Index), lblName(Index))

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

Private Sub LoadMonsters()
Dim nStatus As Integer

lvDatabase.ColumnHeaders.clear
lvDatabase.ColumnHeaders.add 1, "Number", "#", 600, lvwColumnLeft
lvDatabase.ColumnHeaders.add 2, "Name", "Name", 1900, lvwColumnCenter
If Not bOnlyNames Then
    lvDatabase.ColumnHeaders.add 3, "Exp", "Exp", 1000, lvwColumnCenter
    lvDatabase.ColumnHeaders.add 4, "Group", "Group", 1100, lvwColumnCenter
    lvDatabase.ColumnHeaders.add 5, "RGN", "RGN", 700, lvwColumnCenter
    lvDatabase.ColumnHeaders.add 6, "Limit", "Limit", 700, lvwColumnCenter
End If

nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "LoadMonsters, BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

Do While nStatus = 0
    MonsterRowToStruct Monsterdatabuf.buf
    
    Call AddMonsterToLV(Monsterrec.Number)
    
    nStatus = BTRCALL(BGETNEXT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "LoadMonsters, Error: " & BtrieveErrorCode(nStatus)
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
Private Sub AddMonsterToLV(ByVal nNumber As Integer)
Dim nStatus As Integer, oLI As ListItem
On Error GoTo error:

If Not nNumber = Monsterrec.Number Then
    nStatus = BTRCALL(BGETEQUAL, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), nNumber, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then MsgBox "Error getting record " & nNumber & ": " & BtrieveErrorCode(nStatus)
    bLoaded = False
    Exit Sub
End If

Set oLI = lvDatabase.ListItems.add()
oLI.Text = Monsterrec.Number

oLI.ListSubItems.add (1), "Name", ClipNull(Monsterrec.Name)

If Not bOnlyNames Then
    If eDatFileVersion >= v111j Then
        oLI.ListSubItems.add (2), "Exp", (Monsterrec.Experience * Monsterrec.ExpMulti)
    Else
        oLI.ListSubItems.add (2), "Exp", Monsterrec.Experience
    End If
    
    oLI.ListSubItems.add (3), "Group", GetMonGroupName(Monsterrec.Group)
    oLI.ListSubItems.add (4), "RGN", Monsterrec.RegenTime
    oLI.ListSubItems.add (5), "Limit", Monsterrec.GameLimit
End If

Set oLI = Nothing
Exit Sub
error:
Call HandleError
Set oLI = Nothing
End Sub

Private Sub DispMonsterInfo(row() As Byte)
On Error GoTo error:
Dim x As Integer

Call MonsterRowToStruct(row())

bNoRefresh = True

Me.Caption = "Monster Editor -- " & ClipNull(Monsterrec.Name)

txtNumber.Text = Monsterrec.Number
txtName.Text = Monsterrec.Name

If eDatFileVersion >= v111j Then
    txtBase.Text = SLong2ULong(Monsterrec.Experience)
    txtMulti.Text = SLong2ULong(Monsterrec.ExpMulti)
    txtExperience.Text = CDbl(SLong2ULong(Monsterrec.Experience)) * CDbl(SLong2ULong(Monsterrec.ExpMulti))
Else
    txtExperience.Text = SLong2ULong(Monsterrec.Experience)
End If

txtCharmRes.Text = Monsterrec.CharmRes
txtBSDefense.Text = Monsterrec.BSDefence
txtIndex.Text = Monsterrec.Index
txtWeaponNumber.Text = Monsterrec.WeaponNumber
'txtWeaponName.Text = GetItemName(Monsterrec.WeaponNumber)
txtAC.Text = Monsterrec.AC
txtDR.Text = Monsterrec.DR
txtFollow.Text = Monsterrec.Follow
txtMR.Text = Monsterrec.MR
txtHitPoints.Text = Monsterrec.Hitpoints
txtEnergy.Text = Monsterrec.Energy
txtHpRegen.Text = Monsterrec.HPRegen
txtGameLimit.Text = Monsterrec.GameLimit
txtActive.Text = Monsterrec.Active
cmbType.ListIndex = Monsterrec.Type
txtAlignment.ListIndex = Monsterrec.Alignment
txtGender.ListIndex = Monsterrec.Gender
cmbGroup.ListIndex = Monsterrec.Group
txtRegenTime.Text = Monsterrec.RegenTime
txtMoveMsg.Text = Monsterrec.MoveMsg
'txtMoveMsgDisplay.Text = GetMessages(Monsterrec.MoveMsg, 1)
txtDeathMsg.Text = Monsterrec.DeathMsg
'txtDeathMsgDisplay.Text = GetMessages(Monsterrec.DeathMsg, 3)
txtRunic.Text = SLong2ULong(Monsterrec.Runic)
txtPlatinum.Text = SLong2ULong(Monsterrec.Platinum)
txtGold.Text = SLong2ULong(Monsterrec.Gold)
txtSilver.Text = SLong2ULong(Monsterrec.Silver)
txtCopper.Text = SLong2ULong(Monsterrec.Copper)
txtDesc(0).Text = Monsterrec.DescLine1
txtDesc(1).Text = Monsterrec.DescLine2
txtDesc(2).Text = Monsterrec.DescLine3
txtDesc(3).Text = Monsterrec.DescLine4
txtGreetTxt.Text = Monsterrec.GreetTxt
'txtGreetTxtDisplay.Text = GetTextblock(Monsterrec.GreetTxt)
txtCharmlvl.Text = Monsterrec.CharmLvL
txtDescTxt.Text = Monsterrec.DescTxt
'txtDescTxtDisplay.Text = GetTextblock(Monsterrec.DescTxt)
txtTalkTxt.Text = Monsterrec.TalkTxt
'txtTalkTxtDisplay.Text = GetTextblock(Monsterrec.TalkTxt)
txtDeathSpellNumber.Text = Monsterrec.DeathSpellNumber
'txtDeathSpellName.Text = GetSpellName(Monsterrec.DeathSpellNumber)
txtCreateSpellNumber.Text = Monsterrec.CreateSpellNumber
'txtCreateSpellName.Text = GetSpellName(Monsterrec.CreateSpellNumber)
txtDateKilled.Text = DOSDate2Date(SInt2UInt(Monsterrec.DateKilled))
txtTimeKilled.Text = DOSTime2Time(SInt2UInt(Monsterrec.TimeKilled))
If Monsterrec.Undead > 1 Then Monsterrec.Undead = 0
chkUndead.Value = Monsterrec.Undead

For x = 0 To 4
    txtAttackAccuSpell(x).Text = Monsterrec.AttackAccuSpell(x)
    txtAttackPer(x).Text = Monsterrec.AttackPer(x)
    txtAttackMinHCastPer(x).Text = Monsterrec.AttackMinHCastPer(x)
    txtAttackMaxHCastLvL(x).Text = Monsterrec.AttackMaxHCastLvl(x)
    txtAttackHitMsg(x).Text = Monsterrec.AttackHitMsg(x)
    txtAttackDodgeMsg(x).Text = Monsterrec.AttackDodgeMsg(x)
    txtAttackMissMsg(x).Text = Monsterrec.AttackMissMsg(x)
    txtAttackEnergy(x).Text = Monsterrec.AttackEnergy(x)
    txtAttackHitSpell(x).Text = Monsterrec.AttackHitSpell(x)
    txtSpellNumber(x).Text = Monsterrec.SpellNumber(x)
    txtSpellCastPer(x).Text = Monsterrec.SpellCastPer(x)
    txtSpellCastLvL(x).Text = Monsterrec.SpellCastLvl(x)
    
    If Monsterrec.AttackType(x) > 3 Then Monsterrec.AttackType(x) = 0
    cmbAttackType(x).ListIndex = Monsterrec.AttackType(x)
    
    'txtAttackHitMsgDisplay(x) = GetMessages(Monsterrec.AttackHitMsg(x), 1)
    'txtAttackDodgeMsgDisplay(x).Text = GetMessages(Monsterrec.AttackDodgeMsg(x), 3)
    'txtAttackMissMsgDisplay(x).Text = GetMessages(Monsterrec.AttackMissMsg(x), 2)
    'txtSpellName(x).Text = GetSpellName(Monsterrec.SpellNumber(x))
    'txtAttackHitSpellName(x).Text = GetSpellName(Monsterrec.AttackHitSpell(x))
Next

For x = 0 To 9
    txtItemNumber(x).Text = Monsterrec.ItemNumber(x)
    'txtItemName(x).Text = GetItemName(Monsterrec.ItemNumber(x))
    txtAbilityA(x).Text = Monsterrec.AbilityA(x)
    txtAbilityB(x).Text = Monsterrec.AbilityB(x)
    txtItemUses(x).Text = Monsterrec.ItemUses(x)
    txtItemDropPer(x).Text = Monsterrec.ItemDropPer(x)
Next

out:
bNoRefresh = False
Call RefreshAttackWindow

Exit Sub
error:
Call HandleError("DispMonsterInfo")
MsgBox "Warning, record was not completely displayed." & vbCrLf _
    & "Previous records stats may still be in memory.  Select 'Disable DB Writing'" & vbCrLf _
    & "from the file menu and then reload the editor.", vbExclamation
Resume out:
End Sub

Private Sub saverecord(ByVal nRecord As Long)
On Error GoTo error:
Dim nStatus As Integer, x As Integer, temp As Long

If nRecord = 0 Then Exit Sub

nStatus = BTRCALL(BGETEQUAL, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), nRecord, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error on BGETEQUAL: " & BtrieveErrorCode(nStatus)
    Exit Sub
Else
    MonsterRowToStruct Monsterdatabuf.buf
End If

'DoEvents
Monsterrec.Name = RTrim(txtName.Text) & Chr(0)

If eDatFileVersion >= v111j Then
    If Val(txtBase.Text) * Val(txtMulti.Text) > 2147483646 Then
        MsgBox "Total Experience cannot be greater than 2,147,483,646 ... setting it to that.", vbExclamation
        Monsterrec.Experience = 65538
        Monsterrec.ExpMulti = 32767
    Else
        Monsterrec.Experience = ULong2SLong(Val(txtBase.Text))
        Monsterrec.ExpMulti = ULong2SLong(Val(txtMulti.Text))
    End If
Else
    Monsterrec.Experience = ULong2SLong(Val(txtExperience.Text))
End If

Monsterrec.CharmRes = Val(txtCharmRes.Text)
Monsterrec.BSDefence = Val(txtBSDefense.Text)
Monsterrec.Index = Val(txtIndex.Text)
Monsterrec.WeaponNumber = Val(txtWeaponNumber.Text)
Monsterrec.AC = Val(txtAC.Text)
Monsterrec.DR = Val(txtDR.Text)
Monsterrec.Follow = Val(txtFollow.Text)
Monsterrec.MR = Val(txtMR.Text)
Monsterrec.Hitpoints = Val(txtHitPoints.Text)
Monsterrec.Energy = Val(txtEnergy.Text)
Monsterrec.HPRegen = Val(txtHpRegen.Text)
Monsterrec.GameLimit = Val(txtGameLimit.Text)
Monsterrec.Active = Val(txtActive.Text)
Monsterrec.Type = cmbType.ListIndex
Monsterrec.Alignment = txtAlignment.ListIndex
Monsterrec.Gender = txtGender.ListIndex
Monsterrec.Group = cmbGroup.ListIndex
Monsterrec.RegenTime = Val(txtRegenTime.Text)
Monsterrec.MoveMsg = Val(txtMoveMsg.Text)
Monsterrec.DeathMsg = Val(txtDeathMsg.Text)
Monsterrec.Runic = ULong2SLong(Val(txtRunic.Text))
Monsterrec.Platinum = ULong2SLong(Val(txtPlatinum.Text))
Monsterrec.Gold = ULong2SLong(Val(txtGold.Text))
Monsterrec.Silver = ULong2SLong(Val(txtSilver.Text))
Monsterrec.Copper = ULong2SLong(Val(txtCopper.Text))
Monsterrec.DescLine1 = Trim(txtDesc(0).Text) & Chr(0)
Monsterrec.DescLine2 = Trim(txtDesc(1).Text) & Chr(0)
Monsterrec.DescLine3 = Trim(txtDesc(2).Text) & Chr(0)
Monsterrec.DescLine4 = Trim(txtDesc(3).Text) & Chr(0)
Monsterrec.GreetTxt = Val(txtGreetTxt.Text)
Monsterrec.CharmLvL = Val(txtCharmlvl.Text)
Monsterrec.DescTxt = Val(txtDescTxt.Text)
Monsterrec.TalkTxt = Val(txtTalkTxt.Text)
Monsterrec.DeathSpellNumber = Val(txtDeathSpellNumber.Text)
Monsterrec.CreateSpellNumber = Val(txtCreateSpellNumber.Text)
Monsterrec.Undead = chkUndead.Value

temp = Time2DOSTime(txtTimeKilled.Text)
If temp <> -1 Then
    Monsterrec.TimeKilled = UInt2SInt(temp)
End If

temp = Date2DOSDate(txtDateKilled.Text)
If temp <> -1 Then
    Monsterrec.DateKilled = UInt2SInt(temp)
End If

For x = 0 To 4
    Monsterrec.AttackAccuSpell(x) = Val(txtAttackAccuSpell(x).Text)
    Monsterrec.AttackPer(x) = Val(txtAttackPer(x).Text)
    Monsterrec.AttackMinHCastPer(x) = Val(txtAttackMinHCastPer(x).Text)
    Monsterrec.AttackMaxHCastLvl(x) = Val(txtAttackMaxHCastLvL(x).Text)
    Monsterrec.AttackHitMsg(x) = Val(txtAttackHitMsg(x).Text)
    Monsterrec.AttackDodgeMsg(x) = Val(txtAttackDodgeMsg(x).Text)
    Monsterrec.AttackMissMsg(x) = Val(txtAttackMissMsg(x).Text)
    Monsterrec.AttackEnergy(x) = Val(txtAttackEnergy(x).Text)
    Monsterrec.AttackHitSpell(x) = Val(txtAttackHitSpell(x).Text)
    Monsterrec.SpellNumber(x) = Val(txtSpellNumber(x).Text)
    Monsterrec.SpellCastPer(x) = Val(txtSpellCastPer(x).Text)
    Monsterrec.SpellCastLvl(x) = Val(txtSpellCastLvL(x).Text)
    Monsterrec.AttackType(x) = cmbAttackType(x).ListIndex
Next

For x = 0 To 9
    Monsterrec.ItemNumber(x) = Val(txtItemNumber(x).Text)
    Monsterrec.AbilityA(x) = Val(txtAbilityA(x).Text)
    Monsterrec.AbilityB(x) = Val(txtAbilityB(x).Text)
    Monsterrec.ItemUses(x) = Val(txtItemUses(x).Text)
    Monsterrec.ItemDropPer(x) = Val(txtItemDropPer(x).Text)
Next

nStatus = UpdateMonster
If Not nStatus = 0 Then
    MsgBox "SaveRecord, Error: " & BtrieveErrorCode(nStatus)
Else
    DispMonsterInfo Monsterdatabuf.buf
End If

Exit Sub
error:
Call HandleError
End Sub


Private Sub Form_Unload(Cancel As Integer)
'Set TTtxtBox = Nothing
If bLoaded = True Then Call saverecord(nCurrentRecord)
If Me.WindowState = vbMinimized Then Exit Sub

If Me.WindowState = vbMaximized Then
    Call WriteINI("Windows", "MonsterMaxed", 1)
Else
    Call WriteINI("Windows", "MonsterMaxed", 0)
    Call WriteINI("Windows", "MonsterTop", Me.Top)
    Call WriteINI("Windows", "MonsterLeft", Me.Left)
    Call WriteINI("Windows", "MonsterWidth", Me.Width)
    Call WriteINI("Windows", "MonsterHeight", Me.Height)
End If
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

nStatus = BTRCALL(BGETEQUAL, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), nCurrentRecord, KEY_BUF_LEN, 0)
If nStatus = 0 Then
    nStatus = BTRCALL(BDELETE, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
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
Dim nNewMonsterNumber As String, oLI As ListItem

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

If bLoaded = True Then Call saverecord(nCurrentRecord)

nNewMonsterNumber = InputBox("New Monster Number:" & vbCrLf & vbCrLf & "Enter 0 for the next highest number.", "Insert", "0")
If nNewMonsterNumber = "" Then Exit Sub

Monsterrec.Number = Val(nNewMonsterNumber)
'Monsterrec.Name = "New Monster" & Chr(0)
Call MonsterStructToRow(Monsterdatabuf.buf)

nStatus = BTRCALL(BINSERT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdInsert, BINSERT, Error: " & BtrieveErrorCode(nStatus)
    bLoaded = False
Else
    MonsterRowToStruct Monsterdatabuf.buf
    
    Call AddMonsterToLV(Monsterrec.Number)
    
    nCurrentRecord = Monsterrec.Number
    DispMonsterInfo Monsterdatabuf.buf
    
    SortListView lvDatabase, 1, ldtNumber, True
    
    Set oLI = lvDatabase.FindItem(Monsterrec.Number, lvwText, , 0)
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

Private Sub txtAbilityB_GotFocus(Index As Integer)
Call SelectAll(txtAbilityB(Index))

End Sub

Private Sub txtAC_GotFocus()
Call SelectAll(txtAC)

End Sub

Private Sub txtActive_GotFocus()
Call SelectAll(txtActive)

End Sub

Private Sub txtAttackAccuSpell_Change(Index As Integer)
On Error GoTo error:

Call RefreshAttackWindow(Index)

out:
Exit Sub
error:
Call HandleError("txtAttackAccuSpell_Change")
Resume out:
End Sub

Private Sub RefreshAttackWindow(Optional ByVal nSingleAttackNumber As Integer = -1)
Dim x As Integer, y1 As Integer, y2 As Integer
On Error GoTo error:

If bNoRefresh Then Exit Sub

If nSingleAttackNumber >= 0 Then
    y2 = nSingleAttackNumber
    y1 = nSingleAttackNumber
Else
    y1 = 0
    y2 = 4
End If

For x = y1 To y2
    Select Case cmbAttackType(x).ListIndex
        Case 1: 'Normal
            lblAttackAccuSpell(x).Caption = "Accuracy"
            lblAttackMinHCastPer(x).Caption = "Min Hit"
            lblAttackMaxHCastLvl(x).Caption = "Max Hit"
            txtAttackAccuSpellName(x).Text = "N/A"
            txtAttackSpellDamage(x).Text = ""
            lblAttackSpellRange(x).Visible = False
            txtAttackSpellDamage(x).Visible = False
            
        Case 2: 'Spell
            lblAttackAccuSpell(x).Caption = "Spell"
            lblAttackMinHCastPer(x).Caption = "Cast %"
            lblAttackMaxHCastLvl(x).Caption = "Cast LVL"
            If Val(txtAttackAccuSpell(x).Text) > 0 Then
                txtAttackAccuSpellName(x).Text = GetSpellName(Val(txtAttackAccuSpell(x).Text))
                txtAttackSpellDamage(x).Text = GetSpellRange(Val(txtAttackAccuSpell(x).Text), _
                    Val(txtAttackMaxHCastLvL(x).Text), Val(txtAttackEnergy(x).Text))
            Else
                txtAttackAccuSpellName(x).Text = "N/A"
                txtAttackSpellDamage(x).Text = ""
            End If
            lblAttackSpellRange(x).Visible = True
            txtAttackSpellDamage(x).Visible = True
            
        Case Else:
            lblAttackAccuSpell(x).Caption = "Accu/Spell"
            lblAttackMinHCastPer(x).Caption = "MinH/Cast%"
            lblAttackMaxHCastLvl(x).Caption = "MaxH/CastLVL"
            
            If Val(txtAttackAccuSpell(x).Text) > 0 Then
                txtAttackAccuSpellName(x).Text = GetSpellName(Val(txtAttackAccuSpell(x).Text))
                txtAttackSpellDamage(x).Text = GetSpellRange(Val(txtAttackAccuSpell(x).Text), _
                    Val(txtAttackMaxHCastLvL(x).Text), Val(txtAttackEnergy(x).Text))
                lblAttackSpellRange(x).Visible = True
                txtAttackSpellDamage(x).Visible = True
            Else
                txtAttackAccuSpellName(x).Text = "N/A"
                txtAttackSpellDamage(x).Text = ""
                lblAttackSpellRange(x).Visible = False
                txtAttackSpellDamage(x).Visible = False
            End If
            
    End Select
Next x

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("RefreshAttackWindow")
Resume out:

End Sub
Private Sub txtAttackAccuSpell_GotFocus(Index As Integer)
Call SelectAll(txtAttackAccuSpell(Index))

End Sub

Private Sub txtAttackDodgeMsg_Change(Index As Integer)
On Error GoTo error:

txtAttackDodgeMsgDisplay(Index).Text = GetMessages(Val(txtAttackDodgeMsg(Index).Text), 3)

out:
Exit Sub
error:
Call HandleError("txtAttackDodgeMsg_Change")
Resume out:
End Sub

Private Sub txtAttackDodgeMsg_GotFocus(Index As Integer)
Call SelectAll(txtAttackDodgeMsg(Index))

End Sub

Private Sub txtAttackEnergy_Change(Index As Integer)
Call RefreshAttackWindow(Index)
End Sub

Private Sub txtAttackEnergy_GotFocus(Index As Integer)
Call SelectAll(txtAttackEnergy(Index))

End Sub

Private Sub txtAttackHitMsg_Change(Index As Integer)
On Error GoTo error:

txtAttackHitMsgDisplay(Index).Text = GetMessages(Val(txtAttackHitMsg(Index).Text), 1)

out:
Exit Sub
error:
Call HandleError("txtAttackHitMsg_Change")
Resume out:
End Sub

Private Sub txtAttackHitMsg_GotFocus(Index As Integer)
Call SelectAll(txtAttackHitMsg(Index))

End Sub

Private Sub txtAttackHitSpell_Change(Index As Integer)
On Error GoTo error:

txtAttackHitSpellName(Index).Text = GetSpellName(Val(txtAttackHitSpell(Index).Text))

out:
Exit Sub
error:
Call HandleError("txtAttackHitSpell_Change")
Resume out:
End Sub

Private Sub txtAttackHitSpell_GotFocus(Index As Integer)
Call SelectAll(txtAttackHitSpell(Index))

End Sub

Private Sub txtAttackMaxHCastLvL_Change(Index As Integer)
Call RefreshAttackWindow(Index)
End Sub

Private Sub txtAttackMaxHCastLvL_GotFocus(Index As Integer)
Call SelectAll(txtAttackMaxHCastLvL(Index))

End Sub

Private Sub txtAttackMinHCastPer_GotFocus(Index As Integer)
Call SelectAll(txtAttackMinHCastPer(Index))

End Sub

Private Sub txtAttackMissMsg_Change(Index As Integer)
On Error GoTo error:

txtAttackMissMsgDisplay(Index).Text = GetMessages(Val(txtAttackMissMsg(Index).Text), 2)

out:
Exit Sub
error:
Call HandleError("txtAttackMissMsg_Change")
Resume out:
End Sub

Private Sub txtAttackMissMsg_GotFocus(Index As Integer)
Call SelectAll(txtAttackMissMsg(Index))

End Sub

Private Sub txtAttackPer_GotFocus(Index As Integer)
Call SelectAll(txtAttackPer(Index))

End Sub

Private Sub txtBase_Change()
'If Val(txtBase.Text) > 65535 Then txtBase.Text = 65535
txtExperience.Text = Val(txtBase.Text) * Val(txtMulti.Text)
End Sub

Private Sub txtBase_GotFocus()
Call SelectAll(txtBase)

End Sub

Private Sub txtBase_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtBSDefense_GotFocus()
Call SelectAll(txtBSDefense)

End Sub

Private Sub txtCharmlvl_GotFocus()
Call SelectAll(txtCharmlvl)

End Sub

Private Sub txtCharmRes_GotFocus()
Call SelectAll(txtCharmRes)

End Sub

Private Sub txtCopper_GotFocus()
Call SelectAll(txtCopper)

End Sub

Private Sub txtCreateSpellNumber_Change()
On Error GoTo error:

txtCreateSpellName.Text = GetSpellName(Val(txtCreateSpellNumber.Text))

out:
Exit Sub
error:
Call HandleError("txtCreateSpellNumber_Change")
Resume out:
End Sub

Private Sub txtCreateSpellNumber_GotFocus()
Call SelectAll(txtCreateSpellNumber)

End Sub

Private Sub txtDateKilled_GotFocus()
Call SelectAll(txtDateKilled)

End Sub


Private Sub txtDeathMsg_Change()
On Error GoTo error:

txtDeathMsgDisplay.Text = GetMessages(Val(txtDeathMsg.Text), 3)

out:
Exit Sub
error:
Call HandleError("txtDeathMsg_Change")
Resume out:
End Sub

Private Sub txtDeathMsg_GotFocus()
Call SelectAll(txtDeathMsg)

End Sub

Private Sub txtDeathSpellNumber_Change()
On Error GoTo error:

txtDeathSpellName.Text = GetSpellName(Val(txtDeathSpellNumber.Text))

out:
Exit Sub
error:
Call HandleError("txtDeathSpellNumber_Change")
Resume out:
End Sub

Private Sub txtDeathSpellNumber_GotFocus()
Call SelectAll(txtDeathSpellNumber)

End Sub

Private Sub txtDesc_Change(Index As Integer)
If Index = 3 Then Exit Sub
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


Private Sub txtDescTxt_Change()
On Error GoTo error:

txtDescTxtDisplay.Text = GetTextblock(Val(txtDescTxt.Text))

out:
Exit Sub
error:
Call HandleError("txtDescTxt_Change")
Resume out:
End Sub

Private Sub txtDescTxt_GotFocus()
Call SelectAll(txtDescTxt)

End Sub

Private Sub txtDR_GotFocus()
Call SelectAll(txtDR)

End Sub

Private Sub txtEnergy_GotFocus()
Call SelectAll(txtEnergy)

End Sub

Private Sub txtExperience_GotFocus()
Call SelectAll(txtExperience)

End Sub

Private Sub txtExperience_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtFollow_GotFocus()
Call SelectAll(txtFollow)

End Sub

Private Sub txtGameLimit_GotFocus()
Call SelectAll(txtGameLimit)

End Sub

Private Sub txtGold_GotFocus()
Call SelectAll(txtGold)

End Sub

Private Sub txtGreetTxt_Change()
On Error GoTo error:

txtGreetTxtDisplay.Text = GetTextblock(Val(txtGreetTxt.Text))

out:
Exit Sub
error:
Call HandleError("txtGreetTxt_Change")
Resume out:
End Sub

Private Sub txtGreetTxt_GotFocus()
Call SelectAll(txtGreetTxt)

End Sub

Private Sub txtHitPoints_GotFocus()
Call SelectAll(txtHitPoints)

End Sub

Private Sub txtHpRegen_GotFocus()
Call SelectAll(txtHpRegen)

End Sub

Private Sub txtIndex_GotFocus()
Call SelectAll(txtIndex)

End Sub

Private Sub txtItemDropPer_GotFocus(Index As Integer)
Call SelectAll(txtItemDropPer(Index))

End Sub

Private Sub txtItemNumber_Change(Index As Integer)
On Error GoTo error:

txtItemName(Index).Text = GetItemName(Val(txtItemNumber(Index).Text))

out:
Exit Sub
error:
Call HandleError("txtItemNumber_Change")
Resume out:
End Sub

Private Sub txtItemNumber_GotFocus(Index As Integer)
Call SelectAll(txtItemNumber(Index))

End Sub

Private Sub txtItemUses_GotFocus(Index As Integer)
Call SelectAll(txtItemUses(Index))

End Sub

Private Sub txtMoveMsg_Change()
On Error GoTo error:

txtMoveMsgDisplay.Text = GetMessages(Val(txtMoveMsg.Text), 1)

out:
Exit Sub
error:
Call HandleError("txtMoveMsg_Change")
Resume out:
End Sub

Private Sub txtMoveMsg_GotFocus()
Call SelectAll(txtMoveMsg)

End Sub

Private Sub txtMR_GotFocus()
Call SelectAll(txtMR)

End Sub

Private Sub txtMulti_Change()
'If Val(txtMulti.Text) > 65535 Then txtMulti.Text = 65535
txtExperience.Text = Val(txtBase.Text) * Val(txtMulti.Text)
End Sub

Private Sub txtMulti_GotFocus()
Call SelectAll(txtMulti)

End Sub

Private Sub txtMulti_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
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

Private Sub txtPlatinum_GotFocus()
Call SelectAll(txtPlatinum)

End Sub

Private Sub txtRegenTime_GotFocus()
Call SelectAll(txtRegenTime)

End Sub

Private Sub txtRunic_GotFocus()
Call SelectAll(txtRunic)

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

Private Sub txtSilver_GotFocus()
Call SelectAll(txtSilver)

End Sub

Private Sub txtSpellCastLvL_GotFocus(Index As Integer)
Call SelectAll(txtSpellCastLvL(Index))

End Sub

Private Sub txtSpellCastPer_GotFocus(Index As Integer)
Call SelectAll(txtSpellCastPer(Index))

End Sub

Private Sub txtSpellNumber_Change(Index As Integer)
On Error GoTo error:

txtSpellName(Index).Text = GetSpellName(Val(txtSpellNumber(Index).Text))

out:
Exit Sub
error:
Call HandleError("txtSpellNumber_Change")
Resume out:
End Sub

Private Sub txtSpellNumber_GotFocus(Index As Integer)
Call SelectAll(txtSpellNumber(Index))

End Sub

Private Sub txtTalkTxt_Change()
On Error GoTo error:

txtTalkTxtDisplay.Text = GetTextblock(Val(txtTalkTxt.Text))

out:
Exit Sub
error:
Call HandleError("txtTalkTxt_Change")
Resume out:
End Sub

Private Sub txtTalkTxt_GotFocus()
Call SelectAll(txtTalkTxt)

End Sub

Private Sub txtTimeKilled_GotFocus()
Call SelectAll(txtTimeKilled)

End Sub

Private Sub txtWeaponNumber_Change()
On Error GoTo error:

txtWeaponName.Text = GetItemName(Val(txtWeaponNumber.Text))

out:
Exit Sub
error:
Call HandleError("txtWeaponNumber_Change")
Resume out:
End Sub

Private Sub txtWeaponNumber_GotFocus()
Call SelectAll(txtWeaponNumber)

End Sub


