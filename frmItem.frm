VERSION 5.00
Object = "{20D5284F-7B23-4F0A-B8B1-6C9D18B64F1C}#1.0#0"; "exlimiter.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItem 
   Caption         =   "Item Editor"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8970
   Icon            =   "frmItem.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   8970
   Begin VB.Frame fraFilter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   60
      TabIndex        =   243
      Top             =   420
      Visible         =   0   'False
      Width           =   5235
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
         Height          =   5295
         Left            =   180
         TabIndex        =   244
         Top             =   180
         Width           =   4875
         Begin VB.ComboBox cmbFilter 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   9
            ItemData        =   "frmItem.frx":08CA
            Left            =   3120
            List            =   "frmItem.frx":08CC
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   285
            Top             =   2220
            Width           =   1575
         End
         Begin VB.CheckBox chkFilterExcludeZero 
            Caption         =   "Exclude 0 value on <= search"
            Height          =   255
            Left            =   2160
            TabIndex        =   284
            Top             =   240
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Spell"
            Enabled         =   0   'False
            Height          =   255
            Index           =   14
            Left            =   180
            TabIndex        =   283
            Top             =   4800
            Width           =   1095
         End
         Begin VB.TextBox txtFilterSpell 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1500
            MaxLength       =   29
            TabIndex        =   282
            Text            =   "0"
            Top             =   4800
            Width           =   1095
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Message"
            Enabled         =   0   'False
            Height          =   255
            Index           =   12
            Left            =   180
            TabIndex        =   281
            Top             =   4080
            Width           =   1095
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Textblock"
            Enabled         =   0   'False
            Height          =   255
            Index           =   13
            Left            =   180
            TabIndex        =   280
            Top             =   4440
            Width           =   1095
         End
         Begin VB.TextBox txtFilterMessage 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1500
            MaxLength       =   29
            TabIndex        =   279
            Text            =   "0"
            Top             =   4080
            Width           =   1095
         End
         Begin VB.TextBox txtFilterTB 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1500
            MaxLength       =   29
            TabIndex        =   278
            Text            =   "0"
            Top             =   4440
            Width           =   1095
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Ability"
            Enabled         =   0   'False
            Height          =   255
            Index           =   7
            Left            =   180
            TabIndex        =   277
            Top             =   3300
            Width           =   915
         End
         Begin VB.ComboBox cmbFilter 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   7
            ItemData        =   "frmItem.frx":08CE
            Left            =   1500
            List            =   "frmItem.frx":08D0
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   276
            Top             =   3300
            Width           =   1575
         End
         Begin VB.TextBox txtFilterAbilityValue 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   7
            Left            =   3900
            TabIndex        =   275
            Text            =   "0"
            Top             =   3300
            Width           =   555
         End
         Begin VB.ComboBox cmbFilterAbilityGL 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   7
            ItemData        =   "frmItem.frx":08D2
            Left            =   3120
            List            =   "frmItem.frx":08E2
            Style           =   2  'Dropdown List
            TabIndex        =   274
            Top             =   3300
            Width           =   735
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Ability"
            Enabled         =   0   'False
            Height          =   255
            Index           =   6
            Left            =   180
            TabIndex        =   273
            Top             =   2940
            Width           =   915
         End
         Begin VB.ComboBox cmbFilter 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   6
            ItemData        =   "frmItem.frx":08F6
            Left            =   1500
            List            =   "frmItem.frx":08F8
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   272
            Top             =   2940
            Width           =   1575
         End
         Begin VB.TextBox txtFilterAbilityValue 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   6
            Left            =   3900
            TabIndex        =   271
            Text            =   "0"
            Top             =   2940
            Width           =   555
         End
         Begin VB.ComboBox cmbFilterAbilityGL 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   6
            ItemData        =   "frmItem.frx":08FA
            Left            =   3120
            List            =   "frmItem.frx":090A
            Style           =   2  'Dropdown List
            TabIndex        =   270
            Top             =   2940
            Width           =   735
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Non-Magical"
            Enabled         =   0   'False
            Height          =   255
            Index           =   11
            Left            =   3360
            TabIndex        =   269
            Top             =   1860
            Width           =   1395
         End
         Begin VB.TextBox txtFilterAbilityValue 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   8
            Left            =   3900
            TabIndex        =   268
            Text            =   "0"
            Top             =   3660
            Width           =   555
         End
         Begin VB.ComboBox cmbFilterAbilityGL 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   8
            ItemData        =   "frmItem.frx":091E
            Left            =   3120
            List            =   "frmItem.frx":092E
            Style           =   2  'Dropdown List
            TabIndex        =   267
            Top             =   3660
            Width           =   735
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Cost"
            Enabled         =   0   'False
            Height          =   255
            Index           =   8
            Left            =   180
            TabIndex        =   266
            Top             =   3660
            Width           =   915
         End
         Begin VB.ComboBox cmbFilter 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   8
            ItemData        =   "frmItem.frx":0942
            Left            =   1500
            List            =   "frmItem.frx":0955
            Style           =   2  'Dropdown List
            TabIndex        =   265
            Top             =   3660
            Width           =   1575
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Uses > 0"
            Enabled         =   0   'False
            Height          =   255
            Index           =   10
            Left            =   3360
            TabIndex        =   264
            Top             =   1500
            Width           =   1395
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Limit > 0"
            Enabled         =   0   'False
            Height          =   255
            Index           =   9
            Left            =   3360
            TabIndex        =   263
            Top             =   1140
            Width           =   1395
         End
         Begin VB.ComboBox cmbFilter 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   3
            ItemData        =   "frmItem.frx":0980
            Left            =   1500
            List            =   "frmItem.frx":09A2
            Style           =   2  'Dropdown List
            TabIndex        =   262
            Top             =   1860
            Width           =   1575
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Armour"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   180
            TabIndex        =   261
            Top             =   1860
            Width           =   1035
         End
         Begin VB.ComboBox cmbFilter 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   2
            ItemData        =   "frmItem.frx":0A29
            Left            =   1500
            List            =   "frmItem.frx":0A39
            Style           =   2  'Dropdown List
            TabIndex        =   260
            Top             =   1500
            Width           =   1575
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Weapon"
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   259
            Top             =   1500
            Width           =   1035
         End
         Begin VB.ComboBox cmbFilterAbilityGL 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   5
            ItemData        =   "frmItem.frx":0A65
            Left            =   3120
            List            =   "frmItem.frx":0A75
            Style           =   2  'Dropdown List
            TabIndex        =   258
            Top             =   2580
            Width           =   735
         End
         Begin VB.TextBox txtFilterAbilityValue 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   5
            Left            =   3900
            TabIndex        =   257
            Text            =   "0"
            Top             =   2580
            Width           =   555
         End
         Begin VB.CommandButton cmdFilterReset 
            Caption         =   "Reset"
            Height          =   315
            Left            =   3420
            TabIndex        =   256
            Top             =   540
            Width           =   1275
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
            Left            =   3120
            TabIndex        =   255
            Top             =   4080
            Width           =   1575
         End
         Begin VB.CommandButton cmdFilterCancel 
            Caption         =   "Cancel"
            Height          =   495
            Left            =   3120
            TabIndex        =   254
            Top             =   4620
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
            Left            =   180
            TabIndex        =   253
            Top             =   360
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.ComboBox cmbFilter 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   0
            ItemData        =   "frmItem.frx":0A89
            Left            =   1500
            List            =   "frmItem.frx":0AAE
            Style           =   2  'Dropdown List
            TabIndex        =   252
            Top             =   780
            Width           =   1575
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Type"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   251
            Top             =   780
            Width           =   915
         End
         Begin VB.ComboBox cmbFilter 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   5
            ItemData        =   "frmItem.frx":0B09
            Left            =   1500
            List            =   "frmItem.frx":0B0B
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   250
            Top             =   2580
            Width           =   1575
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Ability"
            Enabled         =   0   'False
            Height          =   255
            Index           =   5
            Left            =   180
            TabIndex        =   249
            Top             =   2580
            Width           =   915
         End
         Begin VB.ComboBox cmbFilter 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   4
            ItemData        =   "frmItem.frx":0B0D
            Left            =   1500
            List            =   "frmItem.frx":0B0F
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   248
            Top             =   2220
            Width           =   1575
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Class/Race"
            Enabled         =   0   'False
            Height          =   255
            Index           =   4
            Left            =   180
            TabIndex        =   247
            Top             =   2220
            Width           =   1155
         End
         Begin VB.ComboBox cmbFilter 
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   1
            ItemData        =   "frmItem.frx":0B11
            Left            =   1500
            List            =   "frmItem.frx":0B51
            Style           =   2  'Dropdown List
            TabIndex        =   246
            Top             =   1140
            Width           =   1575
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Worn"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   245
            Top             =   1140
            Width           =   915
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
      Top             =   60
      Width           =   2895
   End
   Begin VB.Frame framNav 
      BorderStyle     =   0  'None
      Height          =   6075
      Left            =   3060
      TabIndex        =   6
      Top             =   0
      Width           =   5835
      Begin VB.CheckBox chkAutoSave 
         Caption         =   "Auto-Save"
         Height          =   195
         Left            =   2640
         TabIndex        =   242
         Top             =   120
         Value           =   1  'Checked
         Width           =   1155
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   345
         Left            =   960
         TabIndex        =   8
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdDiscard 
         Caption         =   "Dis&card"
         Height          =   345
         Left            =   4860
         TabIndex        =   10
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "&Insert"
         Height          =   345
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   345
         Left            =   3900
         TabIndex        =   9
         Top             =   0
         Width           =   975
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5655
         Left            =   0
         TabIndex        =   11
         Top             =   360
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   9975
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "frmItem.frx":0BEF
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "label(21)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "label(20)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "label(19)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "label(18)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "label(17)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "label(15)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "label(14)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "label(13)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "label(12)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "label(11)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "label(10)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "label(9)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Line1"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "label(8)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "label(1)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "label(0)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Label4"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Label5"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "lblACDR"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "label(6)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Label7"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "txtDR"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "chkDestroy"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "chkGettable"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "chkNotDroppable"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "txtAC"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "txtReqSTR"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "txtAccuracy"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "txtSpeed"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "txtMaxHit"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "txtMinHit"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "cmbCostType"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "txtCost"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "txtUses"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "txtLimit"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "txtWeight"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "cmbWeapon"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).Control(37)=   "cmbArmour"
         Tab(0).Control(37).Enabled=   0   'False
         Tab(0).Control(38)=   "cmbWornOn"
         Tab(0).Control(38).Enabled=   0   'False
         Tab(0).Control(39)=   "cmbType"
         Tab(0).Control(39).Enabled=   0   'False
         Tab(0).Control(40)=   "txtName"
         Tab(0).Control(40).Enabled=   0   'False
         Tab(0).Control(41)=   "txtNumber"
         Tab(0).Control(41).Enabled=   0   'False
         Tab(0).Control(42)=   "chkRetainAfterUses"
         Tab(0).Control(42).Enabled=   0   'False
         Tab(0).Control(43)=   "chkRobable"
         Tab(0).Control(43).Enabled=   0   'False
         Tab(0).Control(44)=   "txtUnknown1"
         Tab(0).Control(44).Enabled=   0   'False
         Tab(0).Control(45)=   "txtUnknown8"
         Tab(0).Control(45).Enabled=   0   'False
         Tab(0).Control(46)=   "cmdCalcSwings"
         Tab(0).Control(46).Enabled=   0   'False
         Tab(0).Control(47)=   "Text1"
         Tab(0).Control(47).Enabled=   0   'False
         Tab(0).ControlCount=   48
         TabCaption(1)   =   "Desc/Msg/Cash"
         TabPicture(1)   =   "frmItem.frx":0C0B
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame6"
         Tab(1).Control(1)=   "Frame5"
         Tab(1).Control(2)=   "Frame2"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Race/Class/Negate"
         TabPicture(2)   =   "frmItem.frx":0C27
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame1"
         Tab(2).Control(1)=   "Frame4"
         Tab(2).Control(2)=   "Frame3"
         Tab(2).ControlCount=   3
         TabCaption(3)   =   "Abilities p1"
         TabPicture(3)   =   "frmItem.frx":0C43
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "cmdAbilsClear(0)"
         Tab(3).Control(1)=   "frmAbilities(0)"
         Tab(3).Control(2)=   "Label8"
         Tab(3).ControlCount=   3
         TabCaption(4)   =   "Abilities p2"
         TabPicture(4)   =   "frmItem.frx":0C5F
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Label9"
         Tab(4).Control(1)=   "frmAbilities(1)"
         Tab(4).Control(2)=   "cmdAbilsClear(1)"
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
            Index           =   1
            Left            =   -74760
            TabIndex        =   241
            Top             =   600
            Width           =   675
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
            Index           =   0
            Left            =   -74760
            TabIndex        =   149
            Top             =   600
            Width           =   675
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   4500
            TabIndex        =   57
            Top             =   4620
            Width           =   675
         End
         Begin VB.Frame Frame6 
            Caption         =   "Cash when opened (container item type only)"
            Height          =   1035
            Left            =   -74160
            TabIndex        =   82
            Top             =   4380
            Width           =   4095
            Begin VB.TextBox txtRunic 
               Height          =   315
               Left            =   180
               TabIndex        =   88
               Top             =   540
               Width           =   615
            End
            Begin VB.TextBox txtPlatinum 
               Height          =   315
               Left            =   960
               TabIndex        =   89
               Top             =   540
               Width           =   615
            End
            Begin VB.TextBox txtGold 
               Height          =   315
               Left            =   1740
               TabIndex        =   90
               Top             =   540
               Width           =   615
            End
            Begin VB.TextBox txtSilver 
               Height          =   315
               Left            =   2520
               TabIndex        =   91
               Top             =   540
               Width           =   615
            End
            Begin VB.TextBox txtCopper 
               Height          =   315
               Left            =   3300
               TabIndex        =   92
               Top             =   540
               Width           =   615
            End
            Begin VB.Label label 
               Caption         =   "Runic"
               Height          =   195
               Index           =   25
               Left            =   180
               TabIndex        =   83
               Top             =   360
               Width           =   615
            End
            Begin VB.Label label 
               Caption         =   "Platinum"
               Height          =   195
               Index           =   26
               Left            =   960
               TabIndex        =   84
               Top             =   360
               Width           =   615
            End
            Begin VB.Label label 
               Caption         =   "Gold"
               Height          =   195
               Index           =   27
               Left            =   1740
               TabIndex        =   85
               Top             =   360
               Width           =   435
            End
            Begin VB.Label label 
               Caption         =   "Silver"
               Height          =   195
               Index           =   28
               Left            =   2520
               TabIndex        =   86
               Top             =   360
               Width           =   495
            End
            Begin VB.Label label 
               Caption         =   "Copper"
               Height          =   195
               Index           =   29
               Left            =   3300
               TabIndex        =   87
               Top             =   360
               Width           =   615
            End
         End
         Begin VB.CommandButton cmdCalcSwings 
            Caption         =   "Calculate Swings"
            Height          =   555
            Left            =   4680
            TabIndex        =   12
            Top             =   420
            Width           =   975
         End
         Begin VB.TextBox txtUnknown8 
            Height          =   285
            Left            =   3840
            TabIndex        =   56
            Top             =   4620
            Width           =   615
         End
         Begin VB.TextBox txtUnknown1 
            Height          =   285
            Left            =   3180
            TabIndex        =   55
            Top             =   4620
            Width           =   615
         End
         Begin VB.CheckBox chkRobable 
            Caption         =   "Robable"
            Height          =   255
            Left            =   3180
            TabIndex        =   41
            Top             =   2580
            Width           =   1095
         End
         Begin VB.CheckBox chkRetainAfterUses 
            Caption         =   "Retain After Uses Expire"
            Height          =   255
            Left            =   3180
            TabIndex        =   40
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Frame Frame1 
            Caption         =   "Negate Spells"
            Height          =   1935
            Left            =   -74880
            TabIndex        =   117
            Top             =   3600
            Width           =   5595
            Begin VB.CommandButton cmdResetNegate 
               Caption         =   "Reset"
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
               Left            =   4800
               TabIndex        =   118
               Top             =   60
               Width           =   735
            End
            Begin VB.TextBox txtNegate 
               Height          =   315
               Index           =   9
               Left            =   3120
               TabIndex        =   147
               Top             =   1500
               Width           =   615
            End
            Begin VB.TextBox txtNegateName 
               BackColor       =   &H8000000F&
               Height          =   315
               Index           =   9
               Left            =   3720
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   148
               TabStop         =   0   'False
               Top             =   1500
               Width           =   1755
            End
            Begin VB.CommandButton cmdEditNegate 
               Height          =   195
               Index           =   9
               Left            =   2880
               TabIndex        =   146
               Top             =   1560
               Width           =   195
            End
            Begin VB.TextBox txtNegate 
               Height          =   315
               Index           =   8
               Left            =   3120
               TabIndex        =   144
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtNegateName 
               BackColor       =   &H8000000F&
               Height          =   315
               Index           =   8
               Left            =   3720
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   145
               TabStop         =   0   'False
               Top             =   1200
               Width           =   1755
            End
            Begin VB.CommandButton cmdEditNegate 
               Height          =   195
               Index           =   8
               Left            =   2880
               TabIndex        =   143
               Top             =   1260
               Width           =   195
            End
            Begin VB.TextBox txtNegate 
               Height          =   315
               Index           =   7
               Left            =   3120
               TabIndex        =   141
               Top             =   900
               Width           =   615
            End
            Begin VB.TextBox txtNegateName 
               BackColor       =   &H8000000F&
               Height          =   315
               Index           =   7
               Left            =   3720
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   142
               TabStop         =   0   'False
               Top             =   900
               Width           =   1755
            End
            Begin VB.CommandButton cmdEditNegate 
               Height          =   195
               Index           =   7
               Left            =   2880
               TabIndex        =   140
               Top             =   960
               Width           =   195
            End
            Begin VB.TextBox txtNegate 
               Height          =   315
               Index           =   6
               Left            =   3120
               TabIndex        =   138
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox txtNegateName 
               BackColor       =   &H8000000F&
               Height          =   315
               Index           =   6
               Left            =   3720
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   139
               TabStop         =   0   'False
               Top             =   600
               Width           =   1755
            End
            Begin VB.CommandButton cmdEditNegate 
               Height          =   195
               Index           =   6
               Left            =   2880
               TabIndex        =   137
               Top             =   660
               Width           =   195
            End
            Begin VB.TextBox txtNegate 
               Height          =   315
               Index           =   5
               Left            =   3120
               TabIndex        =   135
               Top             =   300
               Width           =   615
            End
            Begin VB.TextBox txtNegateName 
               BackColor       =   &H8000000F&
               Height          =   315
               Index           =   5
               Left            =   3720
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   136
               TabStop         =   0   'False
               Top             =   300
               Width           =   1755
            End
            Begin VB.CommandButton cmdEditNegate 
               Height          =   195
               Index           =   5
               Left            =   2880
               TabIndex        =   134
               Top             =   360
               Width           =   195
            End
            Begin VB.CommandButton cmdEditNegate 
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   131
               Top             =   1560
               Width           =   195
            End
            Begin VB.CommandButton cmdEditNegate 
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   128
               Top             =   1260
               Width           =   195
            End
            Begin VB.CommandButton cmdEditNegate 
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   125
               Top             =   960
               Width           =   195
            End
            Begin VB.CommandButton cmdEditNegate 
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   122
               Top             =   660
               Width           =   195
            End
            Begin VB.CommandButton cmdEditNegate 
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   119
               Top             =   360
               Width           =   195
            End
            Begin VB.TextBox txtNegateName 
               BackColor       =   &H8000000F&
               Height          =   315
               Index           =   4
               Left            =   960
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   133
               TabStop         =   0   'False
               Top             =   1500
               Width           =   1815
            End
            Begin VB.TextBox txtNegateName 
               BackColor       =   &H8000000F&
               Height          =   315
               Index           =   3
               Left            =   960
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   130
               TabStop         =   0   'False
               Top             =   1200
               Width           =   1815
            End
            Begin VB.TextBox txtNegateName 
               BackColor       =   &H8000000F&
               Height          =   315
               Index           =   2
               Left            =   960
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   127
               TabStop         =   0   'False
               Top             =   900
               Width           =   1815
            End
            Begin VB.TextBox txtNegateName 
               BackColor       =   &H8000000F&
               Height          =   315
               Index           =   1
               Left            =   960
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   124
               TabStop         =   0   'False
               Top             =   600
               Width           =   1815
            End
            Begin VB.TextBox txtNegateName 
               BackColor       =   &H8000000F&
               Height          =   315
               Index           =   0
               Left            =   960
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   121
               TabStop         =   0   'False
               Top             =   300
               Width           =   1815
            End
            Begin VB.TextBox txtNegate 
               Height          =   315
               Index           =   4
               Left            =   360
               TabIndex        =   132
               Top             =   1500
               Width           =   615
            End
            Begin VB.TextBox txtNegate 
               Height          =   315
               Index           =   3
               Left            =   360
               TabIndex        =   129
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtNegate 
               Height          =   315
               Index           =   2
               Left            =   360
               TabIndex        =   126
               Top             =   900
               Width           =   615
            End
            Begin VB.TextBox txtNegate 
               Height          =   315
               Index           =   1
               Left            =   360
               TabIndex        =   123
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox txtNegate 
               Height          =   315
               Index           =   0
               Left            =   360
               TabIndex        =   120
               Top             =   300
               Width           =   615
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Messages/Textblocks"
            Height          =   1635
            Left            =   -74760
            TabIndex        =   65
            Top             =   2580
            Width           =   5295
            Begin VB.CommandButton cmdEditDestructMsg 
               Height          =   195
               Left            =   120
               TabIndex        =   78
               Top             =   1200
               Width           =   195
            End
            Begin VB.CommandButton cmdEditMissMsg 
               Height          =   195
               Left            =   120
               TabIndex        =   74
               Top             =   900
               Width           =   195
            End
            Begin VB.CommandButton cmdEditReadMsg 
               Height          =   195
               Left            =   120
               TabIndex        =   70
               Top             =   600
               Width           =   195
            End
            Begin VB.CommandButton cmdEditHitMsg 
               Height          =   195
               Left            =   120
               TabIndex        =   66
               Top             =   300
               Width           =   195
            End
            Begin VB.TextBox txtMissMsg 
               Height          =   285
               Left            =   1440
               TabIndex        =   76
               Top             =   900
               Width           =   615
            End
            Begin VB.TextBox txtDistructMsg 
               Height          =   285
               Left            =   1440
               TabIndex        =   80
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtHitMsg 
               Height          =   285
               Left            =   1440
               TabIndex        =   68
               Top             =   300
               Width           =   615
            End
            Begin VB.TextBox txtReadMsg 
               Height          =   285
               Left            =   1440
               TabIndex        =   72
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox txtHitMsgDisplay 
               BackColor       =   &H8000000F&
               Height          =   285
               Left            =   2040
               Locked          =   -1  'True
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   300
               Width           =   3135
            End
            Begin VB.TextBox txtReadMsgDisplay 
               BackColor       =   &H8000000F&
               Height          =   285
               Left            =   2040
               Locked          =   -1  'True
               TabIndex        =   73
               TabStop         =   0   'False
               Top             =   600
               Width           =   3135
            End
            Begin VB.TextBox txtMissMsgDisplay 
               BackColor       =   &H8000000F&
               Height          =   285
               Left            =   2040
               Locked          =   -1  'True
               TabIndex        =   77
               TabStop         =   0   'False
               Top             =   900
               Width           =   3135
            End
            Begin VB.TextBox txtDistructMsgDisplay 
               BackColor       =   &H8000000F&
               Height          =   285
               Left            =   2040
               Locked          =   -1  'True
               TabIndex        =   81
               TabStop         =   0   'False
               Top             =   1200
               Width           =   3135
            End
            Begin VB.Label label 
               Alignment       =   1  'Right Justify
               Caption         =   "Read (tb)"
               Height          =   255
               Index           =   5
               Left            =   420
               TabIndex        =   71
               Top             =   600
               Width           =   975
            End
            Begin VB.Label label 
               Alignment       =   1  'Right Justify
               Caption         =   "Destrct (msg)"
               Height          =   255
               Index           =   4
               Left            =   420
               TabIndex        =   79
               Top             =   1200
               Width           =   975
            End
            Begin VB.Label label 
               Alignment       =   1  'Right Justify
               Caption         =   "Hit (msg)"
               Height          =   255
               Index           =   3
               Left            =   420
               TabIndex        =   67
               Top             =   300
               Width           =   975
            End
            Begin VB.Label label 
               Alignment       =   1  'Right Justify
               Caption         =   "Miss (msg)"
               Height          =   255
               Index           =   2
               Left            =   420
               TabIndex        =   75
               Top             =   900
               Width           =   975
            End
         End
         Begin VB.Frame frmAbilities 
            Caption         =   "Abilities"
            Height          =   4095
            Index           =   1
            Left            =   -73920
            TabIndex        =   194
            Top             =   540
            Width           =   3555
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   19
               Left            =   3300
               TabIndex        =   237
               Top             =   3720
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   18
               Left            =   3300
               TabIndex        =   233
               Top             =   3360
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   17
               Left            =   3300
               TabIndex        =   229
               Top             =   3000
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   16
               Left            =   3300
               TabIndex        =   225
               Top             =   2640
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   15
               Left            =   3300
               TabIndex        =   221
               Top             =   2280
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   14
               Left            =   3300
               TabIndex        =   217
               Top             =   1920
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   13
               Left            =   3300
               TabIndex        =   213
               Top             =   1560
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   12
               Left            =   3300
               TabIndex        =   209
               Top             =   1200
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   11
               Left            =   3300
               TabIndex        =   205
               Top             =   840
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   10
               Left            =   3300
               TabIndex        =   201
               Top             =   480
               Width           =   135
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   19
               Left            =   120
               TabIndex        =   234
               Top             =   3720
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   18
               Left            =   120
               TabIndex        =   230
               Top             =   3360
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   17
               Left            =   120
               TabIndex        =   226
               Top             =   3000
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   16
               Left            =   120
               TabIndex        =   222
               Top             =   2640
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   15
               Left            =   120
               TabIndex        =   218
               Top             =   2280
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   14
               Left            =   120
               TabIndex        =   214
               Top             =   1920
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   13
               Left            =   120
               TabIndex        =   210
               Top             =   1560
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   12
               Left            =   120
               TabIndex        =   206
               Top             =   1200
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   11
               Left            =   120
               TabIndex        =   202
               Top             =   840
               Width           =   495
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   19
               Left            =   2640
               TabIndex        =   236
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   3720
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   18
               Left            =   2640
               TabIndex        =   232
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   3360
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   17
               Left            =   2640
               TabIndex        =   228
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   3000
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   16
               Left            =   2640
               TabIndex        =   224
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   2640
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   15
               Left            =   2640
               TabIndex        =   220
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   2280
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   14
               Left            =   2640
               TabIndex        =   216
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   1920
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   13
               Left            =   2640
               TabIndex        =   212
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   1560
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   12
               Left            =   2640
               TabIndex        =   208
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   11
               Left            =   2640
               TabIndex        =   204
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   840
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   10
               Left            =   2640
               TabIndex        =   200
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   10
               Left            =   120
               TabIndex        =   198
               Top             =   480
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
               Index           =   19
               Left            =   720
               TabIndex        =   235
               Text            =   "empty"
               Top             =   3720
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
               Index           =   18
               Left            =   720
               TabIndex        =   231
               Text            =   "empty"
               Top             =   3360
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
               Index           =   17
               Left            =   720
               TabIndex        =   227
               Text            =   "empty"
               Top             =   3000
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
               Index           =   16
               Left            =   720
               TabIndex        =   223
               Text            =   "empty"
               Top             =   2640
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
               Index           =   15
               Left            =   720
               TabIndex        =   219
               Text            =   "empty"
               Top             =   2280
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
               Index           =   14
               Left            =   720
               TabIndex        =   215
               Text            =   "empty"
               Top             =   1920
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
               Index           =   13
               Left            =   720
               TabIndex        =   211
               Text            =   "empty"
               Top             =   1560
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
               Index           =   12
               Left            =   720
               TabIndex        =   207
               Text            =   "empty"
               Top             =   1200
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
               Index           =   11
               Left            =   720
               TabIndex        =   203
               Text            =   "empty"
               Top             =   840
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
               Index           =   10
               Left            =   720
               TabIndex        =   199
               Text            =   "empty"
               Top             =   480
               Width           =   1815
            End
            Begin VB.Label Label1 
               Caption         =   "#"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   195
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label2 
               Caption         =   "Name"
               Height          =   255
               Index           =   1
               Left            =   720
               TabIndex        =   196
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label Label3 
               Caption         =   "Value"
               Height          =   255
               Index           =   1
               Left            =   2640
               TabIndex        =   197
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.Frame frmAbilities 
            Caption         =   "Abilities"
            Height          =   4095
            Index           =   0
            Left            =   -73920
            TabIndex        =   150
            Top             =   540
            Width           =   3555
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   9
               Left            =   3300
               TabIndex        =   193
               Top             =   3720
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   8
               Left            =   3300
               TabIndex        =   189
               Top             =   3360
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   7
               Left            =   3300
               TabIndex        =   185
               Top             =   3000
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   6
               Left            =   3300
               TabIndex        =   181
               Top             =   2640
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   5
               Left            =   3300
               TabIndex        =   177
               Top             =   2280
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   4
               Left            =   3300
               TabIndex        =   173
               Top             =   1920
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   3
               Left            =   3300
               TabIndex        =   169
               Top             =   1560
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   2
               Left            =   3300
               TabIndex        =   165
               Top             =   1200
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   1
               Left            =   3300
               TabIndex        =   161
               Top             =   840
               Width           =   135
            End
            Begin VB.CommandButton cmdAbilityLookup 
               Height          =   255
               Index           =   0
               Left            =   3300
               TabIndex        =   157
               Top             =   480
               Width           =   135
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
               TabIndex        =   191
               Text            =   "empty"
               Top             =   3720
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
               TabIndex        =   187
               Text            =   "empty"
               Top             =   3360
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
               TabIndex        =   183
               Text            =   "empty"
               Top             =   3000
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
               TabIndex        =   179
               Text            =   "empty"
               Top             =   2640
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
               TabIndex        =   175
               Text            =   "empty"
               Top             =   2280
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
               TabIndex        =   171
               Text            =   "empty"
               Top             =   1920
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
               TabIndex        =   167
               Text            =   "empty"
               Top             =   1560
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
               TabIndex        =   163
               Text            =   "empty"
               Top             =   1200
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
               TabIndex        =   159
               Text            =   "empty"
               Top             =   840
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
               Index           =   0
               Left            =   720
               TabIndex        =   155
               Text            =   "empty"
               Top             =   480
               Width           =   1815
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   0
               Left            =   120
               TabIndex        =   154
               Top             =   480
               Width           =   495
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   4
               Left            =   2640
               TabIndex        =   172
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   1920
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   3
               Left            =   2640
               TabIndex        =   168
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   1560
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   2
               Left            =   2640
               TabIndex        =   164
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   1
               Left            =   2640
               TabIndex        =   160
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   840
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   0
               Left            =   2640
               TabIndex        =   156
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   9
               Left            =   2640
               TabIndex        =   192
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   3720
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   8
               Left            =   2640
               TabIndex        =   188
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   3360
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   7
               Left            =   2640
               TabIndex        =   184
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   3000
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   6
               Left            =   2640
               TabIndex        =   180
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   2640
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   5
               Left            =   2640
               TabIndex        =   176
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   2280
               Width           =   615
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   1
               Left            =   120
               TabIndex        =   158
               Top             =   840
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   2
               Left            =   120
               TabIndex        =   162
               Top             =   1200
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   3
               Left            =   120
               TabIndex        =   166
               Top             =   1560
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   4
               Left            =   120
               TabIndex        =   170
               Top             =   1920
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   5
               Left            =   120
               TabIndex        =   174
               Top             =   2280
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   6
               Left            =   120
               TabIndex        =   178
               Top             =   2640
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   7
               Left            =   120
               TabIndex        =   182
               Top             =   3000
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   8
               Left            =   120
               TabIndex        =   186
               Top             =   3360
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   9
               Left            =   120
               TabIndex        =   190
               Top             =   3720
               Width           =   495
            End
            Begin VB.Label Label3 
               Caption         =   "Value"
               Height          =   255
               Index           =   0
               Left            =   2640
               TabIndex        =   153
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label2 
               Caption         =   "Name"
               Height          =   255
               Index           =   0
               Left            =   720
               TabIndex        =   152
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label Label1 
               Caption         =   "#"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   151
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Class Restrictions"
            Height          =   1575
            Left            =   -74760
            TabIndex        =   105
            Top             =   1980
            Width           =   5295
            Begin VB.CommandButton cmdResetClass 
               Caption         =   "Reset"
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
               Left            =   4440
               TabIndex        =   106
               Top             =   1260
               Width           =   735
            End
            Begin VB.ComboBox cmbClass 
               Height          =   315
               Index           =   0
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   107
               Top             =   240
               Width           =   1695
            End
            Begin VB.ComboBox cmbClass 
               Height          =   315
               Index           =   1
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   108
               Top             =   540
               Width           =   1695
            End
            Begin VB.ComboBox cmbClass 
               Height          =   315
               Index           =   2
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   109
               Top             =   840
               Width           =   1695
            End
            Begin VB.ComboBox cmbClass 
               Height          =   315
               Index           =   3
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   110
               Top             =   240
               Width           =   1695
            End
            Begin VB.ComboBox cmbClass 
               Height          =   315
               Index           =   4
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   111
               Top             =   540
               Width           =   1695
            End
            Begin VB.ComboBox cmbClass 
               Height          =   315
               Index           =   5
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   112
               Top             =   840
               Width           =   1695
            End
            Begin VB.ComboBox cmbClass 
               Height          =   315
               Index           =   6
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   113
               Top             =   1140
               Width           =   1695
            End
            Begin VB.ComboBox cmbClass 
               Height          =   315
               Index           =   7
               Left            =   3480
               Style           =   2  'Dropdown List
               TabIndex        =   114
               Top             =   240
               Width           =   1695
            End
            Begin VB.ComboBox cmbClass 
               Height          =   315
               Index           =   8
               Left            =   3480
               Style           =   2  'Dropdown List
               TabIndex        =   115
               Top             =   540
               Width           =   1695
            End
            Begin VB.ComboBox cmbClass 
               Height          =   315
               Index           =   9
               Left            =   3480
               Style           =   2  'Dropdown List
               TabIndex        =   116
               Top             =   840
               Width           =   1695
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Description"
            Height          =   2175
            Left            =   -74760
            TabIndex        =   58
            Top             =   360
            Width           =   5295
            Begin VB.TextBox txtDesc 
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   5
               Left            =   120
               MaxLength       =   60
               TabIndex        =   64
               Top             =   1740
               Width           =   5055
            End
            Begin VB.TextBox txtDesc 
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   4
               Left            =   120
               MaxLength       =   60
               TabIndex        =   63
               Top             =   1440
               Width           =   5055
            End
            Begin VB.TextBox txtDesc 
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   3
               Left            =   120
               MaxLength       =   60
               TabIndex        =   62
               Top             =   1140
               Width           =   5055
            End
            Begin VB.TextBox txtDesc 
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   2
               Left            =   120
               MaxLength       =   60
               TabIndex        =   61
               Top             =   840
               Width           =   5055
            End
            Begin VB.TextBox txtDesc 
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   0
               Left            =   120
               MaxLength       =   60
               TabIndex        =   59
               Top             =   240
               Width           =   5055
            End
            Begin VB.TextBox txtDesc 
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   1
               Left            =   120
               MaxLength       =   60
               TabIndex        =   60
               Top             =   540
               Width           =   5055
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Race Restrictions"
            Height          =   1575
            Left            =   -74760
            TabIndex        =   93
            Top             =   360
            Width           =   5295
            Begin VB.CommandButton cmdResetRace 
               Caption         =   "Reset"
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
               Left            =   4440
               TabIndex        =   94
               Top             =   1260
               Width           =   735
            End
            Begin VB.ComboBox cmbRace 
               Height          =   315
               Index           =   0
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   95
               Top             =   240
               Width           =   1695
            End
            Begin VB.ComboBox cmbRace 
               Height          =   315
               Index           =   1
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   96
               Top             =   540
               Width           =   1695
            End
            Begin VB.ComboBox cmbRace 
               Height          =   315
               Index           =   2
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   97
               Top             =   840
               Width           =   1695
            End
            Begin VB.ComboBox cmbRace 
               Height          =   315
               Index           =   3
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   98
               Top             =   240
               Width           =   1695
            End
            Begin VB.ComboBox cmbRace 
               Height          =   315
               Index           =   4
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   99
               Top             =   540
               Width           =   1695
            End
            Begin VB.ComboBox cmbRace 
               Height          =   315
               Index           =   5
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   100
               Top             =   840
               Width           =   1695
            End
            Begin VB.ComboBox cmbRace 
               Height          =   315
               Index           =   6
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   101
               Top             =   1140
               Width           =   1695
            End
            Begin VB.ComboBox cmbRace 
               Height          =   315
               Index           =   7
               Left            =   3480
               Style           =   2  'Dropdown List
               TabIndex        =   102
               Top             =   240
               Width           =   1695
            End
            Begin VB.ComboBox cmbRace 
               Height          =   315
               Index           =   8
               Left            =   3480
               Style           =   2  'Dropdown List
               TabIndex        =   103
               Top             =   540
               Width           =   1695
            End
            Begin VB.ComboBox cmbRace 
               Height          =   315
               Index           =   9
               Left            =   3480
               Style           =   2  'Dropdown List
               TabIndex        =   104
               Top             =   840
               Width           =   1695
            End
         End
         Begin VB.TextBox txtNumber 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   540
            Width           =   675
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   1080
            MaxLength       =   29
            TabIndex        =   16
            Top             =   900
            Width           =   2715
         End
         Begin VB.ComboBox cmbType 
            Height          =   315
            ItemData        =   "frmItem.frx":0C7B
            Left            =   1080
            List            =   "frmItem.frx":0CA0
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1380
            Width           =   1575
         End
         Begin VB.ComboBox cmbWornOn 
            Height          =   315
            ItemData        =   "frmItem.frx":0CFB
            Left            =   1080
            List            =   "frmItem.frx":0D3B
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1740
            Width           =   1575
         End
         Begin VB.ComboBox cmbArmour 
            Height          =   315
            ItemData        =   "frmItem.frx":0DD9
            Left            =   1080
            List            =   "frmItem.frx":0DFB
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   2460
            Width           =   1575
         End
         Begin VB.ComboBox cmbWeapon 
            Height          =   315
            ItemData        =   "frmItem.frx":0E82
            Left            =   1080
            List            =   "frmItem.frx":0E92
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   2100
            Width           =   1575
         End
         Begin VB.TextBox txtWeight 
            Height          =   285
            Left            =   1080
            TabIndex        =   28
            Top             =   3180
            Width           =   615
         End
         Begin VB.TextBox txtLimit 
            Height          =   285
            Left            =   1080
            TabIndex        =   26
            Top             =   2820
            Width           =   615
         End
         Begin VB.TextBox txtUses 
            Height          =   285
            Left            =   1080
            TabIndex        =   36
            Top             =   4620
            Width           =   615
         End
         Begin VB.TextBox txtCost 
            Height          =   315
            Left            =   3180
            TabIndex        =   43
            Top             =   3120
            Width           =   615
         End
         Begin VB.ComboBox cmbCostType 
            Height          =   315
            ItemData        =   "frmItem.frx":0EBE
            Left            =   3900
            List            =   "frmItem.frx":0ED1
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   3120
            Width           =   1095
         End
         Begin VB.TextBox txtMinHit 
            Height          =   285
            Left            =   3180
            TabIndex        =   46
            Top             =   3540
            Width           =   615
         End
         Begin VB.TextBox txtMaxHit 
            Height          =   285
            Left            =   3900
            TabIndex        =   48
            Top             =   3540
            Width           =   615
         End
         Begin VB.TextBox txtSpeed 
            Height          =   285
            Left            =   1080
            TabIndex        =   30
            Top             =   3540
            Width           =   615
         End
         Begin VB.TextBox txtAccuracy 
            Height          =   285
            Left            =   1080
            TabIndex        =   34
            Top             =   4260
            Width           =   615
         End
         Begin VB.TextBox txtReqSTR 
            Height          =   285
            Left            =   1080
            TabIndex        =   32
            Top             =   3900
            Width           =   615
         End
         Begin VB.TextBox txtAC 
            Height          =   285
            Left            =   3180
            TabIndex        =   50
            Top             =   3900
            Width           =   615
         End
         Begin VB.CheckBox chkNotDroppable 
            Caption         =   "Not Droppable"
            Height          =   255
            Left            =   3180
            TabIndex        =   38
            Top             =   1680
            Width           =   1395
         End
         Begin VB.CheckBox chkGettable 
            Caption         =   "Getable"
            Height          =   255
            Left            =   3180
            TabIndex        =   37
            Top             =   1380
            Width           =   975
         End
         Begin VB.CheckBox chkDestroy 
            Caption         =   "Destroys On Death"
            Height          =   255
            Left            =   3180
            TabIndex        =   39
            Top             =   1980
            Width           =   1815
         End
         Begin VB.TextBox txtDR 
            Height          =   285
            Left            =   3900
            TabIndex        =   52
            Top             =   3900
            Width           =   615
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "Type + and - in the '#' field to cycle though the abilities.  You can also type the name of the ability in the 'Name' field."
            Height          =   615
            Left            =   -73800
            TabIndex        =   238
            Top             =   4800
            Width           =   3195
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "Type + and - in the '#' field to cycle though the abilities.  You can also type the name of the ability in the 'Name' field."
            Height          =   615
            Left            =   -73800
            TabIndex        =   239
            Top             =   4800
            Width           =   3195
         End
         Begin VB.Label Label7 
            Caption         =   "-"
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
            Left            =   3810
            TabIndex        =   47
            Top             =   3540
            Width           =   135
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Unknowns"
            Height          =   195
            Index           =   6
            Left            =   2160
            TabIndex        =   54
            Top             =   4680
            Width           =   915
         End
         Begin VB.Label lblACDR 
            Alignment       =   2  'Center
            Caption         =   "(XX/XX)"
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
            Left            =   3180
            TabIndex        =   53
            Top             =   4200
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "/"
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
            Left            =   3810
            TabIndex        =   51
            Top             =   3900
            Width           =   135
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "the unknown values are not exported/imported.  if you know what they are tell us!"
            Height          =   435
            Left            =   2100
            TabIndex        =   240
            Top             =   5100
            Width           =   3375
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
            Left            =   240
            TabIndex        =   13
            Top             =   540
            Width           =   855
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
            TabIndex        =   15
            Top             =   900
            Width           =   855
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Type"
            Height          =   255
            Index           =   8
            Left            =   300
            TabIndex        =   17
            Top             =   1380
            Width           =   675
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   4680
            Y1              =   1260
            Y2              =   1260
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Worn"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   19
            Top             =   1740
            Width           =   735
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Armour"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   23
            Top             =   2460
            Width           =   735
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Weapon"
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   21
            Top             =   2100
            Width           =   735
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Cost"
            Height          =   255
            Index           =   12
            Left            =   2460
            TabIndex        =   42
            Top             =   3120
            Width           =   615
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Game Limit"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   25
            Top             =   2820
            Width           =   855
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Max Uses"
            Height          =   255
            Index           =   14
            Left            =   240
            TabIndex        =   35
            Top             =   4620
            Width           =   735
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Weight"
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   27
            Top             =   3180
            Width           =   735
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Min/Max Hit"
            Height          =   255
            Index           =   17
            Left            =   2040
            TabIndex        =   45
            Top             =   3540
            Width           =   1035
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Accuracy"
            Height          =   255
            Index           =   18
            Left            =   240
            TabIndex        =   33
            Top             =   4260
            Width           =   735
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Speed"
            Height          =   255
            Index           =   19
            Left            =   240
            TabIndex        =   29
            Top             =   3540
            Width           =   735
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "AC/DR"
            Height          =   255
            Index           =   20
            Left            =   2340
            TabIndex        =   49
            Top             =   3900
            Width           =   735
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            Caption         =   "Req STR"
            Height          =   255
            Index           =   21
            Left            =   240
            TabIndex        =   31
            Top             =   3900
            Width           =   735
         End
      End
      Begin exlimiter.EL EL1 
         Left            =   5100
         Top             =   120
         _ExtentX        =   1270
         _ExtentY        =   1270
      End
   End
   Begin VB.TextBox txtNumberSearch 
      Height          =   285
      Left            =   60
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   2235
   End
   Begin MSComctlLib.ListView lvDatabase 
      Height          =   5175
      Left            =   60
      TabIndex        =   4
      Top             =   900
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   9128
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
      Top             =   420
      Width           =   615
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
      Left            =   720
      TabIndex        =   1
      Top             =   420
      Width           =   1875
   End
End
Attribute VB_Name = "frmItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim bvPWarned As Boolean
Dim bLoaded As Boolean
Dim nCurrentRecord As Long

Private Sub Form_Load()
Dim sCaption As String
On Error Resume Next
bLoaded = False

sCaption = frmMain.Caption
frmMain.Caption = sCaption & " - Loading Items ..."
DoEvents

With EL1
    .FormInQuestion = Me
    .MINHEIGHT = 440
    .MINWIDTH = 610
    .CenterOnLoad = False
    .EnableLimiter = True
End With

Me.Top = ReadINI("Windows", "ItemTop")
Me.Left = ReadINI("Windows", "ItemLeft")
Me.Width = ReadINI("Windows", "ItemWidth")
Me.Height = ReadINI("Windows", "ItemHeight")

lvDatabase.ListItems.clear

Dim i&, j%

For j = 0 To 9
    cmbRace(j).clear
    cmbFilter(9).clear
Next j
For j = 0 To 9
    cmbClass(j).clear
    cmbFilter(4).clear
Next j
cmbFilter(4).clear

For i = 0 To UBound(Races)
    For j = 0 To 9
        cmbRace(j).AddItem Races(i).Name
    Next j
    If Not LCase(Races(i).Name) = "none" Then
        cmbFilter(9).AddItem Races(i).Name
        cmbFilter(9).ItemData(cmbFilter(9).NewIndex) = i
    End If
Next i
For i = 0 To UBound(Classes)
    For j = 0 To 9
        cmbClass(j).AddItem Classes(i).Name
    Next j
    If Not LCase(Classes(i).Name) = "none" Then
        cmbFilter(4).AddItem Classes(i).Name
        cmbFilter(4).ItemData(cmbFilter(4).NewIndex) = i
    End If
Next i

cmbFilter(9).AddItem "Any", 0
cmbFilter(9).ItemData(cmbFilter(9).NewIndex) = 0
cmbFilter(4).AddItem "Any", 0
cmbFilter(4).ItemData(cmbFilter(4).NewIndex) = 0

Call LoadAbilities

For j = 0 To 9
    cmbFilter(j).ListIndex = 0
    Call AutoSizeDropDownWidth(cmbFilter(j))
    Call ExpandCombo(cmbFilter(j), HeightOnly, TripleWidth, fraFilter2.hwnd)
Next j

For j = 5 To 9
    cmbFilterAbilityGL(j).ListIndex = 0
Next j

Call LoadItems

Me.Show
Me.SetFocus
txtSearch.SetFocus
frmMain.Caption = sCaption
If ReadINI("Windows", "ItemMaxed") = "1" Then Me.WindowState = vbMaximized
End Sub

Private Sub chkFilterNone_Click()
On Error Resume Next
Dim x As Integer, bAction As Boolean

If chkFilterNone.Value = 0 Then bAction = True

For x = 0 To 14
    chkFilter(x).Enabled = bAction
    If x <= 9 Then cmbFilter(x).Enabled = bAction
Next x

For x = 5 To 9
    cmbFilterAbilityGL(x).Enabled = bAction
    txtFilterAbilityValue(x).Enabled = bAction
Next x

txtFilterMessage.Enabled = bAction
txtFilterTB.Enabled = bAction
txtFilterSpell.Enabled = bAction

End Sub

Private Sub cmdAbilsClear_Click(Index As Integer)
Dim x As Integer
On Error GoTo error:

For x = 0 To 19
    txtAbilityA(x).Text = 0
    txtAbilityB(x).Text = 0
Next x

out:
Exit Sub
error:
Call HandleError("cmdAbilsClear_Click")
Resume out:

End Sub

Private Sub cmdFilter_Click()
'On Error GoTo error:

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
'Dim nStatus As Integer
'
'If bLoaded Then Call saverecord(nCurrentRecord)
'
'nStatus = BTRCALL(BGETFIRST, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
'If Not nStatus = 0 Then
'    MsgBox "LoadItem, BGETFIRST, Item, Error: " & BtrieveErrorCode(nStatus)
'    Exit Sub
'End If
'
'lvDatabase.ListItems.clear
'
'If cmbFilter.ListIndex = 0 Then
'    Call LoadItems
'    Exit Sub
'End If
'
'Do While nStatus = 0
'    ItemRowToStruct Itemdatabuf.buf
'
'    If Itemrec.Type = cmbFilter.ListIndex - 1 Then Call AddItemToLV(Itemrec.Number)
'
'    nStatus = BTRCALL(BGETNEXT, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
'Loop
'If Not nStatus = 0 And Not nStatus = 9 Then
'    MsgBox "LoadItems, Error: " & BtrieveErrorCode(nStatus)
'End If
'
'If lvDatabase.ListItems.Count >= 1 Then Call lvDatabase_ItemClick(lvDatabase.ListItems(1))
'
'lvDatabase.refresh
'SortListView lvDatabase, 1, ldtNumber, True
'bLoaded = True
'
'Exit Sub
'error:
'Call HandleError
End Sub

Private Sub cmdFilterApply_Click()
Dim nStatus As Integer, bAdd As Boolean, x As Integer, bFiltered As Boolean
Dim z As Integer, bAbilMatch(5 To 7) As Boolean, nVal As Long

On Error GoTo error:

If bLoaded Then Call saverecord(nCurrentRecord)

nStatus = BTRCALL(BGETFIRST, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "BGETFIRST, Item, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

Me.MousePointer = vbHourglass

bLoaded = False
lvDatabase.ListItems.clear

If chkFilterNone.Value = 1 Then
    Call LoadItems
    Call cmdFilter_Click
    cmdFilter.Caption = "Filter"
    GoTo out:
End If

Do While nStatus = 0
    bAdd = True
    ItemRowToStruct Itemdatabuf.buf
    
    If chkFilter(0).Value = 1 And bAdd Then 'type
        If Not Itemrec.Type = cmbFilter(0).ListIndex Then bAdd = False
    End If
    If chkFilter(1).Value = 1 And bAdd Then 'worn
        If Not Itemrec.WornOn = cmbFilter(1).ListIndex Then bAdd = False
    End If
    If chkFilter(2).Value = 1 And bAdd Then 'weapon
        If Not Itemrec.Weapon = cmbFilter(2).ListIndex Then bAdd = False
    End If
    If chkFilter(3).Value = 1 And bAdd Then 'armour
        If Not Itemrec.Armour = cmbFilter(3).ListIndex Then bAdd = False
    End If
    If chkFilter(4).Value = 1 And bAdd Then 'class
        For x = 0 To 9
            If Itemrec.Class(x) = cmbFilter(4).ItemData(cmbFilter(4).ListIndex) Then Exit For
        Next x
        If x = 10 Then bAdd = False
        If bAdd Then
            For x = 0 To 9
                If Itemrec.Race(x) = cmbFilter(9).ItemData(cmbFilter(9).ListIndex) Then Exit For
            Next x
            If x = 10 Then bAdd = False
        End If
    End If
    If (chkFilter(5).Value = 1 Or chkFilter(6).Value = 1 Or chkFilter(7).Value = 1) And bAdd Then 'ability
        If chkFilter(5).Value = 1 Then bAbilMatch(5) = False Else bAbilMatch(5) = True
        If chkFilter(6).Value = 1 Then bAbilMatch(6) = False Else bAbilMatch(6) = True
        If chkFilter(7).Value = 1 Then bAbilMatch(7) = False Else bAbilMatch(7) = True
        
        For x = 0 To 19
            For z = 5 To 7
                If chkFilter(z).Value = 1 Then
                    If Itemrec.AbilityA(x) = cmbFilter(z).ItemData(cmbFilter(z).ListIndex) Then
                        nVal = Val(txtFilterAbilityValue(z).Text)
                        bAbilMatch(z) = True
                        
                        If cmbFilterAbilityGL(z).ListIndex = 0 Then  'ANY
                        ElseIf cmbFilterAbilityGL(z).ListIndex = 1 Then  '<=
                            If Itemrec.AbilityB(x) > nVal Then bAbilMatch(z) = False
                            If chkFilterExcludeZero.Value = 1 And Itemrec.AbilityB(x) = 0 Then bAbilMatch(z) = False
                        ElseIf cmbFilterAbilityGL(z).ListIndex = 2 Then  '>=
                            If Itemrec.AbilityB(x) < nVal Then bAbilMatch(z) = False
                        ElseIf cmbFilterAbilityGL(z).ListIndex = 3 Then  '=
                            If Not Itemrec.AbilityB(x) = nVal Then bAbilMatch(z) = False
                        End If
                        
                        If Not bAbilMatch(z) Then GoTo abil_out:
                    End If
                End If
            Next z
        Next x
abil_out:
        If Not (bAbilMatch(5) And bAbilMatch(6) And bAbilMatch(7)) Then bAdd = False
    End If
    
    If chkFilter(8).Value = 1 And bAdd Then 'cost
        If Itemrec.CostType = cmbFilter(8).ListIndex Then
            If cmbFilterAbilityGL(8).ListIndex = 0 Then 'ANY
            ElseIf cmbFilterAbilityGL(8).ListIndex = 1 Then '<=
                If Itemrec.Cost > Val(txtFilterAbilityValue(8).Text) Then bAdd = False
                If chkFilterExcludeZero.Value = 1 And Itemrec.Cost = 0 Then bAdd = False
            ElseIf cmbFilterAbilityGL(8).ListIndex = 2 Then '>=
                If Itemrec.Cost < Val(txtFilterAbilityValue(8).Text) Then bAdd = False
            Else '=
                If Not Itemrec.Cost = Val(txtFilterAbilityValue(8).Text) Then bAdd = False
            End If
        Else
            bAdd = False
        End If
    End If
    If chkFilter(9).Value = 1 And bAdd Then 'limit>0
        If Not Itemrec.GameLimit > 0 Then bAdd = False
    End If
    If chkFilter(10).Value = 1 And bAdd Then 'uses>0
        If Not Itemrec.Uses > 0 Then bAdd = False
    End If
    If chkFilter(11).Value = 1 And bAdd Then 'non-magic
        For x = 0 To 19
            If Itemrec.AbilityA(x) = 28 Then
                bAdd = False
                Exit For
            End If
        Next x
    End If
    If chkFilter(12).Value = 1 And bAdd Then 'message
        nVal = Val(txtFilterMessage.Text)
        If Itemrec.HitMsg = nVal Then GoTo msg_match:
        If Itemrec.MissMsg = nVal Then GoTo msg_match:
        If Itemrec.DistructMsg = nVal Then GoTo msg_match:
        bAdd = False
msg_match:
    End If
    If chkFilter(13).Value = 1 And bAdd Then 'textblock
        If Not Itemrec.ReadTB = Val(txtFilterTB.Text) Then bAdd = False
    End If
    If chkFilter(14).Value = 1 And bAdd Then 'spell
        nVal = Val(txtFilterSpell.Text)
        For x = 0 To 9
            Select Case Itemrec.AbilityA(x)
                Case 42, 43:
                    If Itemrec.AbilityB(x) = nVal Then GoTo spell_match:
            End Select
        Next x
        bAdd = False
spell_match:
    End If
    
    If bAdd Then
        Call AddItemToLV(Itemrec.Number)
    Else
        bFiltered = True
    End If
    
    nStatus = BTRCALL(BGETNEXT, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
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

Private Sub cmdFilterCancel_Click()
    fraFilter.Visible = False
    txtNumberSearch.Enabled = True
    framNav.Enabled = True
    lvDatabase.Enabled = True
    txtSearch.Enabled = True
End Sub

Private Sub cmdFilterReset_Click()
Dim x As Integer

For x = 0 To 14
    chkFilter(x).Value = 0
    If x <= 8 Then cmbFilter(x).ListIndex = 0
Next x

For x = 5 To 8
    cmbFilterAbilityGL(x).ListIndex = 0
    txtFilterAbilityValue(x).Text = 0
Next x

txtFilterSpell.Text = 0
txtFilterTB.Text = 0
txtFilterMessage.Text = 0

End Sub

Private Sub cmdResetClass_Click()
Dim x As Integer
On Error Resume Next

For x = 0 To 9
    cmbClass(x).ListIndex = 0
Next x
End Sub

Private Sub cmdResetNegate_Click()
Dim x As Integer
On Error Resume Next

For x = 0 To 9
    txtNegate(x).Text = "0"
Next x
End Sub

Private Sub cmdResetRace_Click()
Dim x As Integer
On Error Resume Next

For x = 0 To 9
    cmbRace(x).ListIndex = 0
Next x

End Sub



Private Sub LoadAbilities()
Dim x As Integer
On Error GoTo error:

For x = 5 To 7
    cmbFilter(x).clear
Next x
rsAbilities.MoveFirst
Do Until rsAbilities.EOF
    If Not rsAbilities.Fields("Number") = 0 Then
        For x = 5 To 7
            cmbFilter(x).AddItem rsAbilities.Fields("Name") & " (" & rsAbilities.Fields("Number") & ")"
            cmbFilter(x).ItemData(cmbFilter(x).NewIndex) = rsAbilities.Fields("Number")
        Next x
    End If
    rsAbilities.MoveNext
Loop

For x = 5 To 7
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



Private Sub cmdCalcSwings_Click()
If lvDatabase.ListItems.Count < 1 Then Exit Sub

If FormIsLoaded("frmSwingCalc") Then
    Call frmSwingCalc.LoadWeapons
Else
    Load frmSwingCalc
End If

frmSwingCalc.GotoWeapon (Val(lvDatabase.SelectedItem.Text))
frmSwingCalc.Show
frmSwingCalc.SetFocus

End Sub

Private Sub cmdDiscard_Click()
Dim nStatus As Integer

If lvDatabase.SelectedItem Is Nothing Or nCurrentRecord = 0 Then
    MsgBox "No current record."
    Exit Sub
End If

nStatus = BTRCALL(BGETEQUAL, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), nCurrentRecord, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error on BGETEQUAL: " & BtrieveErrorCode(nStatus)
Else
    DispItemInfo Itemdatabuf.buf
End If

End Sub

Private Sub cmdEditDestructMsg_Click()
Call frmMessage.GotoMSG(Val(txtDistructMsg.Text))
End Sub

Private Sub cmdEditHitMsg_Click()
Call frmMessage.GotoMSG(Val(txtHitMsg.Text))
End Sub

Private Sub cmdEditMissMsg_Click()
Call frmMessage.GotoMSG(Val(txtMissMsg.Text))
End Sub

Private Sub cmdEditNegate_Click(Index As Integer)
Call frmSpell.GotoSpell(Val(txtNegate(Index).Text))
frmSpell.Show
frmSpell.SetFocus
End Sub
Public Sub GotoItem(ByVal nRecnum As Long)
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
Private Sub cmdEditReadMsg_Click()
    Call frmTextblock.GotoTB(Val(txtReadMsg.Text))
'    DoEvents
'    Call frmTextblock.cmdShowPreview_Click
End Sub

Private Sub cmdSave_Click()
On Error GoTo error:

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
If lvDatabase.SelectedItem Is Nothing Then Exit Sub

Call saverecord(nCurrentRecord)
'Call lvDatabase_ItemClick(lvDatabase.SelectedItem)

Dim oLI As ListItem
Set oLI = lvDatabase.FindItem(Itemrec.Number, lvwText, , 0)
If Not oLI Is Nothing Then
    oLI.ListSubItems(1).Text = ClipNull(Itemrec.Name)
    If Not bOnlyNames Then
        oLI.ListSubItems(2).Text = GetItemType(Itemrec.Type)
        oLI.ListSubItems(3).Text = (Itemrec.AC / 10) & "/" & (Itemrec.DR / 10)
        oLI.ListSubItems(4).Text = GetWornType(Itemrec.WornOn)
        oLI.ListSubItems(5).Text = GetArmourType(Itemrec.Armour)
        oLI.ListSubItems(6).Text = GetWeaponType(Itemrec.Weapon)
        oLI.ListSubItems(7).Text = Itemrec.GameLimit
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
lvDatabase.Height = Me.Height - 1385 - TITLEBAR_OFFSET
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
    Case 1, 4, 8: nSort = ldtNumber
    Case Else: nSort = ldtString
End Select
SortListView lvDatabase, ColumnHeader.Index, nSort, lvDatabase.SortOrder
End Sub

Private Sub lvDatabase_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim temp As Long, nStatus As Integer

If bLoaded = True And chkAutoSave.Value = 1 Then Call saverecord(nCurrentRecord)

temp = Val(Item.Text)
nStatus = BTRCALL(BGETEQUAL, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), temp, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error on BGETEQUAL: " & BtrieveErrorCode(nStatus)
    bLoaded = False
Else
    nCurrentRecord = temp
    DispItemInfo Itemdatabuf.buf
    bLoaded = True
End If
End Sub

Private Sub Text1_GotFocus()
Call SelectAll(Text1)

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
Private Sub LoadItems()
On Error GoTo error:
Dim nStatus As Integer

lvDatabase.ColumnHeaders.clear
lvDatabase.ColumnHeaders.add 1, "Number", "#", 600, lvwColumnLeft
lvDatabase.ColumnHeaders.add 2, "Name", "Name", 1900, lvwColumnCenter
If Not bOnlyNames Then
    lvDatabase.ColumnHeaders.add 3, "Type", "Type", 1000, lvwColumnCenter
    lvDatabase.ColumnHeaders.add 4, "AC/DR", "AC/DR", 1000, lvwColumnCenter
    lvDatabase.ColumnHeaders.add 5, "Worn", "Worn", 1000, lvwColumnCenter
    lvDatabase.ColumnHeaders.add 6, "Armour", "Armour", 1000, lvwColumnCenter
    lvDatabase.ColumnHeaders.add 7, "Weapon", "Weapon", 1000, lvwColumnCenter
    lvDatabase.ColumnHeaders.add 8, "Limit", "Limit", 700, lvwColumnCenter
End If

nStatus = BTRCALL(BGETFIRST, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "LoadItem, BGETFIRST, Item, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

Do While nStatus = 0
    ItemRowToStruct Itemdatabuf.buf

    Call AddItemToLV(Itemrec.Number)

    nStatus = BTRCALL(BGETNEXT, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "LoadItems, Error: " & BtrieveErrorCode(nStatus)
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
''''On Error GoTo error:
''''Dim nStatus As Integer, oLI As ListItem
''''
''''lvDatabase.ColumnHeaders.clear
''''lvDatabase.ColumnHeaders.add 1, "Number", "#", 600, lvwColumnLeft
''''lvDatabase.ColumnHeaders.add 2, "Name", "Name", 1900, lvwColumnCenter
''''lvDatabase.ColumnHeaders.add 3, "Type", "Type", 1000, lvwColumnCenter
''''lvDatabase.ColumnHeaders.add 4, "AC/DR", "AC/DR", 1000, lvwColumnCenter
''''lvDatabase.ColumnHeaders.add 5, "Worn", "Worn", 1000, lvwColumnCenter
''''lvDatabase.ColumnHeaders.add 6, "Armour", "Armour", 1000, lvwColumnCenter
''''lvDatabase.ColumnHeaders.add 7, "Weapon", "Weapon", 1000, lvwColumnCenter
''''lvDatabase.ColumnHeaders.add 8, "Limit", "Limit", 700, lvwColumnCenter
''''
''''nStatus = BTRCALL(BGETFIRST, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
''''If Not nStatus = 0 Then
''''    MsgBox "LoadItem, BGETFIRST, Item, Error: " & BtrieveErrorCode(nStatus)
''''    Exit Sub
''''End If
''''
'''''GetNextExRec.HeaderDataBufLength = 35
''''Call SetBinaryValue(48, GetNextExDataBuf.buf(), 1, 2)
'''''GetNextExRec.HeaderBeginCode = "UC"
''''Call SetBinaryString("UC", GetNextExDataBuf.buf(), 3, 2)
'''''GetNextExRec.HeaderMaxReject = 20000
''''Call SetBinaryValue(65535, GetNextExDataBuf.buf(), 5, 2)
'''''GetNextExRec.HeaderNumFilterTerms = 1
''''Call SetBinaryValue(0, GetNextExDataBuf.buf(), 7, 2)
''''
'''''GetNextExRec.HeaderNumRecordsReturned = 1
''''Call SetBinaryValue(1, GetNextExDataBuf.buf(), 9, 2)
'''''GetNextExRec.HeaderNumFieldsExtracted = 3
''''Call SetBinaryValue(9, GetNextExDataBuf.buf(), 11, 2)
''''
'''''GetNextExRec.Field1Length
'''''GetNextExRec.Field1Offset
''''Call SetBinaryValue(4, GetNextExDataBuf.buf(), 13, 2) 'number
''''Call SetBinaryValue(0, GetNextExDataBuf.buf(), 15, 2)
''''
''''Call SetBinaryValue(2, GetNextExDataBuf.buf(), 17, 2) 'limit
''''Call SetBinaryValue(6, GetNextExDataBuf.buf(), 19, 2)
''''
''''Call SetBinaryValue(29, GetNextExDataBuf.buf(), 21, 2) 'name
''''Call SetBinaryValue(173, GetNextExDataBuf.buf(), 23, 2)
''''
''''Call SetBinaryValue(2, GetNextExDataBuf.buf(), 25, 2) 'type
''''Call SetBinaryValue(756, GetNextExDataBuf.buf(), 27, 2)
''''
''''Call SetBinaryValue(2, GetNextExDataBuf.buf(), 29, 2) 'ac
''''Call SetBinaryValue(834, GetNextExDataBuf.buf(), 31, 2)
''''
''''Call SetBinaryValue(2, GetNextExDataBuf.buf(), 33, 2) 'dr
''''Call SetBinaryValue(924, GetNextExDataBuf.buf(), 35, 2)
''''
''''Call SetBinaryValue(2, GetNextExDataBuf.buf(), 37, 2) 'weapon
''''Call SetBinaryValue(916, GetNextExDataBuf.buf(), 39, 2)
''''
''''Call SetBinaryValue(2, GetNextExDataBuf.buf(), 41, 2) 'armour
''''Call SetBinaryValue(918, GetNextExDataBuf.buf(), 43, 2)
''''
''''Call SetBinaryValue(2, GetNextExDataBuf.buf(), 45, 2) 'worn
''''Call SetBinaryValue(920, GetNextExDataBuf.buf(), 47, 2)
''''
''''nStatus = BTRCALL(BGETNEXTEXTENDED, ItemPosBlock, GetNextExDataBuf, Len(GetNextExDataBuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
''''If Not nStatus = 0 Then
''''    MsgBox "LoadItem, BGETFIRST, Item, Error: " & BtrieveErrorCode(nStatus)
''''    Exit Sub
''''End If
''''
''''Do While nStatus = 0
''''    'ItemRowToStruct Itemdatabuf.buf
''''
''''    Set oLI = lvDatabase.ListItems.add()
''''    oLI.Text = GetBinaryValue(GetNextExDataBuf.buf(), 9, 4)
''''
''''    oLI.ListSubItems.add (1), "Name", ClipNull(GetBinaryString(GetNextExDataBuf.buf(), 15, 29))
''''    oLI.ListSubItems.add (2), "Type", GetItemType(GetBinaryValue(GetNextExDataBuf.buf(), 44, 2))
''''    oLI.ListSubItems.add (3), "AC/DR", (GetBinaryValue(GetNextExDataBuf.buf(), 46, 2) / 10) _
''''        & "/" & (GetBinaryValue(GetNextExDataBuf.buf(), 48, 2) / 10)
''''    oLI.ListSubItems.add (4), "Worn", GetWornType(GetBinaryValue(GetNextExDataBuf.buf(), 54, 2))
''''    oLI.ListSubItems.add (5), "Armour", GetArmourType(GetBinaryValue(GetNextExDataBuf.buf(), 52, 2))
''''    oLI.ListSubItems.add (6), "Weapon", GetWeaponType(GetBinaryValue(GetNextExDataBuf.buf(), 50, 2))
''''    oLI.ListSubItems.add (7), "Limit", GetBinaryValue(GetNextExDataBuf.buf(), 13, 2)
''''
''''    Call SetBinaryValue(48, GetNextExDataBuf.buf(), 1, 2)
''''    Call SetBinaryString("EG", GetNextExDataBuf.buf(), 3, 2)
''''    Call SetBinaryValue(65535, GetNextExDataBuf.buf(), 5, 2)
''''    Call SetBinaryValue(0, GetNextExDataBuf.buf(), 7, 2)
''''
''''    Call SetBinaryValue(1, GetNextExDataBuf.buf(), 9, 2)
''''    Call SetBinaryValue(9, GetNextExDataBuf.buf(), 11, 2)
''''
''''    Call SetBinaryValue(4, GetNextExDataBuf.buf(), 13, 2) 'number +0
''''    Call SetBinaryValue(0, GetNextExDataBuf.buf(), 15, 2)
''''    Call SetBinaryValue(2, GetNextExDataBuf.buf(), 17, 2) 'limit +4
''''    Call SetBinaryValue(6, GetNextExDataBuf.buf(), 19, 2)
''''    Call SetBinaryValue(29, GetNextExDataBuf.buf(), 21, 2) 'name +6
''''    Call SetBinaryValue(173, GetNextExDataBuf.buf(), 23, 2)
''''    Call SetBinaryValue(2, GetNextExDataBuf.buf(), 25, 2) 'type +35
''''    Call SetBinaryValue(756, GetNextExDataBuf.buf(), 27, 2)
''''    Call SetBinaryValue(2, GetNextExDataBuf.buf(), 29, 2) 'ac +37
''''    Call SetBinaryValue(834, GetNextExDataBuf.buf(), 31, 2)
''''    Call SetBinaryValue(2, GetNextExDataBuf.buf(), 33, 2) 'dr +39
''''    Call SetBinaryValue(924, GetNextExDataBuf.buf(), 35, 2)
''''    Call SetBinaryValue(2, GetNextExDataBuf.buf(), 37, 2) 'weapon +41
''''    Call SetBinaryValue(916, GetNextExDataBuf.buf(), 39, 2)
''''    Call SetBinaryValue(2, GetNextExDataBuf.buf(), 41, 2) 'armour +43
''''    Call SetBinaryValue(918, GetNextExDataBuf.buf(), 43, 2)
''''    Call SetBinaryValue(2, GetNextExDataBuf.buf(), 45, 2) 'worn +45
''''    Call SetBinaryValue(920, GetNextExDataBuf.buf(), 47, 2)
''''
''''    Set oLI = Nothing
''''    nStatus = BTRCALL(BGETNEXTEXTENDED, ItemPosBlock, GetNextExDataBuf, Len(GetNextExDataBuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
''''Loop
''''If Not nStatus = 0 And Not nStatus = 9 Then
''''    MsgBox "LoadItems, Error: " & BtrieveErrorCode(nStatus)
''''End If
''''
''''If lvDatabase.ListItems.Count >= 1 Then Call lvDatabase_ItemClick(lvDatabase.ListItems(1))
''''
''''lvDatabase.refresh
''''SortListView lvDatabase, 1, ldtNumber, True
''''bLoaded = True
''''Set oLI = Nothing
''''
''''Exit Sub
''''error:
''''Call HandleError
''''Set oLI = Nothing
End Sub
Private Sub AddItemToLV(ByVal nNumber As Long)
Dim nStatus As Integer, oLI As ListItem
On Error GoTo error:

If Not nNumber = Itemrec.Number Then
    nStatus = BTRCALL(BGETEQUAL, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), nNumber, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then MsgBox "Error getting record " & nNumber & ": " & BtrieveErrorCode(nStatus)
    bLoaded = False
    Exit Sub
End If

Set oLI = lvDatabase.ListItems.add()
oLI.Text = Itemrec.Number

oLI.ListSubItems.add (1), "Name", ClipNull(Itemrec.Name)
If Not bOnlyNames Then
    oLI.ListSubItems.add (2), "Type", GetItemType(Itemrec.Type)
    oLI.ListSubItems.add (3), "AC/DR", (Itemrec.AC / 10) & "/" & (Itemrec.DR / 10)
    oLI.ListSubItems.add (4), "Worn", GetWornType(Itemrec.WornOn)
    oLI.ListSubItems.add (5), "Armour", GetArmourType(Itemrec.Armour)
    oLI.ListSubItems.add (6), "Weapon", GetWeaponType(Itemrec.Weapon)
    oLI.ListSubItems.add (7), "Limit", Itemrec.GameLimit
End If

Set oLI = Nothing
Exit Sub
error:
Call HandleError
Set oLI = Nothing
End Sub

Private Sub DispItemInfo(row() As Byte)
On Error GoTo error:
Dim x As Integer, j As Integer, i As Integer
bLoaded = True

Call ItemRowToStruct(row())

Me.Caption = "Item Editor -- " & ClipNull(Itemrec.Name)

txtNumber.Text = Itemrec.Number
txtName.Text = Itemrec.Name
txtDesc(0).Text = Itemrec.Desc1
txtDesc(1).Text = Itemrec.Desc2
txtDesc(2).Text = Itemrec.Desc3
txtDesc(3).Text = Itemrec.Desc4
txtDesc(4).Text = Itemrec.Desc5
txtDesc(5).Text = Itemrec.Desc6
txtLimit.Text = Itemrec.GameLimit
txtWeight.Text = Itemrec.Weight
cmbType.ListIndex = Itemrec.Type
If Itemrec.Type = 1 Then
    cmdCalcSwings.Enabled = True
Else
    cmdCalcSwings.Enabled = False
End If
txtUses.Text = Itemrec.Uses
txtCost.Text = Itemrec.Cost
chkRobable.Value = Itemrec.Robable
cmbCostType.ListIndex = Itemrec.CostType
chkDestroy.Value = Itemrec.DestroyOnDeath
chkRetainAfterUses.Value = Itemrec.RetainAfterUses
txtMinHit.Text = Itemrec.Minhit
txtMaxHit.Text = Itemrec.Maxhit
cmbWeapon.ListIndex = Itemrec.Weapon
cmbArmour.ListIndex = Itemrec.Armour
cmbWornOn.ListIndex = Itemrec.WornOn
txtAccuracy.Text = Itemrec.Accuracy
txtAC.Text = Itemrec.AC
txtDR.Text = Itemrec.DR
lblACDR.Caption = "(" & Round(Itemrec.AC / 10, 1) & "/" & Round(Itemrec.DR / 10, 1) & ")"
chkGettable.Value = Itemrec.Gettable
txtReqSTR.Text = Itemrec.ReqStr
txtRunic.Text = Itemrec.OpenRunic
txtPlatinum.Text = Itemrec.OpenPlatinum
txtGold.Text = Itemrec.OpenGold
txtSilver.Text = Itemrec.OpenSilver
txtCopper.Text = Itemrec.OpenCopper
txtSpeed.Text = Itemrec.Speed
txtMissMsg.Text = Itemrec.MissMsg
'txtMissMsgDisplay.Text = GetMessages(Itemrec.MissMsg, 1)
txtHitMsg.Text = Itemrec.HitMsg
'txtHitMsgDisplay.Text = GetMessages(Itemrec.HitMsg, 1)
txtDistructMsg.Text = Itemrec.DistructMsg
'txtDistructMsgDisplay.Text = GetMessages(Itemrec.DistructMsg, 1)
txtReadMsg.Text = Itemrec.ReadTB
'txtReadMsgDisplay.Text = GetTextblock(Itemrec.ReadTB)
chkNotDroppable.Value = Itemrec.NotDroppable

txtUnknown1.Text = Itemrec.unknown1
txtUnknown8.Text = Itemrec.unknown8
Text1.Text = Itemrec.unknown7

For x = 0 To 9
    
    If cmbRace(x).ListCount <= Itemrec.Race(x) Then
        Call Add2RaceArray(Itemrec.Race(x))
        For j = 0 To 9
            cmbRace(j).clear
        Next j
        For i = 0 To UBound(Races)
            For j = 0 To 9
                cmbRace(j).AddItem Races(i).Name
            Next j
        Next i
    End If
        
    If cmbClass(x).ListCount <= Itemrec.Class(x) Then
        Call Add2ClassArray(Itemrec.Class(x))
        For j = 0 To 9
            cmbClass(j).clear
        Next j
        For i = 0 To UBound(Classes)
            For j = 0 To 9
                cmbClass(j).AddItem Classes(i).Name
            Next j
        Next i
    End If
    
    cmbClass(x).ListIndex = Itemrec.Class(x)
    cmbRace(x).ListIndex = Itemrec.Race(x)
    txtNegate(x).Text = Itemrec.Negate(x * 2)
    'txtNegateName(x).Text = GetSpellName(Itemrec.Negate(x * 2))

Next

For x = 0 To 19
    txtAbilityA(x).Text = Itemrec.AbilityA(x)
    txtAbilityB(x).Text = Itemrec.AbilityB(x)
Next

Exit Sub
error:
Call HandleError
MsgBox "Warning, record was not completely displayed." & vbCrLf _
    & "Previous records stats may still be in memory.  Select 'Disable DB Writing'" & vbCrLf _
    & "from the file menu and then reload the editor.", vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
        'Set TTtxtBox = Nothing
        If bLoaded = True Then Call saverecord(nCurrentRecord)
        If Me.WindowState = vbMinimized Then Exit Sub
        
        If Me.WindowState = vbMaximized Then
            Call WriteINI("Windows", "ItemMaxed", 1)
        Else
            Call WriteINI("Windows", "ItemMaxed", 0)
            Call WriteINI("Windows", "ItemTop", Me.Top)
            Call WriteINI("Windows", "ItemLeft", Me.Left)
            Call WriteINI("Windows", "ItemWidth", Me.Width)
            Call WriteINI("Windows", "ItemHeight", Me.Height)
        End If
End Sub

Private Sub saverecord(ByVal nRecord As Long)
On Error GoTo error:
Dim nStatus As Integer, x As Integer

If nRecord = 0 Then Exit Sub

nStatus = BTRCALL(BGETEQUAL, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), nRecord, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Save Error on BGETEQUAL: " & BtrieveErrorCode(nStatus)
    Exit Sub
Else
    ItemRowToStruct Itemdatabuf.buf
End If

'DoEvents
Itemrec.Name = RTrim(txtName.Text) & Chr(0)
Itemrec.Desc1 = Trim(txtDesc(0).Text) & Chr(0)
Itemrec.Desc2 = Trim(txtDesc(1).Text) & Chr(0)
Itemrec.Desc3 = Trim(txtDesc(2).Text) & Chr(0)
Itemrec.Desc4 = Trim(txtDesc(3).Text) & Chr(0)
Itemrec.Desc5 = Trim(txtDesc(4).Text) & Chr(0)
Itemrec.Desc6 = Trim(txtDesc(5).Text) & Chr(0)
Itemrec.GameLimit = Val(txtLimit.Text)
Itemrec.Weight = Val(txtWeight.Text)
Itemrec.Type = cmbType.ListIndex
Itemrec.Uses = Val(txtUses.Text)
Itemrec.Cost = Val(txtCost.Text)
Itemrec.CostType = cmbCostType.ListIndex
Itemrec.DestroyOnDeath = chkDestroy.Value
Itemrec.Minhit = Val(txtMinHit.Text)
Itemrec.Maxhit = Val(txtMaxHit.Text)
Itemrec.AC = Val(txtAC.Text)
Itemrec.Weapon = cmbWeapon.ListIndex
Itemrec.Armour = cmbArmour.ListIndex
Itemrec.WornOn = cmbWornOn.ListIndex
Itemrec.Accuracy = Val(txtAccuracy.Text)
Itemrec.DR = Val(txtDR.Text)
Itemrec.Gettable = chkGettable.Value
Itemrec.ReqStr = Val(txtReqSTR.Text)

Itemrec.OpenRunic = Val(txtRunic.Text)
Itemrec.OpenPlatinum = Val(txtPlatinum.Text)
Itemrec.OpenGold = Val(txtGold.Text)
Itemrec.OpenSilver = Val(txtSilver.Text)
Itemrec.OpenCopper = Val(txtCopper.Text)

Itemrec.Speed = Val(txtSpeed.Text)
Itemrec.MissMsg = Val(txtMissMsg.Text)
Itemrec.HitMsg = Val(txtHitMsg.Text)
Itemrec.DistructMsg = ULong2SLong(Val(txtDistructMsg.Text))
Itemrec.ReadTB = Val(txtReadMsg.Text)
Itemrec.NotDroppable = chkNotDroppable.Value
Itemrec.RetainAfterUses = chkRetainAfterUses.Value

Itemrec.unknown1 = Val(txtUnknown1.Text)
Itemrec.unknown8 = Val(txtUnknown8.Text)
Itemrec.Robable = chkRobable.Value

For x = 0 To 9
    Itemrec.Class(x) = cmbClass(x).ListIndex
    Itemrec.Race(x) = cmbRace(x).ListIndex
    Itemrec.Negate(x * 2) = Val(txtNegate(x).Text)
Next

For x = 0 To 19
    Itemrec.AbilityA(x) = Val(txtAbilityA(x).Text)
    Itemrec.AbilityB(x) = Val(txtAbilityB(x).Text)
Next

nStatus = UpdateItem
If Not nStatus = 0 Then
    MsgBox "SaveRecord, BUPDATE: " & BtrieveErrorCode(nStatus)
Else
    DispItemInfo Itemdatabuf.buf
End If

Exit Sub
error:
Call HandleError
End Sub

Private Sub txtAbilityB_GotFocus(Index As Integer)
Call SelectAll(txtAbilityB(Index))

End Sub

Private Sub txtAC_Change()
lblACDR.Caption = "(" & Round(Val(txtAC.Text) / 10, 1) & "/" & Round(Val(txtDR.Text) / 10, 1) & ")"
End Sub

Private Sub txtAC_GotFocus()
Call SelectAll(txtAC)

End Sub

Private Sub txtAccuracy_GotFocus()
Call SelectAll(txtAccuracy)

End Sub

Private Sub txtCopper_GotFocus()
Call SelectAll(txtCopper)

End Sub

Private Sub txtCost_GotFocus()
Call SelectAll(txtCost)

End Sub

Private Sub txtDesc1_Change()

End Sub

Private Sub txtDesc_Change(Index As Integer)
If Index = 5 Then Exit Sub
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

Private Sub txtDistructMsg_Change()
On Error GoTo error:

txtDistructMsgDisplay.Text = GetMessages(Val(txtDistructMsg.Text), 1)

out:
Exit Sub
error:
Call HandleError("txtDistructMsg_Change")
Resume out:
End Sub

Private Sub txtDistructMsg_GotFocus()
Call SelectAll(txtDistructMsg)

End Sub

Private Sub txtDR_Change()
lblACDR.Caption = "(" & Round(Val(txtAC.Text) / 10, 1) & "/" & Round(Val(txtDR.Text) / 10, 1) & ")"
End Sub

Private Sub txtDR_GotFocus()
Call SelectAll(txtDR)

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

Private Sub txtGold_GotFocus()
Call SelectAll(txtGold)

End Sub

Private Sub txtHitMsg_Change()
On Error GoTo error:

txtHitMsgDisplay.Text = GetMessages(Val(txtHitMsg.Text), 1)

out:
Exit Sub
error:
Call HandleError("txtHitMsg_Change")
Resume out:
End Sub

Private Sub txtHitMsg_GotFocus()
Call SelectAll(txtHitMsg)

End Sub

Private Sub txtLimit_GotFocus()
Call SelectAll(txtLimit)

End Sub

Private Sub txtMaxHit_GotFocus()
Call SelectAll(txtMaxHit)

End Sub

Private Sub txtMinHit_GotFocus()
Call SelectAll(txtMinHit)
End Sub

Private Sub txtMissMsg_Change()
On Error GoTo error:

txtMissMsgDisplay.Text = GetMessages(Val(txtMissMsg.Text), 1)

out:
Exit Sub
error:
Call HandleError("txtMissMsg_Change")
Resume out:
End Sub

Private Sub txtMissMsg_GotFocus()
Call SelectAll(txtMissMsg)

End Sub

Private Sub txtName_GotFocus()
Call SelectAll(txtName)

End Sub

Private Sub txtNegate_Change(Index As Integer)
On Error GoTo error:

txtNegateName(Index).Text = GetSpellName(Val(txtNegate(Index).Text))

out:
Exit Sub
error:
Call HandleError("txtNegate_Change")
Resume out:
End Sub

Private Sub txtNegate_GotFocus(Index As Integer)
Call SelectAll(txtNegate(Index))

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
If KeyCode = vbKeyControl Then Exit Sub
If KeyCode = 18 Then Exit Sub 'alt
If KeyCode = vbKeyTab Then Exit Sub
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

Private Sub txtReadMsg_Change()
On Error GoTo error:

txtReadMsgDisplay.Text = GetTextblock(Val(txtReadMsg.Text))

out:
Exit Sub
error:
Call HandleError("txtReadMsg_Change")
Resume out:
End Sub

Private Sub txtReadMsg_GotFocus()
Call SelectAll(txtReadMsg)

End Sub

Private Sub txtReqSTR_GotFocus()
Call SelectAll(txtReqSTR)

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

nStatus = BTRCALL(BGETEQUAL, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), nCurrentRecord, KEY_BUF_LEN, 0)
If nStatus = 0 Then
    nStatus = BTRCALL(BDELETE, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
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
Dim nNewItemNumber As String, oLI As ListItem

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

If bLoaded = True Then Call saverecord(nCurrentRecord)

nNewItemNumber = InputBox("New Item Number:" & vbCrLf & vbCrLf & "Enter 0 for the next highest number.", "Insert", "0")
If nNewItemNumber = "" Then Exit Sub

Itemrec.Number = Val(nNewItemNumber)
'Itemrec.Name = "New Item" & Chr(0)
Call ItemStructToRow(Itemdatabuf.buf)

nStatus = BTRCALL(BINSERT, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdInsert, BINSERT, Error: " & BtrieveErrorCode(nStatus)
    bLoaded = False
Else
    ItemRowToStruct Itemdatabuf.buf
    
    Call AddItemToLV(Itemrec.Number)
    
    nCurrentRecord = Itemrec.Number
    DispItemInfo Itemdatabuf.buf
    
    SortListView lvDatabase, 1, ldtNumber, True
    
    Set oLI = lvDatabase.FindItem(Itemrec.Number, lvwText, , 0)
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

Private Sub cmbWornOn_Click()
If cmbWornOn.ListIndex >= 17 And cmbWornOn.ListIndex <= 19 And _
    eDatFileVersion < v111p13 And bvPWarned = False Then
        MsgBox "Note: eyes, face, and second wrist slot work only in v1.11p-beta12+", vbInformation
        bvPWarned = True
End If
End Sub


Private Sub txtSilver_GotFocus()
Call SelectAll(txtSilver)

End Sub

Private Sub txtSpeed_GotFocus()
Call SelectAll(txtSpeed)

End Sub

Private Sub txtUnknown1_GotFocus()
Call SelectAll(txtUnknown1)

End Sub

Private Sub txtUnknown8_GotFocus()
Call SelectAll(txtUnknown8)

End Sub

Private Sub txtUses_GotFocus()
Call SelectAll(txtUses)

End Sub

Private Sub txtWeight_GotFocus()
Call SelectAll(txtWeight)

End Sub
