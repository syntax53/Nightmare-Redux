VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{20D5284F-7B23-4F0A-B8B1-6C9D18B64F1C}#1.0#0"; "exlimiter.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmUser 
   Caption         =   "User Editor"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17325
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   17325
   Begin VB.Frame framImport 
      Caption         =   "Import"
      Height          =   5295
      Left            =   10260
      TabIndex        =   579
      Top             =   60
      Visible         =   0   'False
      Width           =   6915
      Begin VB.CheckBox chkImportDeets 
         Caption         =   "Import Lives, Suicide, Channel, Cash"
         Height          =   195
         Left            =   2640
         TabIndex        =   590
         Top             =   1620
         Width           =   3195
      End
      Begin VB.CheckBox chkImportRooms 
         Caption         =   "Import Rooms"
         Height          =   195
         Left            =   2640
         TabIndex        =   589
         Top             =   1980
         Width           =   1875
      End
      Begin VB.CommandButton cmdImportGo 
         Caption         =   "Cancel"
         Height          =   435
         Index           =   1
         Left            =   2520
         TabIndex        =   588
         Top             =   4620
         Width           =   1575
      End
      Begin VB.CommandButton cmdImportGo 
         Caption         =   "Import from Clipboard"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   0
         Left            =   2520
         TabIndex        =   587
         Top             =   3780
         Width           =   1575
      End
      Begin VB.CheckBox chkImportClearAbils 
         Caption         =   "Clear Abilities First"
         Height          =   195
         Left            =   2640
         TabIndex        =   586
         Top             =   3420
         Width           =   1875
      End
      Begin VB.CheckBox chkImportClearKeys 
         Caption         =   "Clear Keys First"
         Height          =   195
         Left            =   2640
         TabIndex        =   585
         Top             =   3060
         Width           =   1875
      End
      Begin VB.CheckBox chkImportClearSpells 
         Caption         =   "Clear Spells First"
         Height          =   195
         Left            =   2640
         TabIndex        =   584
         Top             =   2340
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkImportClearItems 
         Caption         =   "Clear Items First"
         Height          =   195
         Left            =   2640
         TabIndex        =   583
         Top             =   2700
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkImportGang 
         Caption         =   "Import Gang Name"
         Height          =   195
         Left            =   2640
         TabIndex        =   582
         Top             =   1260
         Width           =   1875
      End
      Begin VB.CheckBox chkImportName 
         Caption         =   "Import Character Name"
         Height          =   195
         Left            =   2640
         TabIndex        =   580
         Top             =   900
         Width           =   2295
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Caption         =   "Character Import"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   581
         Top             =   300
         Width           =   6615
      End
   End
   Begin VB.Frame framNav 
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   3300
      TabIndex        =   3
      Top             =   60
      Width           =   6915
      Begin VB.CommandButton cmdUserImport 
         Caption         =   "Import"
         Height          =   255
         Left            =   2760
         TabIndex        =   473
         ToolTipText     =   "Import User from TXT"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdUserExport 
         Caption         =   "Export"
         Height          =   255
         Left            =   1980
         TabIndex        =   472
         ToolTipText     =   "Export User to TXT"
         Top             =   0
         Width           =   735
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   4935
         Left            =   0
         TabIndex        =   9
         Top             =   300
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   8705
         _Version        =   393216
         Tabs            =   8
         TabsPerRow      =   8
         TabHeight       =   520
         TabCaption(0)   =   "Stats"
         TabPicture(0)   =   "frmUser.frx":08CA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frameGeneral"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame7"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdReviveChar"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "cmdCalcExp"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmdPasteStatQ"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmdPasteChar"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Inven."
         TabPicture(1)   =   "frmUser.frx":08E6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frmKeys"
         Tab(1).Control(1)=   "frmItems"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Spellbk."
         TabPicture(2)   =   "frmUser.frx":0902
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frmSpells"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Abilities"
         TabPicture(3)   =   "frmUser.frx":091E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "cmdAbilsClear"
         Tab(3).Control(1)=   "SSTab2"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "Rooms"
         TabPicture(4)   =   "frmUser.frx":093A
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "txtCurrRoomDisp"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).Control(1)=   "cmdEditCurrentRoom"
         Tab(4).Control(2)=   "frmMapTrail"
         Tab(4).Control(3)=   "txtCurrentRoom"
         Tab(4).Control(4)=   "txtCurrentMap"
         Tab(4).Control(5)=   "Label74"
         Tab(4).ControlCount=   6
         TabCaption(5)   =   "Worn"
         TabPicture(5)   =   "frmUser.frx":0956
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "cmdPasteItems(2)"
         Tab(5).Control(1)=   "cmdEditWeapon"
         Tab(5).Control(2)=   "cmdClearWorn"
         Tab(5).Control(3)=   "Frame5"
         Tab(5).Control(4)=   "txtWeaponName"
         Tab(5).Control(4).Enabled=   0   'False
         Tab(5).Control(5)=   "txtWeaponNumber"
         Tab(5).Control(6)=   "Label75"
         Tab(5).ControlCount=   7
         TabCaption(6)   =   "Misc"
         TabPicture(6)   =   "frmUser.frx":0972
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "cmdUserHitPointMax"
         Tab(6).Control(1)=   "cmdUserHitPointRollQ"
         Tab(6).Control(2)=   "txtHitPointRolls"
         Tab(6).Control(3)=   "chkEdited"
         Tab(6).Control(4)=   "Frame4"
         Tab(6).Control(5)=   "txtTitle"
         Tab(6).Control(6)=   "txtCurrentEncum"
         Tab(6).Control(7)=   "txtMaxEncum"
         Tab(6).Control(8)=   "txtEvilPoints"
         Tab(6).Control(9)=   "txtBroadcastChan"
         Tab(6).Control(10)=   "txtCopper"
         Tab(6).Control(11)=   "txtSilver"
         Tab(6).Control(12)=   "txtGold"
         Tab(6).Control(13)=   "txtPlatinum"
         Tab(6).Control(14)=   "txtRunic"
         Tab(6).Control(15)=   "txtGang"
         Tab(6).Control(16)=   "txtSuicide"
         Tab(6).Control(17)=   "Label76(12)"
         Tab(6).Control(18)=   "Label76(11)"
         Tab(6).Control(19)=   "Label76(10)"
         Tab(6).Control(20)=   "Label76(9)"
         Tab(6).Control(21)=   "Label76(8)"
         Tab(6).Control(22)=   "Label76(7)"
         Tab(6).Control(23)=   "Label76(6)"
         Tab(6).Control(24)=   "Label76(5)"
         Tab(6).Control(25)=   "Label76(4)"
         Tab(6).Control(26)=   "Label76(3)"
         Tab(6).Control(27)=   "Label76(2)"
         Tab(6).Control(28)=   "Label76(1)"
         Tab(6).Control(29)=   "Label76(0)"
         Tab(6).ControlCount=   30
         TabCaption(7)   =   "???"
         TabPicture(7)   =   "frmUser.frx":098E
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "cmdUserUnknownsNote"
         Tab(7).Control(1)=   "optUserUnknowns(2)"
         Tab(7).Control(2)=   "optUserUnknowns(1)"
         Tab(7).Control(3)=   "optUserUnknowns(0)"
         Tab(7).Control(4)=   "txtUserUnknowns(48)"
         Tab(7).Control(5)=   "txtUserUnknowns(47)"
         Tab(7).Control(6)=   "txtUserUnknowns(46)"
         Tab(7).Control(7)=   "txtUserUnknowns(45)"
         Tab(7).Control(8)=   "txtUserUnknowns(44)"
         Tab(7).Control(9)=   "txtUserUnknowns(43)"
         Tab(7).Control(10)=   "txtUserUnknowns(42)"
         Tab(7).Control(11)=   "txtUserUnknowns(41)"
         Tab(7).Control(12)=   "txtUserUnknowns(40)"
         Tab(7).Control(13)=   "txtUserUnknowns(39)"
         Tab(7).Control(14)=   "txtUserUnknowns(38)"
         Tab(7).Control(15)=   "txtUserUnknowns(37)"
         Tab(7).Control(16)=   "txtUserUnknowns(36)"
         Tab(7).Control(17)=   "txtUserUnknowns(35)"
         Tab(7).Control(18)=   "txtUserUnknowns(34)"
         Tab(7).Control(19)=   "txtUserUnknowns(33)"
         Tab(7).Control(20)=   "txtUserUnknowns(32)"
         Tab(7).Control(21)=   "txtUserUnknowns(31)"
         Tab(7).Control(22)=   "txtUserUnknowns(30)"
         Tab(7).Control(23)=   "txtUserUnknowns(29)"
         Tab(7).Control(24)=   "txtUserUnknowns(28)"
         Tab(7).Control(25)=   "txtUserUnknowns(27)"
         Tab(7).Control(26)=   "txtUserUnknowns(26)"
         Tab(7).Control(27)=   "txtUserUnknowns(25)"
         Tab(7).Control(28)=   "txtUserUnknowns(24)"
         Tab(7).Control(29)=   "txtUserUnknowns(23)"
         Tab(7).Control(30)=   "txtUserUnknowns(22)"
         Tab(7).Control(31)=   "txtUserUnknowns(21)"
         Tab(7).Control(32)=   "txtUserUnknowns(20)"
         Tab(7).Control(33)=   "txtUserUnknowns(19)"
         Tab(7).Control(34)=   "txtUserUnknowns(18)"
         Tab(7).Control(35)=   "txtUserUnknowns(17)"
         Tab(7).Control(36)=   "txtUserUnknowns(16)"
         Tab(7).Control(37)=   "txtUserUnknowns(15)"
         Tab(7).Control(38)=   "txtUserUnknowns(14)"
         Tab(7).Control(39)=   "txtUserUnknowns(13)"
         Tab(7).Control(40)=   "txtUserUnknowns(12)"
         Tab(7).Control(41)=   "txtUserUnknowns(11)"
         Tab(7).Control(42)=   "txtUserUnknowns(10)"
         Tab(7).Control(43)=   "txtUserUnknowns(9)"
         Tab(7).Control(44)=   "txtUserUnknowns(8)"
         Tab(7).Control(45)=   "txtUserUnknowns(7)"
         Tab(7).Control(46)=   "txtUserUnknowns(6)"
         Tab(7).Control(47)=   "txtUserUnknowns(5)"
         Tab(7).Control(48)=   "txtUserUnknowns(4)"
         Tab(7).Control(49)=   "txtUserUnknowns(3)"
         Tab(7).Control(50)=   "txtUserUnknowns(2)"
         Tab(7).Control(51)=   "txtUserUnknowns(1)"
         Tab(7).Control(52)=   "txtUserUnknowns(0)"
         Tab(7).Control(53)=   "lblUserUnknowns(48)"
         Tab(7).Control(54)=   "lblUserUnknowns(47)"
         Tab(7).Control(55)=   "lblUserUnknowns(46)"
         Tab(7).Control(56)=   "lblUserUnknowns(45)"
         Tab(7).Control(57)=   "lblUserUnknowns(44)"
         Tab(7).Control(58)=   "lblUserUnknowns(43)"
         Tab(7).Control(59)=   "lblUserUnknowns(42)"
         Tab(7).Control(60)=   "lblUserUnknowns(41)"
         Tab(7).Control(61)=   "lblUserUnknowns(40)"
         Tab(7).Control(62)=   "lblUserUnknowns(39)"
         Tab(7).Control(63)=   "lblUserUnknowns(38)"
         Tab(7).Control(64)=   "lblUserUnknowns(37)"
         Tab(7).Control(65)=   "lblUserUnknowns(36)"
         Tab(7).Control(66)=   "lblUserUnknowns(35)"
         Tab(7).Control(67)=   "lblUserUnknowns(34)"
         Tab(7).Control(68)=   "lblUserUnknowns(33)"
         Tab(7).Control(69)=   "lblUserUnknowns(32)"
         Tab(7).Control(70)=   "lblUserUnknowns(31)"
         Tab(7).Control(71)=   "lblUserUnknowns(30)"
         Tab(7).Control(72)=   "lblUserUnknowns(29)"
         Tab(7).Control(73)=   "lblUserUnknowns(28)"
         Tab(7).Control(74)=   "lblUserUnknowns(27)"
         Tab(7).Control(75)=   "lblUserUnknowns(26)"
         Tab(7).Control(76)=   "lblUserUnknowns(25)"
         Tab(7).Control(77)=   "lblUserUnknowns(24)"
         Tab(7).Control(78)=   "lblUserUnknowns(23)"
         Tab(7).Control(79)=   "lblUserUnknowns(22)"
         Tab(7).Control(80)=   "lblUserUnknowns(21)"
         Tab(7).Control(81)=   "lblUserUnknowns(20)"
         Tab(7).Control(82)=   "lblUserUnknowns(19)"
         Tab(7).Control(83)=   "lblUserUnknowns(18)"
         Tab(7).Control(84)=   "lblUserUnknowns(17)"
         Tab(7).Control(85)=   "lblUserUnknowns(16)"
         Tab(7).Control(86)=   "lblUserUnknowns(15)"
         Tab(7).Control(87)=   "lblUserUnknowns(14)"
         Tab(7).Control(88)=   "lblUserUnknowns(13)"
         Tab(7).Control(89)=   "lblUserUnknowns(12)"
         Tab(7).Control(90)=   "lblUserUnknowns(11)"
         Tab(7).Control(91)=   "lblUserUnknowns(10)"
         Tab(7).Control(92)=   "lblUserUnknowns(9)"
         Tab(7).Control(93)=   "lblUserUnknowns(8)"
         Tab(7).Control(94)=   "lblUserUnknowns(7)"
         Tab(7).Control(95)=   "lblUserUnknowns(6)"
         Tab(7).Control(96)=   "lblUserUnknowns(5)"
         Tab(7).Control(97)=   "lblUserUnknowns(4)"
         Tab(7).Control(98)=   "lblUserUnknowns(3)"
         Tab(7).Control(99)=   "lblUserUnknowns(2)"
         Tab(7).Control(100)=   "lblUserUnknowns(1)"
         Tab(7).Control(101)=   "lblUserUnknowns(0)"
         Tab(7).ControlCount=   102
         Begin VB.CommandButton cmdPasteChar 
            Caption         =   "&Paste Stats"
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
            Left            =   5100
            TabIndex        =   13
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton cmdPasteStatQ 
            Caption         =   "?"
            Height          =   315
            Left            =   6240
            TabIndex        =   469
            Top             =   420
            Width           =   435
         End
         Begin VB.CommandButton cmdCalcExp 
            Caption         =   "Calc E&xp"
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
            Left            =   3660
            TabIndex        =   471
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton cmdReviveChar 
            Caption         =   "Revive Character"
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
            Left            =   3660
            Style           =   1  'Graphical
            TabIndex        =   595
            Top             =   420
            Width           =   2595
         End
         Begin VB.CommandButton cmdUserHitPointMax 
            Caption         =   "Max"
            Height          =   255
            Left            =   -69180
            TabIndex        =   591
            Top             =   1380
            Width           =   615
         End
         Begin VB.CommandButton cmdUserHitPointRollQ 
            Caption         =   "?"
            Height          =   315
            Left            =   -68520
            TabIndex        =   578
            Top             =   1680
            Width           =   195
         End
         Begin VB.TextBox txtHitPointRolls 
            Height          =   315
            Left            =   -69180
            MaxLength       =   5
            TabIndex        =   576
            Top             =   1680
            Width           =   615
         End
         Begin VB.CommandButton cmdUserUnknownsNote 
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
            Height          =   255
            Left            =   -71940
            TabIndex        =   575
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton optUserUnknowns 
            Caption         =   "Set 3"
            Height          =   315
            Index           =   2
            Left            =   -72840
            TabIndex        =   574
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton optUserUnknowns 
            Caption         =   "Set 2"
            Height          =   315
            Index           =   1
            Left            =   -73740
            TabIndex        =   573
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton optUserUnknowns 
            Caption         =   "Set 1"
            Height          =   315
            Index           =   0
            Left            =   -74640
            TabIndex        =   572
            Top             =   360
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   48
            Left            =   -69240
            TabIndex        =   564
            Top             =   4500
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   47
            Left            =   -70140
            TabIndex        =   563
            Top             =   4500
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   46
            Left            =   -71040
            TabIndex        =   562
            Top             =   4500
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   45
            Left            =   -71940
            TabIndex        =   561
            Top             =   4500
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   44
            Left            =   -72840
            TabIndex        =   560
            Top             =   4500
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   43
            Left            =   -73740
            TabIndex        =   559
            Top             =   4500
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   42
            Left            =   -74640
            TabIndex        =   558
            Top             =   4500
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   41
            Left            =   -69240
            TabIndex        =   550
            Top             =   3900
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   40
            Left            =   -70140
            TabIndex        =   549
            Top             =   3900
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   39
            Left            =   -71040
            TabIndex        =   548
            Top             =   3900
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   38
            Left            =   -71940
            TabIndex        =   547
            Top             =   3900
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   37
            Left            =   -72840
            TabIndex        =   546
            Top             =   3900
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   36
            Left            =   -73740
            TabIndex        =   545
            Top             =   3900
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   35
            Left            =   -74640
            TabIndex        =   544
            Top             =   3900
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   34
            Left            =   -69240
            TabIndex        =   536
            Top             =   3300
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   33
            Left            =   -70140
            TabIndex        =   535
            Top             =   3300
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   32
            Left            =   -71040
            TabIndex        =   534
            Top             =   3300
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   31
            Left            =   -71940
            TabIndex        =   533
            Top             =   3300
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   30
            Left            =   -72840
            TabIndex        =   532
            Top             =   3300
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   29
            Left            =   -73740
            TabIndex        =   531
            Top             =   3300
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   28
            Left            =   -74640
            TabIndex        =   530
            Top             =   3300
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   27
            Left            =   -69240
            TabIndex        =   522
            Top             =   2700
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   26
            Left            =   -70140
            TabIndex        =   521
            Top             =   2700
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   25
            Left            =   -71040
            TabIndex        =   520
            Top             =   2700
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   24
            Left            =   -71940
            TabIndex        =   519
            Top             =   2700
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   23
            Left            =   -72840
            TabIndex        =   518
            Top             =   2700
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   22
            Left            =   -73740
            TabIndex        =   517
            Top             =   2700
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   21
            Left            =   -74640
            TabIndex        =   516
            Top             =   2700
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   20
            Left            =   -69240
            TabIndex        =   508
            Top             =   2100
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   19
            Left            =   -70140
            TabIndex        =   507
            Top             =   2100
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   18
            Left            =   -71040
            TabIndex        =   506
            Top             =   2100
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   17
            Left            =   -71940
            TabIndex        =   505
            Top             =   2100
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   16
            Left            =   -72840
            TabIndex        =   504
            Top             =   2100
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   15
            Left            =   -73740
            TabIndex        =   503
            Top             =   2100
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   14
            Left            =   -74640
            TabIndex        =   502
            Top             =   2100
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   13
            Left            =   -69240
            TabIndex        =   487
            Top             =   1500
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   12
            Left            =   -70140
            TabIndex        =   486
            Top             =   1500
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   11
            Left            =   -71040
            TabIndex        =   485
            Top             =   1500
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   10
            Left            =   -71940
            TabIndex        =   484
            Top             =   1500
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   9
            Left            =   -72840
            TabIndex        =   483
            Top             =   1500
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   8
            Left            =   -73740
            TabIndex        =   482
            Top             =   1500
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   7
            Left            =   -74640
            TabIndex        =   481
            Top             =   1500
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   6
            Left            =   -69240
            TabIndex        =   480
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   5
            Left            =   -70140
            TabIndex        =   479
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   4
            Left            =   -71040
            TabIndex        =   478
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   3
            Left            =   -71940
            TabIndex        =   477
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   2
            Left            =   -72840
            TabIndex        =   476
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   1
            Left            =   -73740
            TabIndex        =   475
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox txtUserUnknowns 
            Height          =   285
            Index           =   0
            Left            =   -74640
            TabIndex        =   474
            Top             =   900
            Width           =   735
         End
         Begin VB.CheckBox chkEdited 
            Caption         =   "EDITED Flag"
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
            Left            =   -69960
            TabIndex        =   470
            Top             =   4380
            Width           =   1515
         End
         Begin VB.TextBox txtCurrRoomDisp 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   -71280
            Locked          =   -1  'True
            TabIndex        =   467
            TabStop         =   0   'False
            Top             =   540
            Width           =   2955
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
            Left            =   -74760
            TabIndex        =   108
            Top             =   1080
            Width           =   675
         End
         Begin VB.CommandButton cmdPasteItems 
            Caption         =   "&Paste Items/Keys/Worn/$$"
            Height          =   315
            Index           =   2
            Left            =   -74580
            TabIndex        =   300
            Top             =   480
            Width           =   2415
         End
         Begin VB.Frame Frame4 
            Caption         =   "Spells Casted on User"
            Height          =   3315
            Left            =   -74880
            TabIndex        =   391
            Top             =   1500
            Width           =   4455
            Begin VB.CommandButton cmdSomeFlagQ 
               Caption         =   "?"
               Height          =   315
               Left            =   2040
               TabIndex        =   594
               Top             =   180
               Width           =   195
            End
            Begin VB.TextBox txtSomeSortOfFlag 
               Height          =   285
               Left            =   240
               TabIndex        =   592
               Top             =   240
               Width           =   615
            End
            Begin VB.CommandButton cmdEditSpellCasted 
               Height          =   135
               Index           =   9
               Left            =   60
               TabIndex        =   442
               Top             =   3000
               Width           =   135
            End
            Begin VB.CommandButton cmdEditSpellCasted 
               Height          =   135
               Index           =   8
               Left            =   60
               TabIndex        =   437
               Top             =   2760
               Width           =   135
            End
            Begin VB.CommandButton cmdEditSpellCasted 
               Height          =   135
               Index           =   7
               Left            =   60
               TabIndex        =   432
               Top             =   2520
               Width           =   135
            End
            Begin VB.CommandButton cmdEditSpellCasted 
               Height          =   135
               Index           =   6
               Left            =   60
               TabIndex        =   427
               Top             =   2280
               Width           =   135
            End
            Begin VB.CommandButton cmdEditSpellCasted 
               Height          =   135
               Index           =   5
               Left            =   60
               TabIndex        =   422
               Top             =   2040
               Width           =   135
            End
            Begin VB.CommandButton cmdEditSpellCasted 
               Height          =   135
               Index           =   4
               Left            =   60
               TabIndex        =   417
               Top             =   1800
               Width           =   135
            End
            Begin VB.CommandButton cmdEditSpellCasted 
               Height          =   135
               Index           =   3
               Left            =   60
               TabIndex        =   412
               Top             =   1560
               Width           =   135
            End
            Begin VB.CommandButton cmdEditSpellCasted 
               Height          =   135
               Index           =   2
               Left            =   60
               TabIndex        =   407
               Top             =   1320
               Width           =   135
            End
            Begin VB.CommandButton cmdEditSpellCasted 
               Height          =   135
               Index           =   1
               Left            =   60
               TabIndex        =   402
               Top             =   1080
               Width           =   135
            End
            Begin VB.CommandButton cmdEditSpellCasted 
               Height          =   135
               Index           =   0
               Left            =   60
               TabIndex        =   397
               Top             =   840
               Width           =   135
            End
            Begin VB.CommandButton cmdClearSpellsCasted 
               Caption         =   "Clear All"
               Height          =   255
               Left            =   3000
               TabIndex        =   392
               Top             =   210
               Width           =   1335
            End
            Begin VB.TextBox txtSpellNumber 
               Height          =   285
               Index           =   0
               Left            =   240
               MaxLength       =   5
               TabIndex        =   398
               Top             =   780
               Width           =   615
            End
            Begin VB.TextBox txtSpellName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   0
               Left            =   840
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   399
               TabStop         =   0   'False
               Top             =   780
               Width           =   2295
            End
            Begin VB.TextBox txtSpellNumber 
               Height          =   285
               Index           =   1
               Left            =   240
               MaxLength       =   5
               TabIndex        =   403
               Top             =   1020
               Width           =   615
            End
            Begin VB.TextBox txtSpellName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   1
               Left            =   840
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   404
               TabStop         =   0   'False
               Top             =   1020
               Width           =   2295
            End
            Begin VB.TextBox txtSpellNumber 
               Height          =   285
               Index           =   2
               Left            =   240
               MaxLength       =   5
               TabIndex        =   408
               Top             =   1260
               Width           =   615
            End
            Begin VB.TextBox txtSpellName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   2
               Left            =   840
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   409
               TabStop         =   0   'False
               Top             =   1260
               Width           =   2295
            End
            Begin VB.TextBox txtSpellNumber 
               Height          =   285
               Index           =   3
               Left            =   240
               MaxLength       =   5
               TabIndex        =   413
               Top             =   1500
               Width           =   615
            End
            Begin VB.TextBox txtSpellName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   3
               Left            =   840
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   414
               TabStop         =   0   'False
               Top             =   1500
               Width           =   2295
            End
            Begin VB.TextBox txtSpellNumber 
               Height          =   285
               Index           =   4
               Left            =   240
               MaxLength       =   5
               TabIndex        =   418
               Top             =   1740
               Width           =   615
            End
            Begin VB.TextBox txtSpellName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   4
               Left            =   840
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   419
               TabStop         =   0   'False
               Top             =   1740
               Width           =   2295
            End
            Begin VB.TextBox txtSpellNumber 
               Height          =   285
               Index           =   5
               Left            =   240
               MaxLength       =   5
               TabIndex        =   423
               Top             =   1980
               Width           =   615
            End
            Begin VB.TextBox txtSpellName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   5
               Left            =   840
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   424
               TabStop         =   0   'False
               Top             =   1980
               Width           =   2295
            End
            Begin VB.TextBox txtSpellNumber 
               Height          =   285
               Index           =   6
               Left            =   240
               MaxLength       =   5
               TabIndex        =   428
               Top             =   2220
               Width           =   615
            End
            Begin VB.TextBox txtSpellName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   6
               Left            =   840
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   429
               TabStop         =   0   'False
               Top             =   2220
               Width           =   2295
            End
            Begin VB.TextBox txtSpellNumber 
               Height          =   285
               Index           =   7
               Left            =   240
               MaxLength       =   5
               TabIndex        =   433
               Top             =   2460
               Width           =   615
            End
            Begin VB.TextBox txtSpellName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   7
               Left            =   840
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   434
               TabStop         =   0   'False
               Top             =   2460
               Width           =   2295
            End
            Begin VB.TextBox txtSpellNumber 
               Height          =   285
               Index           =   8
               Left            =   240
               MaxLength       =   5
               TabIndex        =   438
               Top             =   2700
               Width           =   615
            End
            Begin VB.TextBox txtSpellName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   8
               Left            =   840
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   439
               TabStop         =   0   'False
               Top             =   2700
               Width           =   2295
            End
            Begin VB.TextBox txtSpellNumber 
               Height          =   285
               Index           =   9
               Left            =   240
               MaxLength       =   5
               TabIndex        =   443
               Top             =   2940
               Width           =   615
            End
            Begin VB.TextBox txtSpellName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   9
               Left            =   840
               Locked          =   -1  'True
               MaxLength       =   28
               TabIndex        =   444
               TabStop         =   0   'False
               Top             =   2940
               Width           =   2295
            End
            Begin VB.TextBox txtSpellValue 
               Height          =   285
               Index           =   0
               Left            =   3105
               TabIndex        =   400
               Top             =   780
               Width           =   615
            End
            Begin VB.TextBox txtSpellRounds 
               Height          =   285
               Index           =   0
               Left            =   3690
               TabIndex        =   401
               Top             =   780
               Width           =   615
            End
            Begin VB.TextBox txtSpellValue 
               Height          =   285
               Index           =   1
               Left            =   3105
               TabIndex        =   405
               Top             =   1020
               Width           =   615
            End
            Begin VB.TextBox txtSpellRounds 
               Height          =   285
               Index           =   1
               Left            =   3690
               TabIndex        =   406
               Top             =   1020
               Width           =   615
            End
            Begin VB.TextBox txtSpellValue 
               Height          =   285
               Index           =   2
               Left            =   3105
               TabIndex        =   410
               Top             =   1260
               Width           =   615
            End
            Begin VB.TextBox txtSpellRounds 
               Height          =   285
               Index           =   2
               Left            =   3690
               TabIndex        =   411
               Top             =   1260
               Width           =   615
            End
            Begin VB.TextBox txtSpellValue 
               Height          =   285
               Index           =   3
               Left            =   3105
               TabIndex        =   415
               Top             =   1500
               Width           =   615
            End
            Begin VB.TextBox txtSpellRounds 
               Height          =   285
               Index           =   3
               Left            =   3690
               TabIndex        =   416
               Top             =   1500
               Width           =   615
            End
            Begin VB.TextBox txtSpellValue 
               Height          =   285
               Index           =   4
               Left            =   3105
               TabIndex        =   420
               Top             =   1740
               Width           =   615
            End
            Begin VB.TextBox txtSpellRounds 
               Height          =   285
               Index           =   4
               Left            =   3690
               TabIndex        =   421
               Top             =   1740
               Width           =   615
            End
            Begin VB.TextBox txtSpellValue 
               Height          =   285
               Index           =   5
               Left            =   3105
               TabIndex        =   425
               Top             =   1980
               Width           =   615
            End
            Begin VB.TextBox txtSpellRounds 
               Height          =   285
               Index           =   5
               Left            =   3690
               TabIndex        =   426
               Top             =   1980
               Width           =   615
            End
            Begin VB.TextBox txtSpellValue 
               Height          =   285
               Index           =   6
               Left            =   3105
               TabIndex        =   430
               Top             =   2220
               Width           =   615
            End
            Begin VB.TextBox txtSpellRounds 
               Height          =   285
               Index           =   6
               Left            =   3690
               TabIndex        =   431
               Top             =   2220
               Width           =   615
            End
            Begin VB.TextBox txtSpellValue 
               Height          =   285
               Index           =   7
               Left            =   3105
               TabIndex        =   435
               Top             =   2460
               Width           =   615
            End
            Begin VB.TextBox txtSpellRounds 
               Height          =   285
               Index           =   7
               Left            =   3690
               TabIndex        =   436
               Top             =   2460
               Width           =   615
            End
            Begin VB.TextBox txtSpellValue 
               Height          =   285
               Index           =   8
               Left            =   3105
               TabIndex        =   440
               Top             =   2700
               Width           =   615
            End
            Begin VB.TextBox txtSpellRounds 
               Height          =   285
               Index           =   8
               Left            =   3690
               TabIndex        =   441
               Top             =   2700
               Width           =   615
            End
            Begin VB.TextBox txtSpellValue 
               Height          =   285
               Index           =   9
               Left            =   3105
               TabIndex        =   445
               Top             =   2940
               Width           =   615
            End
            Begin VB.TextBox txtSpellRounds 
               Height          =   285
               Index           =   9
               Left            =   3690
               TabIndex        =   446
               Top             =   2940
               Width           =   615
            End
            Begin VB.Label Label76 
               AutoSize        =   -1  'True
               Caption         =   " <-- Poison flag?"
               Height          =   195
               Index           =   13
               Left            =   840
               TabIndex        =   593
               Top             =   255
               Width           =   1920
            End
            Begin VB.Label Label32 
               Caption         =   "Spell#"
               Height          =   255
               Left            =   240
               TabIndex        =   393
               Top             =   540
               Width           =   495
            End
            Begin VB.Label Label33 
               Caption         =   "Spell Name"
               Height          =   255
               Left            =   840
               TabIndex        =   394
               Top             =   540
               Width           =   1215
            End
            Begin VB.Label Label34 
               Caption         =   "Value"
               Height          =   255
               Left            =   3120
               TabIndex        =   395
               Top             =   540
               Width           =   495
            End
            Begin VB.Label Label35 
               Caption         =   "Rounds"
               Height          =   255
               Left            =   3720
               TabIndex        =   396
               Top             =   540
               Width           =   615
            End
         End
         Begin VB.TextBox txtTitle 
            Height          =   285
            Left            =   -74100
            MaxLength       =   19
            TabIndex        =   368
            Top             =   435
            Width           =   2415
         End
         Begin VB.TextBox txtCurrentEncum 
            Height          =   285
            Left            =   -73440
            TabIndex        =   372
            Top             =   1155
            Width           =   735
         End
         Begin VB.TextBox txtMaxEncum 
            Height          =   285
            Left            =   -72420
            TabIndex        =   374
            Top             =   1155
            Width           =   735
         End
         Begin VB.TextBox txtEvilPoints 
            Height          =   285
            Left            =   -70260
            TabIndex        =   378
            Top             =   795
            Width           =   615
         End
         Begin VB.TextBox txtBroadcastChan 
            Height          =   285
            Left            =   -70260
            TabIndex        =   380
            Top             =   1155
            Width           =   615
         End
         Begin VB.TextBox txtCopper 
            Height          =   315
            Left            =   -69180
            MaxLength       =   5
            TabIndex        =   390
            Top             =   3735
            Width           =   615
         End
         Begin VB.TextBox txtSilver 
            Height          =   315
            Left            =   -69180
            MaxLength       =   5
            TabIndex        =   388
            Top             =   3375
            Width           =   615
         End
         Begin VB.TextBox txtGold 
            Height          =   315
            Left            =   -69180
            MaxLength       =   5
            TabIndex        =   386
            Top             =   3015
            Width           =   615
         End
         Begin VB.TextBox txtPlatinum 
            Height          =   315
            Left            =   -69180
            MaxLength       =   5
            TabIndex        =   384
            Top             =   2655
            Width           =   615
         End
         Begin VB.TextBox txtRunic 
            Height          =   315
            Left            =   -69180
            MaxLength       =   5
            TabIndex        =   382
            Top             =   2295
            Width           =   615
         End
         Begin VB.TextBox txtGang 
            Height          =   285
            Left            =   -74100
            MaxLength       =   19
            TabIndex        =   370
            Top             =   795
            Width           =   2415
         End
         Begin VB.TextBox txtSuicide 
            Height          =   285
            Left            =   -70260
            TabIndex        =   376
            Top             =   420
            Width           =   1335
         End
         Begin VB.CommandButton cmdEditWeapon 
            Height          =   195
            Left            =   -74580
            TabIndex        =   301
            Top             =   1200
            Width           =   195
         End
         Begin VB.CommandButton cmdClearWorn 
            Caption         =   "Clear All"
            Height          =   315
            Left            =   -70260
            TabIndex        =   305
            Top             =   1140
            Width           =   1515
         End
         Begin VB.Frame Frame5 
            Caption         =   "Worn on Body"
            Height          =   2895
            Left            =   -74640
            TabIndex        =   306
            Top             =   1680
            Width           =   6135
            Begin VB.CommandButton cmdEditWornItem 
               Height          =   135
               Index           =   19
               Left            =   3120
               TabIndex        =   364
               Top             =   2520
               Width           =   135
            End
            Begin VB.CommandButton cmdEditWornItem 
               Height          =   135
               Index           =   18
               Left            =   3120
               TabIndex        =   361
               Top             =   2280
               Width           =   135
            End
            Begin VB.CommandButton cmdEditWornItem 
               Height          =   135
               Index           =   17
               Left            =   3120
               TabIndex        =   358
               Top             =   2040
               Width           =   135
            End
            Begin VB.CommandButton cmdEditWornItem 
               Height          =   135
               Index           =   16
               Left            =   3120
               TabIndex        =   355
               Top             =   1800
               Width           =   135
            End
            Begin VB.CommandButton cmdEditWornItem 
               Height          =   135
               Index           =   15
               Left            =   3120
               TabIndex        =   352
               Top             =   1560
               Width           =   135
            End
            Begin VB.CommandButton cmdEditWornItem 
               Height          =   135
               Index           =   14
               Left            =   3120
               TabIndex        =   349
               Top             =   1320
               Width           =   135
            End
            Begin VB.CommandButton cmdEditWornItem 
               Height          =   135
               Index           =   13
               Left            =   3120
               TabIndex        =   346
               Top             =   1080
               Width           =   135
            End
            Begin VB.CommandButton cmdEditWornItem 
               Height          =   135
               Index           =   12
               Left            =   3120
               TabIndex        =   343
               Top             =   840
               Width           =   135
            End
            Begin VB.CommandButton cmdEditWornItem 
               Height          =   135
               Index           =   11
               Left            =   3120
               TabIndex        =   340
               Top             =   600
               Width           =   135
            End
            Begin VB.CommandButton cmdEditWornItem 
               Height          =   135
               Index           =   10
               Left            =   3120
               TabIndex        =   337
               Top             =   360
               Width           =   135
            End
            Begin VB.CommandButton cmdEditWornItem 
               Height          =   135
               Index           =   9
               Left            =   120
               TabIndex        =   334
               Top             =   2520
               Width           =   135
            End
            Begin VB.CommandButton cmdEditWornItem 
               Height          =   135
               Index           =   8
               Left            =   120
               TabIndex        =   331
               Top             =   2280
               Width           =   135
            End
            Begin VB.CommandButton cmdEditWornItem 
               Height          =   135
               Index           =   7
               Left            =   120
               TabIndex        =   328
               Top             =   2040
               Width           =   135
            End
            Begin VB.CommandButton cmdEditWornItem 
               Height          =   135
               Index           =   6
               Left            =   120
               TabIndex        =   325
               Top             =   1800
               Width           =   135
            End
            Begin VB.CommandButton cmdEditWornItem 
               Height          =   135
               Index           =   5
               Left            =   120
               TabIndex        =   322
               Top             =   1560
               Width           =   135
            End
            Begin VB.CommandButton cmdEditWornItem 
               Height          =   135
               Index           =   4
               Left            =   120
               TabIndex        =   319
               Top             =   1320
               Width           =   135
            End
            Begin VB.CommandButton cmdEditWornItem 
               Height          =   135
               Index           =   3
               Left            =   120
               TabIndex        =   316
               Top             =   1080
               Width           =   135
            End
            Begin VB.CommandButton cmdEditWornItem 
               Height          =   135
               Index           =   2
               Left            =   120
               TabIndex        =   313
               Top             =   840
               Width           =   135
            End
            Begin VB.CommandButton cmdEditWornItem 
               Height          =   135
               Index           =   1
               Left            =   120
               TabIndex        =   310
               Top             =   600
               Width           =   135
            End
            Begin VB.CommandButton cmdEditWornItem 
               Height          =   135
               Index           =   0
               Left            =   120
               TabIndex        =   307
               Top             =   360
               Width           =   135
            End
            Begin VB.TextBox txtWornItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   19
               Left            =   4020
               Locked          =   -1  'True
               TabIndex        =   365
               TabStop         =   0   'False
               Top             =   2460
               Width           =   1935
            End
            Begin VB.TextBox txtWornItem 
               Height          =   285
               Index           =   19
               Left            =   3300
               TabIndex        =   366
               Top             =   2460
               Width           =   735
            End
            Begin VB.TextBox txtWornItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   18
               Left            =   4020
               Locked          =   -1  'True
               TabIndex        =   363
               TabStop         =   0   'False
               Top             =   2220
               Width           =   1935
            End
            Begin VB.TextBox txtWornItem 
               Height          =   285
               Index           =   18
               Left            =   3300
               TabIndex        =   362
               Top             =   2220
               Width           =   735
            End
            Begin VB.TextBox txtWornItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   17
               Left            =   4020
               Locked          =   -1  'True
               TabIndex        =   360
               TabStop         =   0   'False
               Top             =   1980
               Width           =   1935
            End
            Begin VB.TextBox txtWornItem 
               Height          =   285
               Index           =   17
               Left            =   3300
               TabIndex        =   359
               Top             =   1980
               Width           =   735
            End
            Begin VB.TextBox txtWornItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   16
               Left            =   4020
               Locked          =   -1  'True
               TabIndex        =   357
               TabStop         =   0   'False
               Top             =   1740
               Width           =   1935
            End
            Begin VB.TextBox txtWornItem 
               Height          =   285
               Index           =   16
               Left            =   3300
               TabIndex        =   356
               Top             =   1740
               Width           =   735
            End
            Begin VB.TextBox txtWornItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   15
               Left            =   4020
               Locked          =   -1  'True
               TabIndex        =   354
               TabStop         =   0   'False
               Top             =   1500
               Width           =   1935
            End
            Begin VB.TextBox txtWornItem 
               Height          =   285
               Index           =   15
               Left            =   3300
               TabIndex        =   353
               Top             =   1500
               Width           =   735
            End
            Begin VB.TextBox txtWornItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   14
               Left            =   4020
               Locked          =   -1  'True
               TabIndex        =   351
               TabStop         =   0   'False
               Top             =   1260
               Width           =   1935
            End
            Begin VB.TextBox txtWornItem 
               Height          =   285
               Index           =   14
               Left            =   3300
               TabIndex        =   350
               Top             =   1260
               Width           =   735
            End
            Begin VB.TextBox txtWornItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   13
               Left            =   4020
               Locked          =   -1  'True
               TabIndex        =   348
               TabStop         =   0   'False
               Top             =   1020
               Width           =   1935
            End
            Begin VB.TextBox txtWornItem 
               Height          =   285
               Index           =   13
               Left            =   3300
               TabIndex        =   347
               Top             =   1020
               Width           =   735
            End
            Begin VB.TextBox txtWornItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   12
               Left            =   4020
               Locked          =   -1  'True
               TabIndex        =   345
               TabStop         =   0   'False
               Top             =   780
               Width           =   1935
            End
            Begin VB.TextBox txtWornItem 
               Height          =   285
               Index           =   12
               Left            =   3300
               TabIndex        =   344
               Top             =   780
               Width           =   735
            End
            Begin VB.TextBox txtWornItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   11
               Left            =   4020
               Locked          =   -1  'True
               TabIndex        =   342
               TabStop         =   0   'False
               Top             =   540
               Width           =   1935
            End
            Begin VB.TextBox txtWornItem 
               Height          =   285
               Index           =   11
               Left            =   3300
               TabIndex        =   341
               Top             =   540
               Width           =   735
            End
            Begin VB.TextBox txtWornItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   10
               Left            =   4020
               Locked          =   -1  'True
               TabIndex        =   339
               TabStop         =   0   'False
               Top             =   300
               Width           =   1935
            End
            Begin VB.TextBox txtWornItem 
               Height          =   285
               Index           =   10
               Left            =   3300
               TabIndex        =   338
               Top             =   300
               Width           =   735
            End
            Begin VB.TextBox txtWornItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   9
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   336
               TabStop         =   0   'False
               Top             =   2460
               Width           =   1935
            End
            Begin VB.TextBox txtWornItem 
               Height          =   285
               Index           =   9
               Left            =   300
               TabIndex        =   335
               Top             =   2460
               Width           =   735
            End
            Begin VB.TextBox txtWornItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   8
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   333
               TabStop         =   0   'False
               Top             =   2220
               Width           =   1935
            End
            Begin VB.TextBox txtWornItem 
               Height          =   285
               Index           =   8
               Left            =   300
               TabIndex        =   332
               Top             =   2220
               Width           =   735
            End
            Begin VB.TextBox txtWornItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   7
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   330
               TabStop         =   0   'False
               Top             =   1980
               Width           =   1935
            End
            Begin VB.TextBox txtWornItem 
               Height          =   285
               Index           =   7
               Left            =   300
               TabIndex        =   329
               Top             =   1980
               Width           =   735
            End
            Begin VB.TextBox txtWornItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   6
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   327
               TabStop         =   0   'False
               Top             =   1740
               Width           =   1935
            End
            Begin VB.TextBox txtWornItem 
               Height          =   285
               Index           =   6
               Left            =   300
               TabIndex        =   326
               Top             =   1740
               Width           =   735
            End
            Begin VB.TextBox txtWornItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   5
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   324
               TabStop         =   0   'False
               Top             =   1500
               Width           =   1935
            End
            Begin VB.TextBox txtWornItem 
               Height          =   285
               Index           =   5
               Left            =   300
               TabIndex        =   323
               Top             =   1500
               Width           =   735
            End
            Begin VB.TextBox txtWornItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   4
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   321
               TabStop         =   0   'False
               Top             =   1260
               Width           =   1935
            End
            Begin VB.TextBox txtWornItem 
               Height          =   285
               Index           =   4
               Left            =   300
               TabIndex        =   320
               Top             =   1260
               Width           =   735
            End
            Begin VB.TextBox txtWornItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   3
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   318
               TabStop         =   0   'False
               Top             =   1020
               Width           =   1935
            End
            Begin VB.TextBox txtWornItem 
               Height          =   285
               Index           =   3
               Left            =   300
               TabIndex        =   317
               Top             =   1020
               Width           =   735
            End
            Begin VB.TextBox txtWornItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   2
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   315
               TabStop         =   0   'False
               Top             =   780
               Width           =   1935
            End
            Begin VB.TextBox txtWornItem 
               Height          =   285
               Index           =   2
               Left            =   300
               TabIndex        =   314
               Top             =   780
               Width           =   735
            End
            Begin VB.TextBox txtWornItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   1
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   312
               TabStop         =   0   'False
               Top             =   540
               Width           =   1935
            End
            Begin VB.TextBox txtWornItem 
               Height          =   285
               Index           =   1
               Left            =   300
               TabIndex        =   311
               Top             =   540
               Width           =   735
            End
            Begin VB.TextBox txtWornItemName 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   0
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   309
               TabStop         =   0   'False
               Top             =   300
               Width           =   1935
            End
            Begin VB.TextBox txtWornItem 
               Height          =   285
               Index           =   0
               Left            =   300
               TabIndex        =   308
               Top             =   300
               Width           =   735
            End
         End
         Begin VB.TextBox txtWeaponName 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   -72660
            Locked          =   -1  'True
            TabIndex        =   304
            TabStop         =   0   'False
            Top             =   1140
            Width           =   1935
         End
         Begin VB.TextBox txtWeaponNumber 
            Height          =   285
            Left            =   -73380
            TabIndex        =   303
            Top             =   1140
            Width           =   735
         End
         Begin VB.CommandButton cmdEditCurrentRoom 
            Height          =   195
            Left            =   -74640
            TabIndex        =   211
            Top             =   600
            Width           =   195
         End
         Begin VB.Frame frmMapTrail 
            Caption         =   "Trail of Last Rooms (# = # of rooms back)"
            Height          =   3495
            Left            =   -74880
            TabIndex        =   215
            Top             =   1140
            Width           =   6615
            Begin VB.CheckBox chkNoRoomNames 
               Caption         =   "Don't Lookup Room Names"
               Height          =   195
               Left            =   120
               TabIndex        =   468
               Top             =   300
               Width           =   2895
            End
            Begin VB.TextBox txtRoomTrailDisp 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   19
               Left            =   5460
               Locked          =   -1  'True
               TabIndex        =   466
               TabStop         =   0   'False
               Top             =   3060
               Width           =   1035
            End
            Begin VB.TextBox txtRoomTrailDisp 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   18
               Left            =   5460
               Locked          =   -1  'True
               TabIndex        =   465
               TabStop         =   0   'False
               Top             =   2820
               Width           =   1035
            End
            Begin VB.TextBox txtRoomTrailDisp 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   17
               Left            =   5460
               Locked          =   -1  'True
               TabIndex        =   464
               TabStop         =   0   'False
               Top             =   2580
               Width           =   1035
            End
            Begin VB.TextBox txtRoomTrailDisp 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   16
               Left            =   5460
               Locked          =   -1  'True
               TabIndex        =   463
               TabStop         =   0   'False
               Top             =   2340
               Width           =   1035
            End
            Begin VB.TextBox txtRoomTrailDisp 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   15
               Left            =   5460
               Locked          =   -1  'True
               TabIndex        =   462
               TabStop         =   0   'False
               Top             =   2100
               Width           =   1035
            End
            Begin VB.TextBox txtRoomTrailDisp 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   14
               Left            =   5460
               Locked          =   -1  'True
               TabIndex        =   461
               TabStop         =   0   'False
               Top             =   1860
               Width           =   1035
            End
            Begin VB.TextBox txtRoomTrailDisp 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   13
               Left            =   5460
               Locked          =   -1  'True
               TabIndex        =   460
               TabStop         =   0   'False
               Top             =   1620
               Width           =   1035
            End
            Begin VB.TextBox txtRoomTrailDisp 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   12
               Left            =   5460
               Locked          =   -1  'True
               TabIndex        =   459
               TabStop         =   0   'False
               Top             =   1380
               Width           =   1035
            End
            Begin VB.TextBox txtRoomTrailDisp 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   11
               Left            =   5460
               Locked          =   -1  'True
               TabIndex        =   458
               TabStop         =   0   'False
               Top             =   1140
               Width           =   1035
            End
            Begin VB.TextBox txtRoomTrailDisp 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   10
               Left            =   5460
               Locked          =   -1  'True
               TabIndex        =   457
               TabStop         =   0   'False
               Top             =   900
               Width           =   1035
            End
            Begin VB.TextBox txtRoomTrailDisp 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   9
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   456
               TabStop         =   0   'False
               Top             =   3060
               Width           =   1875
            End
            Begin VB.TextBox txtRoomTrailDisp 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   8
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   455
               TabStop         =   0   'False
               Top             =   2820
               Width           =   1875
            End
            Begin VB.TextBox txtRoomTrailDisp 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   7
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   454
               TabStop         =   0   'False
               Top             =   2580
               Width           =   1875
            End
            Begin VB.TextBox txtRoomTrailDisp 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   6
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   453
               TabStop         =   0   'False
               Top             =   2340
               Width           =   1875
            End
            Begin VB.TextBox txtRoomTrailDisp 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   5
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   452
               TabStop         =   0   'False
               Top             =   2100
               Width           =   1875
            End
            Begin VB.TextBox txtRoomTrailDisp 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   4
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   451
               TabStop         =   0   'False
               Top             =   1860
               Width           =   1875
            End
            Begin VB.TextBox txtRoomTrailDisp 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   3
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   450
               TabStop         =   0   'False
               Top             =   1620
               Width           =   1875
            End
            Begin VB.TextBox txtRoomTrailDisp 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   2
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   449
               TabStop         =   0   'False
               Top             =   1380
               Width           =   1875
            End
            Begin VB.TextBox txtRoomTrailDisp 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   1
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   448
               TabStop         =   0   'False
               Top             =   1140
               Width           =   1875
            End
            Begin VB.TextBox txtRoomTrailDisp 
               BackColor       =   &H8000000F&
               Height          =   285
               Index           =   0
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   447
               TabStop         =   0   'False
               Top             =   900
               Width           =   1875
            End
            Begin VB.CommandButton cmdGotoRoom 
               Height          =   135
               Index           =   19
               Left            =   3780
               TabIndex        =   277
               Top             =   3120
               Width           =   135
            End
            Begin VB.CommandButton cmdGotoRoom 
               Height          =   135
               Index           =   18
               Left            =   3780
               TabIndex        =   274
               Top             =   2880
               Width           =   135
            End
            Begin VB.CommandButton cmdGotoRoom 
               Height          =   135
               Index           =   17
               Left            =   3780
               TabIndex        =   271
               Top             =   2640
               Width           =   135
            End
            Begin VB.CommandButton cmdGotoRoom 
               Height          =   135
               Index           =   16
               Left            =   3780
               TabIndex        =   268
               Top             =   2400
               Width           =   135
            End
            Begin VB.CommandButton cmdGotoRoom 
               Height          =   135
               Index           =   15
               Left            =   3780
               TabIndex        =   265
               Top             =   2160
               Width           =   135
            End
            Begin VB.CommandButton cmdGotoRoom 
               Height          =   135
               Index           =   14
               Left            =   3780
               TabIndex        =   262
               Top             =   1920
               Width           =   135
            End
            Begin VB.CommandButton cmdGotoRoom 
               Height          =   135
               Index           =   13
               Left            =   3780
               TabIndex        =   259
               Top             =   1680
               Width           =   135
            End
            Begin VB.CommandButton cmdGotoRoom 
               Height          =   135
               Index           =   12
               Left            =   3780
               TabIndex        =   256
               Top             =   1440
               Width           =   135
            End
            Begin VB.CommandButton cmdGotoRoom 
               Height          =   135
               Index           =   11
               Left            =   3780
               TabIndex        =   253
               Top             =   1200
               Width           =   135
            End
            Begin VB.CommandButton cmdGotoRoom 
               Height          =   135
               Index           =   10
               Left            =   3780
               TabIndex        =   250
               Top             =   960
               Width           =   135
            End
            Begin VB.CommandButton cmdGotoRoom 
               Height          =   135
               Index           =   9
               Left            =   120
               TabIndex        =   247
               Top             =   3120
               Width           =   135
            End
            Begin VB.CommandButton cmdGotoRoom 
               Height          =   135
               Index           =   8
               Left            =   120
               TabIndex        =   244
               Top             =   2880
               Width           =   135
            End
            Begin VB.CommandButton cmdGotoRoom 
               Height          =   135
               Index           =   7
               Left            =   120
               TabIndex        =   241
               Top             =   2640
               Width           =   135
            End
            Begin VB.CommandButton cmdGotoRoom 
               Height          =   135
               Index           =   6
               Left            =   120
               TabIndex        =   238
               Top             =   2400
               Width           =   135
            End
            Begin VB.CommandButton cmdGotoRoom 
               Height          =   135
               Index           =   5
               Left            =   120
               TabIndex        =   235
               Top             =   2160
               Width           =   135
            End
            Begin VB.CommandButton cmdGotoRoom 
               Height          =   135
               Index           =   4
               Left            =   120
               TabIndex        =   232
               Top             =   1920
               Width           =   135
            End
            Begin VB.CommandButton cmdGotoRoom 
               Height          =   135
               Index           =   3
               Left            =   120
               TabIndex        =   229
               Top             =   1680
               Width           =   135
            End
            Begin VB.CommandButton cmdGotoRoom 
               Height          =   135
               Index           =   2
               Left            =   120
               TabIndex        =   226
               Top             =   1440
               Width           =   135
            End
            Begin VB.CommandButton cmdGotoRoom 
               Height          =   135
               Index           =   1
               Left            =   120
               TabIndex        =   223
               Top             =   1200
               Width           =   135
            End
            Begin VB.CommandButton cmdGotoRoom 
               Height          =   135
               Index           =   0
               Left            =   120
               TabIndex        =   220
               Top             =   960
               Width           =   135
            End
            Begin VB.TextBox txtRoomTrail 
               Height          =   285
               Index           =   19
               Left            =   4800
               TabIndex        =   279
               Top             =   3060
               Width           =   675
            End
            Begin VB.TextBox txtMapTrail 
               Height          =   285
               Index           =   19
               Left            =   4260
               TabIndex        =   278
               Top             =   3060
               Width           =   495
            End
            Begin VB.TextBox txtRoomTrail 
               Height          =   285
               Index           =   18
               Left            =   4800
               TabIndex        =   276
               Top             =   2820
               Width           =   675
            End
            Begin VB.TextBox txtMapTrail 
               Height          =   285
               Index           =   18
               Left            =   4260
               TabIndex        =   275
               Top             =   2820
               Width           =   495
            End
            Begin VB.TextBox txtRoomTrail 
               Height          =   285
               Index           =   17
               Left            =   4800
               TabIndex        =   273
               Top             =   2580
               Width           =   675
            End
            Begin VB.TextBox txtMapTrail 
               Height          =   285
               Index           =   17
               Left            =   4260
               TabIndex        =   272
               Top             =   2580
               Width           =   495
            End
            Begin VB.TextBox txtRoomTrail 
               Height          =   285
               Index           =   16
               Left            =   4800
               TabIndex        =   270
               Top             =   2340
               Width           =   675
            End
            Begin VB.TextBox txtMapTrail 
               Height          =   285
               Index           =   16
               Left            =   4260
               TabIndex        =   269
               Top             =   2340
               Width           =   495
            End
            Begin VB.TextBox txtRoomTrail 
               Height          =   285
               Index           =   15
               Left            =   4800
               TabIndex        =   267
               Top             =   2100
               Width           =   675
            End
            Begin VB.TextBox txtMapTrail 
               Height          =   285
               Index           =   15
               Left            =   4260
               TabIndex        =   266
               Top             =   2100
               Width           =   495
            End
            Begin VB.TextBox txtRoomTrail 
               Height          =   285
               Index           =   14
               Left            =   4800
               TabIndex        =   264
               Top             =   1860
               Width           =   675
            End
            Begin VB.TextBox txtMapTrail 
               Height          =   285
               Index           =   14
               Left            =   4260
               TabIndex        =   263
               Top             =   1860
               Width           =   495
            End
            Begin VB.TextBox txtRoomTrail 
               Height          =   285
               Index           =   13
               Left            =   4800
               TabIndex        =   261
               Top             =   1620
               Width           =   675
            End
            Begin VB.TextBox txtMapTrail 
               Height          =   285
               Index           =   13
               Left            =   4260
               TabIndex        =   260
               Top             =   1620
               Width           =   495
            End
            Begin VB.TextBox txtRoomTrail 
               Height          =   285
               Index           =   12
               Left            =   4800
               TabIndex        =   258
               Top             =   1380
               Width           =   675
            End
            Begin VB.TextBox txtMapTrail 
               Height          =   285
               Index           =   12
               Left            =   4260
               TabIndex        =   257
               Top             =   1380
               Width           =   495
            End
            Begin VB.TextBox txtRoomTrail 
               Height          =   285
               Index           =   11
               Left            =   4800
               TabIndex        =   255
               Top             =   1140
               Width           =   675
            End
            Begin VB.TextBox txtMapTrail 
               Height          =   285
               Index           =   11
               Left            =   4260
               TabIndex        =   254
               Top             =   1140
               Width           =   495
            End
            Begin VB.TextBox txtRoomTrail 
               Height          =   285
               Index           =   10
               Left            =   4800
               TabIndex        =   252
               Top             =   900
               Width           =   675
            End
            Begin VB.TextBox txtMapTrail 
               Height          =   285
               Index           =   10
               Left            =   4260
               TabIndex        =   251
               Top             =   900
               Width           =   495
            End
            Begin VB.TextBox txtRoomTrail 
               Height          =   285
               Index           =   9
               Left            =   1080
               TabIndex        =   249
               Top             =   3060
               Width           =   675
            End
            Begin VB.TextBox txtMapTrail 
               Height          =   285
               Index           =   9
               Left            =   540
               TabIndex        =   248
               Top             =   3060
               Width           =   495
            End
            Begin VB.TextBox txtRoomTrail 
               Height          =   285
               Index           =   8
               Left            =   1080
               TabIndex        =   246
               Top             =   2820
               Width           =   675
            End
            Begin VB.TextBox txtMapTrail 
               Height          =   285
               Index           =   8
               Left            =   540
               TabIndex        =   245
               Top             =   2820
               Width           =   495
            End
            Begin VB.TextBox txtRoomTrail 
               Height          =   285
               Index           =   7
               Left            =   1080
               TabIndex        =   243
               Top             =   2580
               Width           =   675
            End
            Begin VB.TextBox txtMapTrail 
               Height          =   285
               Index           =   7
               Left            =   540
               TabIndex        =   242
               Top             =   2580
               Width           =   495
            End
            Begin VB.TextBox txtRoomTrail 
               Height          =   285
               Index           =   6
               Left            =   1080
               TabIndex        =   240
               Top             =   2340
               Width           =   675
            End
            Begin VB.TextBox txtMapTrail 
               Height          =   285
               Index           =   6
               Left            =   540
               TabIndex        =   239
               Top             =   2340
               Width           =   495
            End
            Begin VB.TextBox txtRoomTrail 
               Height          =   285
               Index           =   5
               Left            =   1080
               TabIndex        =   237
               Top             =   2100
               Width           =   675
            End
            Begin VB.TextBox txtMapTrail 
               Height          =   285
               Index           =   5
               Left            =   540
               TabIndex        =   236
               Top             =   2100
               Width           =   495
            End
            Begin VB.TextBox txtRoomTrail 
               Height          =   285
               Index           =   4
               Left            =   1080
               TabIndex        =   234
               Top             =   1860
               Width           =   675
            End
            Begin VB.TextBox txtMapTrail 
               Height          =   285
               Index           =   4
               Left            =   540
               TabIndex        =   233
               Top             =   1860
               Width           =   495
            End
            Begin VB.TextBox txtRoomTrail 
               Height          =   285
               Index           =   3
               Left            =   1080
               TabIndex        =   231
               Top             =   1620
               Width           =   675
            End
            Begin VB.TextBox txtMapTrail 
               Height          =   285
               Index           =   3
               Left            =   540
               TabIndex        =   230
               Top             =   1620
               Width           =   495
            End
            Begin VB.TextBox txtRoomTrail 
               Height          =   285
               Index           =   2
               Left            =   1080
               TabIndex        =   228
               Top             =   1380
               Width           =   675
            End
            Begin VB.TextBox txtMapTrail 
               Height          =   285
               Index           =   2
               Left            =   540
               TabIndex        =   227
               Top             =   1380
               Width           =   495
            End
            Begin VB.TextBox txtRoomTrail 
               Height          =   285
               Index           =   1
               Left            =   1080
               TabIndex        =   225
               Top             =   1140
               Width           =   675
            End
            Begin VB.TextBox txtMapTrail 
               Height          =   285
               Index           =   1
               Left            =   540
               TabIndex        =   224
               Top             =   1140
               Width           =   495
            End
            Begin VB.TextBox txtRoomTrail 
               Height          =   285
               Index           =   0
               Left            =   1080
               TabIndex        =   222
               Top             =   900
               Width           =   675
            End
            Begin VB.TextBox txtMapTrail 
               Height          =   285
               Index           =   0
               Left            =   540
               TabIndex        =   221
               Top             =   900
               Width           =   495
            End
            Begin VB.Label Label69 
               Alignment       =   2  'Center
               Caption         =   "Room"
               Height          =   255
               Left            =   4800
               TabIndex        =   219
               Top             =   660
               Width           =   675
            End
            Begin VB.Label Label68 
               Alignment       =   2  'Center
               Caption         =   "Map"
               Height          =   255
               Left            =   4260
               TabIndex        =   218
               Top             =   660
               Width           =   495
            End
            Begin VB.Label Label67 
               Caption         =   "19"
               Height          =   255
               Left            =   4020
               TabIndex        =   299
               Top             =   3090
               Width           =   255
            End
            Begin VB.Label Label66 
               Caption         =   "18"
               Height          =   255
               Left            =   4005
               TabIndex        =   297
               Top             =   2850
               Width           =   255
            End
            Begin VB.Label Label65 
               Caption         =   "17"
               Height          =   255
               Left            =   4005
               TabIndex        =   295
               Top             =   2610
               Width           =   255
            End
            Begin VB.Label Label64 
               Caption         =   "16"
               Height          =   255
               Left            =   4005
               TabIndex        =   289
               Top             =   2370
               Width           =   255
            End
            Begin VB.Label Label63 
               Caption         =   "15"
               Height          =   255
               Left            =   4005
               TabIndex        =   287
               Top             =   2130
               Width           =   255
            End
            Begin VB.Label Label62 
               Caption         =   "14"
               Height          =   255
               Left            =   4005
               TabIndex        =   293
               Top             =   1890
               Width           =   255
            End
            Begin VB.Label Label61 
               Caption         =   "13"
               Height          =   255
               Left            =   4005
               TabIndex        =   292
               Top             =   1650
               Width           =   255
            End
            Begin VB.Label Label60 
               Caption         =   "12"
               Height          =   255
               Left            =   4005
               TabIndex        =   291
               Top             =   1410
               Width           =   255
            End
            Begin VB.Label Label59 
               Caption         =   "11"
               Height          =   255
               Left            =   4005
               TabIndex        =   290
               Top             =   1170
               Width           =   255
            End
            Begin VB.Label Label58 
               Caption         =   "10"
               Height          =   255
               Left            =   4005
               TabIndex        =   281
               Top             =   930
               Width           =   255
            End
            Begin VB.Label Label57 
               Caption         =   "9"
               Height          =   255
               Left            =   360
               TabIndex        =   298
               Top             =   3090
               Width           =   135
            End
            Begin VB.Label Label56 
               Caption         =   "8"
               Height          =   255
               Left            =   360
               TabIndex        =   296
               Top             =   2850
               Width           =   135
            End
            Begin VB.Label Label55 
               Caption         =   "7"
               Height          =   255
               Left            =   360
               TabIndex        =   294
               Top             =   2610
               Width           =   135
            End
            Begin VB.Label Label54 
               Caption         =   "6"
               Height          =   255
               Left            =   360
               TabIndex        =   288
               Top             =   2370
               Width           =   135
            End
            Begin VB.Label Label53 
               Caption         =   "5"
               Height          =   255
               Left            =   360
               TabIndex        =   286
               Top             =   2130
               Width           =   135
            End
            Begin VB.Label Label52 
               Caption         =   "4"
               Height          =   255
               Left            =   360
               TabIndex        =   285
               Top             =   1890
               Width           =   135
            End
            Begin VB.Label Label51 
               Caption         =   "3"
               Height          =   255
               Left            =   360
               TabIndex        =   284
               Top             =   1650
               Width           =   135
            End
            Begin VB.Label Label50 
               Caption         =   "2"
               Height          =   255
               Left            =   360
               TabIndex        =   283
               Top             =   1410
               Width           =   135
            End
            Begin VB.Label Label49 
               Caption         =   "1"
               Height          =   255
               Left            =   360
               TabIndex        =   282
               Top             =   1170
               Width           =   135
            End
            Begin VB.Label Label48 
               Alignment       =   2  'Center
               Caption         =   "Room"
               Height          =   255
               Left            =   1080
               TabIndex        =   217
               Top             =   660
               Width           =   675
            End
            Begin VB.Label Label47 
               Alignment       =   2  'Center
               Caption         =   "Map"
               Height          =   255
               Left            =   540
               TabIndex        =   216
               Top             =   660
               Width           =   495
            End
            Begin VB.Label Label46 
               Caption         =   "0"
               Height          =   255
               Left            =   360
               TabIndex        =   280
               Top             =   930
               Width           =   135
            End
         End
         Begin VB.TextBox txtCurrentRoom 
            Height          =   285
            Left            =   -72180
            TabIndex        =   214
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox txtCurrentMap 
            Height          =   285
            Left            =   -72840
            TabIndex        =   213
            Top             =   540
            Width           =   555
         End
         Begin VB.Frame frmSpells 
            Caption         =   "Spellbook"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4455
            Left            =   -74700
            TabIndex        =   97
            Top             =   360
            Width           =   6255
            Begin VB.CommandButton cmdPasteSpells 
               Caption         =   "&Paste Spellbook"
               Height          =   315
               Left            =   4080
               TabIndex        =   98
               Top             =   180
               Width           =   2055
            End
            Begin VB.CommandButton cmdSpellEditor 
               Caption         =   "Spell Editor"
               Height          =   255
               Left            =   2520
               TabIndex        =   105
               Top             =   4080
               Width           =   1215
            End
            Begin VB.CommandButton cmdClearAllSpell 
               Caption         =   "Clear All"
               Height          =   255
               Left            =   4800
               TabIndex        =   106
               Top             =   4080
               Width           =   1335
            End
            Begin VB.CommandButton cmdEditSpell 
               Caption         =   "Change &Spell"
               Height          =   255
               Left            =   120
               TabIndex        =   104
               Top             =   4080
               Width           =   1335
            End
            Begin VB.ListBox lstSpells 
               Height          =   3375
               ItemData        =   "frmUser.frx":09AA
               Left            =   120
               List            =   "frmUser.frx":09B1
               TabIndex        =   103
               Top             =   540
               Width           =   6015
            End
            Begin VB.Label Label31 
               Caption         =   "Short"
               Height          =   255
               Left            =   1560
               TabIndex        =   101
               Top             =   300
               Width           =   495
            End
            Begin VB.Label Label30 
               Caption         =   "Spell Name"
               Height          =   255
               Left            =   2280
               TabIndex        =   102
               Top             =   300
               Width           =   1095
            End
            Begin VB.Label Label29 
               Caption         =   "Spell #"
               Height          =   255
               Left            =   840
               TabIndex        =   100
               Top             =   300
               Width           =   615
            End
            Begin VB.Label Label28 
               Caption         =   "SB#"
               Height          =   255
               Left            =   180
               TabIndex        =   99
               Top             =   300
               Width           =   495
            End
         End
         Begin VB.Frame frmKeys 
            Caption         =   "Keys"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1920
            Left            =   -74700
            TabIndex        =   87
            Top             =   2865
            Width           =   6255
            Begin VB.CommandButton cmdPasteItems 
               Caption         =   "&Paste Items/Keys/Worn/$$"
               Height          =   255
               Index           =   1
               Left            =   3660
               TabIndex        =   88
               Top             =   180
               Width           =   2475
            End
            Begin VB.CommandButton cmdItemEditorKey 
               Caption         =   "Item Editor"
               Height          =   255
               Left            =   2520
               TabIndex        =   95
               Top             =   1560
               Width           =   1215
            End
            Begin VB.CommandButton cmdClearAllKey 
               Caption         =   "Clear All"
               Height          =   255
               Left            =   4800
               TabIndex        =   96
               Top             =   1560
               Width           =   1335
            End
            Begin VB.CommandButton cmdEditKey 
               Caption         =   "Change &Key"
               Height          =   255
               Left            =   120
               TabIndex        =   94
               Top             =   1560
               Width           =   1335
            End
            Begin VB.ListBox lstKeys 
               Height          =   1035
               ItemData        =   "frmUser.frx":09D1
               Left            =   120
               List            =   "frmUser.frx":09D8
               TabIndex        =   93
               Top             =   480
               Width           =   6015
            End
            Begin VB.Label Label27 
               Caption         =   "Key Name"
               Height          =   255
               Left            =   2310
               TabIndex        =   92
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label26 
               Caption         =   "Uses"
               Height          =   255
               Left            =   1590
               TabIndex        =   91
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Label25 
               Caption         =   "Item#"
               Height          =   255
               Left            =   885
               TabIndex        =   90
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label24 
               Caption         =   "INV#"
               Height          =   255
               Left            =   195
               TabIndex        =   89
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame frmItems 
            Caption         =   "Items"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2505
            Left            =   -74700
            TabIndex        =   77
            Top             =   360
            Width           =   6255
            Begin VB.CommandButton cmdPasteItems 
               Caption         =   "&Paste Items/Keys/Worn/$$"
               Height          =   255
               Index           =   0
               Left            =   3660
               TabIndex        =   78
               Top             =   180
               Width           =   2475
            End
            Begin VB.CommandButton cmdItemEditor 
               Caption         =   "Item Editor"
               Height          =   255
               Left            =   2520
               TabIndex        =   85
               Top             =   2160
               Width           =   1215
            End
            Begin VB.CommandButton cmdClearAllItem 
               Caption         =   "Clear All"
               Height          =   255
               Left            =   4800
               TabIndex        =   86
               Top             =   2160
               Width           =   1335
            End
            Begin VB.CommandButton cmdEditItem 
               Caption         =   "Change &Item"
               Height          =   255
               Left            =   120
               TabIndex        =   84
               Top             =   2160
               Width           =   1335
            End
            Begin VB.ListBox lstItems 
               Height          =   1620
               ItemData        =   "frmUser.frx":09F1
               Left            =   120
               List            =   "frmUser.frx":09F8
               TabIndex        =   83
               Top             =   480
               Width           =   6015
            End
            Begin VB.Label Label23 
               Caption         =   "Item Name"
               Height          =   255
               Left            =   2310
               TabIndex        =   82
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label22 
               Caption         =   "Uses"
               Height          =   255
               Left            =   1620
               TabIndex        =   81
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Label21 
               Caption         =   "Item#"
               Height          =   255
               Left            =   885
               TabIndex        =   80
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label20 
               Caption         =   "INV#"
               Height          =   255
               Left            =   180
               TabIndex        =   79
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "BBS Name"
            Height          =   645
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   3375
            Begin VB.TextBox txtBBSName 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   120
               Locked          =   -1  'True
               MaxLength       =   29
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   225
               Width           =   3135
            End
         End
         Begin VB.Frame frameGeneral 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Stats"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   3735
            Left            =   120
            TabIndex        =   12
            Top             =   1080
            Width           =   6615
            Begin VB.TextBox txtMagicResistance 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Left            =   6000
               TabIndex        =   44
               Top             =   3300
               Width           =   495
            End
            Begin VB.TextBox txtMartialArts 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Left            =   6000
               TabIndex        =   42
               Top             =   2940
               Width           =   495
            End
            Begin VB.TextBox txtTracking 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Left            =   6000
               TabIndex        =   40
               Top             =   2580
               Width           =   495
            End
            Begin VB.TextBox txtPicklocks 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Left            =   6000
               TabIndex        =   38
               Top             =   2220
               Width           =   495
            End
            Begin VB.TextBox txtTraps 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Left            =   6000
               TabIndex        =   36
               Top             =   1860
               Width           =   495
            End
            Begin VB.TextBox txtThievery 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Left            =   6000
               TabIndex        =   34
               Top             =   1500
               Width           =   495
            End
            Begin VB.TextBox txtStealth 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Left            =   6000
               TabIndex        =   32
               Top             =   1080
               Width           =   495
            End
            Begin VB.TextBox txtPerception 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Left            =   6000
               TabIndex        =   30
               Top             =   660
               Width           =   495
            End
            Begin VB.TextBox txtLives 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Left            =   5340
               TabIndex        =   26
               Top             =   300
               Width           =   375
            End
            Begin VB.TextBox txtCP 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Left            =   5910
               TabIndex        =   28
               Top             =   300
               Width           =   585
            End
            Begin VB.TextBox txtSpellcasting 
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Left            =   3420
               TabIndex        =   54
               Top             =   1860
               Width           =   900
            End
            Begin VB.TextBox txtLastName 
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Left            =   2400
               MaxLength       =   18
               TabIndex        =   16
               Top             =   300
               Width           =   1995
            End
            Begin VB.TextBox txtFirstName 
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Left            =   840
               MaxLength       =   10
               TabIndex        =   15
               Top             =   300
               Width           =   1455
            End
            Begin VB.TextBox txtCurrentMana 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF00&
               Height          =   315
               Left            =   825
               TabIndex        =   50
               Top             =   1860
               Width           =   675
            End
            Begin VB.TextBox txtMaxMana 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF00&
               Height          =   315
               Left            =   1680
               TabIndex        =   52
               Top             =   1860
               Width           =   675
            End
            Begin VB.TextBox txtMaxHP 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF00&
               Height          =   315
               Left            =   1680
               TabIndex        =   48
               Top             =   1500
               Width           =   675
            End
            Begin VB.TextBox txtCurrentHP 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF00&
               Height          =   315
               Left            =   825
               TabIndex        =   46
               Top             =   1500
               Width           =   675
            End
            Begin VB.TextBox txtExperience 
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Left            =   2820
               TabIndex        =   22
               Top             =   660
               Width           =   1575
            End
            Begin VB.TextBox txtLevel 
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Left            =   2940
               TabIndex        =   24
               Top             =   1080
               Width           =   1455
            End
            Begin VB.TextBox txtStat 
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Index           =   11
               Left            =   3720
               TabIndex        =   76
               Top             =   3300
               Width           =   615
            End
            Begin VB.TextBox txtStat 
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Index           =   10
               Left            =   3720
               TabIndex        =   70
               Top             =   2580
               Width           =   615
            End
            Begin VB.TextBox txtStat 
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Index           =   9
               Left            =   3720
               TabIndex        =   73
               Top             =   2940
               Width           =   615
            End
            Begin VB.TextBox txtStat 
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Index           =   8
               Left            =   1680
               TabIndex        =   61
               Top             =   2580
               Width           =   615
            End
            Begin VB.TextBox txtStat 
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Index           =   7
               Left            =   1680
               TabIndex        =   67
               Top             =   3300
               Width           =   615
            End
            Begin VB.TextBox txtStat 
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Index           =   6
               Left            =   1680
               TabIndex        =   64
               Top             =   2940
               Width           =   615
            End
            Begin VB.TextBox txtStat 
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Index           =   5
               Left            =   3000
               TabIndex        =   75
               Top             =   3300
               Width           =   615
            End
            Begin VB.TextBox txtStat 
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Index           =   4
               Left            =   3000
               TabIndex        =   69
               Top             =   2580
               Width           =   615
            End
            Begin VB.TextBox txtStat 
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Index           =   3
               Left            =   3000
               TabIndex        =   72
               Top             =   2940
               Width           =   615
            End
            Begin VB.TextBox txtStat 
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Index           =   2
               Left            =   960
               TabIndex        =   60
               Top             =   2580
               Width           =   615
            End
            Begin VB.TextBox txtStat 
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   66
               Top             =   3300
               Width           =   615
            End
            Begin VB.TextBox txtStat 
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
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Index           =   0
               Left            =   960
               TabIndex        =   63
               Top             =   2940
               Width           =   615
            End
            Begin VB.ComboBox cmbClasses 
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
               ForeColor       =   &H00FFFF00&
               Height          =   360
               ItemData        =   "frmUser.frx":0A12
               Left            =   840
               List            =   "frmUser.frx":0A14
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   1080
               Width           =   1455
            End
            Begin VB.ComboBox cmbRaces 
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
               ForeColor       =   &H00FFFF00&
               Height          =   360
               Left            =   840
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   660
               Width           =   1455
            End
            Begin VB.Label Label19 
               BackColor       =   &H00000000&
               Caption         =   "Magic Resistance:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   4560
               TabIndex        =   43
               Top             =   3300
               Width           =   1590
            End
            Begin VB.Label Label18 
               BackColor       =   &H00000000&
               Caption         =   "Martial Arts:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   4560
               TabIndex        =   41
               Top             =   2940
               Width           =   1035
            End
            Begin VB.Label Label17 
               BackColor       =   &H00000000&
               Caption         =   "Tracking:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   4560
               TabIndex        =   39
               Top             =   2580
               Width           =   825
            End
            Begin VB.Label Label16 
               BackColor       =   &H00000000&
               Caption         =   "Picklocks:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   4560
               TabIndex        =   37
               Top             =   2220
               Width           =   900
            End
            Begin VB.Label Label15 
               BackColor       =   &H00000000&
               Caption         =   "Traps:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   4560
               TabIndex        =   35
               Top             =   1860
               Width           =   555
            End
            Begin VB.Label Label14 
               BackColor       =   &H00000000&
               Caption         =   "Thievery:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   4560
               TabIndex        =   33
               Top             =   1500
               Width           =   810
            End
            Begin VB.Label Label13 
               BackColor       =   &H00000000&
               Caption         =   "Stealth:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   4560
               TabIndex        =   31
               Top             =   1080
               Width           =   675
            End
            Begin VB.Label Label12 
               BackColor       =   &H00000000&
               Caption         =   "Perception:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   4560
               TabIndex        =   29
               Top             =   660
               Width           =   990
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               Caption         =   "/"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF00&
               Height          =   375
               Left            =   5760
               TabIndex        =   27
               Top             =   255
               Width           =   135
            End
            Begin VB.Label Label10 
               BackColor       =   &H00000000&
               Caption         =   "Lives/CP:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   4560
               TabIndex        =   25
               Top             =   300
               Width           =   855
            End
            Begin VB.Label Label9 
               BackColor       =   &H00000000&
               Caption         =   "Spellcasting:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   2400
               TabIndex        =   53
               Top             =   1860
               Width           =   1110
            End
            Begin VB.Label Label8 
               BackColor       =   &H00000000&
               Caption         =   "Current"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   3720
               TabIndex        =   58
               Top             =   2340
               Width           =   630
            End
            Begin VB.Label Label7 
               BackColor       =   &H00000000&
               Caption         =   "Set"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   3000
               TabIndex        =   57
               Top             =   2340
               Width           =   600
            End
            Begin VB.Label Label6 
               BackColor       =   &H00000000&
               Caption         =   "Current"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   1680
               TabIndex        =   56
               Top             =   2340
               Width           =   630
            End
            Begin VB.Label Label5 
               BackColor       =   &H00000000&
               Caption         =   "Set"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   960
               TabIndex        =   55
               Top             =   2340
               Width           =   645
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               Caption         =   "/"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF00&
               Height          =   375
               Left            =   1500
               TabIndex        =   51
               Top             =   1800
               Width           =   135
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               Caption         =   "/"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF00&
               Height          =   375
               Index           =   0
               Left            =   1500
               TabIndex        =   47
               Top             =   1455
               Width           =   135
            End
            Begin VB.Label Label2 
               BackColor       =   &H00000000&
               Caption         =   "Name:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   14
               Top             =   300
               Width           =   555
            End
            Begin VB.Label Label1 
               BackColor       =   &H00000000&
               Caption         =   "Mana:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   49
               Top             =   1860
               Width           =   540
            End
            Begin VB.Label label 
               BackColor       =   &H00000000&
               Caption         =   "Hits:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Index           =   10
               Left            =   120
               TabIndex        =   45
               Top             =   1500
               Width           =   405
            End
            Begin VB.Label label 
               BackColor       =   &H00000000&
               Caption         =   "Exp:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Index           =   9
               Left            =   2400
               TabIndex        =   21
               Top             =   660
               Width           =   390
            End
            Begin VB.Label label 
               BackColor       =   &H00000000&
               Caption         =   "Level:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Index           =   8
               Left            =   2400
               TabIndex        =   23
               Top             =   1080
               Width           =   540
            End
            Begin VB.Label label 
               BackColor       =   &H00000000&
               Caption         =   "Charm:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Index           =   7
               Left            =   2400
               TabIndex        =   74
               Top             =   3300
               Width           =   600
            End
            Begin VB.Label label 
               BackColor       =   &H00000000&
               Caption         =   "Agility:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Index           =   6
               Left            =   2400
               TabIndex        =   68
               Top             =   2580
               Width           =   585
            End
            Begin VB.Label label 
               BackColor       =   &H00000000&
               Caption         =   "Health:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Index           =   5
               Left            =   2400
               TabIndex        =   71
               Top             =   2940
               Width           =   630
            End
            Begin VB.Label label 
               BackColor       =   &H00000000&
               Caption         =   "Strength:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   59
               Top             =   2580
               Width           =   795
            End
            Begin VB.Label label 
               BackColor       =   &H00000000&
               Caption         =   "Willpower:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   65
               Top             =   3300
               Width           =   900
            End
            Begin VB.Label label 
               BackColor       =   &H00000000&
               Caption         =   "Intellect:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   62
               Top             =   2940
               Width           =   765
            End
            Begin VB.Label label 
               BackColor       =   &H00000000&
               Caption         =   "Class:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   19
               Top             =   1080
               Width           =   525
            End
            Begin VB.Label label 
               BackColor       =   &H00000000&
               Caption         =   "Race:"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   17
               Top             =   660
               Width           =   525
            End
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   4395
            Left            =   -73800
            TabIndex        =   107
            Top             =   420
            Width           =   4485
            _ExtentX        =   7911
            _ExtentY        =   7752
            _Version        =   393216
            TabOrientation  =   3
            TabHeight       =   520
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Page 1"
            TabPicture(0)   =   "frmUser.frx":0A16
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "frmAblilities"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Page 2"
            TabPicture(1)   =   "frmUser.frx":0A32
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame2"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Page 3"
            TabPicture(2)   =   "frmUser.frx":0A4E
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Frame3"
            Tab(2).ControlCount=   1
            Begin VB.Frame Frame3 
               Caption         =   "Abilities"
               Height          =   4155
               Left            =   -74640
               TabIndex        =   177
               Top             =   120
               Width           =   3375
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   29
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   208
                  Top             =   3720
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   28
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   205
                  Top             =   3360
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   27
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   202
                  Top             =   3000
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   26
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   199
                  Top             =   2640
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   25
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   196
                  Top             =   2280
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   24
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   193
                  Top             =   1920
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   23
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   190
                  Top             =   1560
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   22
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   187
                  Top             =   1200
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   21
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   184
                  Top             =   840
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   29
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   210
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   3720
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   28
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   207
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   3360
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   27
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   204
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   3000
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   26
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   201
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   2640
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   25
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   198
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   2280
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   24
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   195
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   1920
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   23
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   192
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   1560
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   22
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   189
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   1200
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   21
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   186
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   840
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   20
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   183
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   480
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   20
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   181
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
                  Index           =   29
                  Left            =   720
                  TabIndex        =   209
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
                  Index           =   28
                  Left            =   720
                  TabIndex        =   206
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
                  Index           =   27
                  Left            =   720
                  TabIndex        =   203
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
                  Index           =   26
                  Left            =   720
                  TabIndex        =   200
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
                  Index           =   25
                  Left            =   720
                  TabIndex        =   197
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
                  Index           =   24
                  Left            =   720
                  TabIndex        =   194
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
                  Index           =   23
                  Left            =   720
                  TabIndex        =   191
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
                  Index           =   22
                  Left            =   720
                  TabIndex        =   188
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
                  Index           =   21
                  Left            =   720
                  TabIndex        =   185
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
                  Index           =   20
                  Left            =   720
                  TabIndex        =   182
                  Text            =   "empty"
                  Top             =   480
                  Width           =   1815
               End
               Begin VB.Label Label1 
                  Caption         =   "#"
                  Height          =   255
                  Index           =   3
                  Left            =   120
                  TabIndex        =   178
                  Top             =   240
                  Width           =   495
               End
               Begin VB.Label Label2 
                  Caption         =   "Name"
                  Height          =   255
                  Index           =   3
                  Left            =   720
                  TabIndex        =   179
                  Top             =   240
                  Width           =   1815
               End
               Begin VB.Label Label3 
                  Caption         =   "Value"
                  Height          =   255
                  Index           =   3
                  Left            =   2640
                  TabIndex        =   180
                  Top             =   240
                  Width           =   615
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Abilities"
               Height          =   4155
               Left            =   -74640
               TabIndex        =   143
               Top             =   120
               Width           =   3375
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
                  TabIndex        =   175
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
                  TabIndex        =   172
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
                  TabIndex        =   169
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
                  TabIndex        =   166
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
                  TabIndex        =   163
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
                  TabIndex        =   160
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
                  TabIndex        =   157
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
                  TabIndex        =   154
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
                  TabIndex        =   151
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
                  TabIndex        =   148
                  Text            =   "empty"
                  Top             =   480
                  Width           =   1815
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   19
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   174
                  Top             =   3720
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   19
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   176
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   3720
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   18
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   173
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   3360
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   17
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   170
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   3000
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   16
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   167
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   2640
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   15
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   164
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   2280
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   14
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   161
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   1920
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   13
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   158
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   1560
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   12
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   155
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   1200
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   11
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   152
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   840
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   10
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   149
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   480
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   18
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   171
                  Top             =   3360
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   17
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   168
                  Top             =   3000
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   16
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   165
                  Top             =   2640
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   15
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   162
                  Top             =   2280
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   14
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   159
                  Top             =   1920
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   13
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   156
                  Top             =   1560
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   12
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   153
                  Top             =   1200
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   11
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   150
                  Top             =   840
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   10
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   147
                  Top             =   480
                  Width           =   495
               End
               Begin VB.Label Label1 
                  Caption         =   "#"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   144
                  Top             =   240
                  Width           =   495
               End
               Begin VB.Label Label2 
                  Caption         =   "Name"
                  Height          =   255
                  Index           =   2
                  Left            =   720
                  TabIndex        =   145
                  Top             =   240
                  Width           =   1815
               End
               Begin VB.Label Label3 
                  Caption         =   "Value"
                  Height          =   255
                  Index           =   2
                  Left            =   2640
                  TabIndex        =   146
                  Top             =   240
                  Width           =   615
               End
            End
            Begin VB.Frame frmAblilities 
               Caption         =   "Abilities"
               Height          =   4155
               Left            =   360
               TabIndex        =   109
               Top             =   120
               Width           =   3375
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   9
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   140
                  Top             =   3720
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   8
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   137
                  Top             =   3360
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   7
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   134
                  Top             =   3000
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   6
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   131
                  Top             =   2640
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   5
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   128
                  Top             =   2280
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   4
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   125
                  Top             =   1920
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   3
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   122
                  Top             =   1560
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   2
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   119
                  Top             =   1200
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   1
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   116
                  Top             =   840
                  Width           =   495
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   5
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   130
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   2280
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   6
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   133
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   2640
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   7
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   136
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   3000
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   8
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   139
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   3360
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   9
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   142
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   3720
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   0
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   115
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   480
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   1
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   118
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   840
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   2
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   121
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   1200
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   3
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   124
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   1560
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityB 
                  Height          =   315
                  Index           =   4
                  Left            =   2640
                  MaxLength       =   5
                  TabIndex        =   127
                  ToolTipText     =   "Enter the value for the ability here."
                  Top             =   1920
                  Width           =   615
               End
               Begin VB.TextBox txtAbilityA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   0
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   113
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
                  Index           =   0
                  Left            =   720
                  TabIndex        =   114
                  Text            =   "empty"
                  Top             =   480
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
                  TabIndex        =   117
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
                  Index           =   2
                  Left            =   720
                  TabIndex        =   120
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
                  Index           =   3
                  Left            =   720
                  TabIndex        =   123
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
                  Index           =   4
                  Left            =   720
                  TabIndex        =   126
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
                  Index           =   5
                  Left            =   720
                  TabIndex        =   129
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
                  Index           =   6
                  Left            =   720
                  TabIndex        =   132
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
                  Index           =   7
                  Left            =   720
                  TabIndex        =   135
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
                  Index           =   8
                  Left            =   720
                  TabIndex        =   138
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
                  Index           =   9
                  Left            =   720
                  TabIndex        =   141
                  Text            =   "empty"
                  Top             =   3720
                  Width           =   1815
               End
               Begin VB.Label Label1 
                  Caption         =   "#"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   110
                  Top             =   240
                  Width           =   495
               End
               Begin VB.Label Label2 
                  Caption         =   "Name"
                  Height          =   255
                  Index           =   1
                  Left            =   720
                  TabIndex        =   111
                  Top             =   240
                  Width           =   1815
               End
               Begin VB.Label Label3 
                  Caption         =   "Value"
                  Height          =   255
                  Index           =   1
                  Left            =   2640
                  TabIndex        =   112
                  Top             =   240
                  Width           =   615
               End
            End
         End
         Begin VB.Label Label76 
            AutoSize        =   -1  'True
            Caption         =   "Hit Point Rolls"
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
            Index           =   12
            Left            =   -70020
            TabIndex        =   577
            Top             =   1605
            Width           =   870
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   48
            Left            =   -69240
            TabIndex        =   571
            Top             =   4260
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   47
            Left            =   -70140
            TabIndex        =   570
            Top             =   4260
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   46
            Left            =   -71040
            TabIndex        =   569
            Top             =   4260
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   45
            Left            =   -71940
            TabIndex        =   568
            Top             =   4260
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   44
            Left            =   -72840
            TabIndex        =   567
            Top             =   4260
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   43
            Left            =   -73740
            TabIndex        =   566
            Top             =   4260
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   42
            Left            =   -74640
            TabIndex        =   565
            Top             =   4260
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   41
            Left            =   -69240
            TabIndex        =   557
            Top             =   3660
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   40
            Left            =   -70140
            TabIndex        =   556
            Top             =   3660
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   39
            Left            =   -71040
            TabIndex        =   555
            Top             =   3660
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   38
            Left            =   -71940
            TabIndex        =   554
            Top             =   3660
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   37
            Left            =   -72840
            TabIndex        =   553
            Top             =   3660
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   36
            Left            =   -73740
            TabIndex        =   552
            Top             =   3660
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   35
            Left            =   -74640
            TabIndex        =   551
            Top             =   3660
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   34
            Left            =   -69240
            TabIndex        =   543
            Top             =   3060
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   33
            Left            =   -70140
            TabIndex        =   542
            Top             =   3060
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   32
            Left            =   -71040
            TabIndex        =   541
            Top             =   3060
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   31
            Left            =   -71940
            TabIndex        =   540
            Top             =   3060
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   30
            Left            =   -72840
            TabIndex        =   539
            Top             =   3060
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   29
            Left            =   -73740
            TabIndex        =   538
            Top             =   3060
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   28
            Left            =   -74640
            TabIndex        =   537
            Top             =   3060
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   27
            Left            =   -69240
            TabIndex        =   529
            Top             =   2460
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   26
            Left            =   -70140
            TabIndex        =   528
            Top             =   2460
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   25
            Left            =   -71040
            TabIndex        =   527
            Top             =   2460
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   24
            Left            =   -71940
            TabIndex        =   526
            Top             =   2460
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   23
            Left            =   -72840
            TabIndex        =   525
            Top             =   2460
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   22
            Left            =   -73740
            TabIndex        =   524
            Top             =   2460
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   21
            Left            =   -74640
            TabIndex        =   523
            Top             =   2460
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   20
            Left            =   -69240
            TabIndex        =   515
            Top             =   1860
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   19
            Left            =   -70140
            TabIndex        =   514
            Top             =   1860
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   18
            Left            =   -71040
            TabIndex        =   513
            Top             =   1860
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   17
            Left            =   -71940
            TabIndex        =   512
            Top             =   1860
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   16
            Left            =   -72840
            TabIndex        =   511
            Top             =   1860
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   15
            Left            =   -73740
            TabIndex        =   510
            Top             =   1860
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   14
            Left            =   -74640
            TabIndex        =   509
            Top             =   1860
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   13
            Left            =   -69240
            TabIndex        =   501
            Top             =   1260
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   12
            Left            =   -70140
            TabIndex        =   500
            Top             =   1260
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   11
            Left            =   -71040
            TabIndex        =   499
            Top             =   1260
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   10
            Left            =   -71940
            TabIndex        =   498
            Top             =   1260
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   9
            Left            =   -72840
            TabIndex        =   497
            Top             =   1260
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   8
            Left            =   -73740
            TabIndex        =   496
            Top             =   1260
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   7
            Left            =   -74640
            TabIndex        =   495
            Top             =   1260
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   6
            Left            =   -69240
            TabIndex        =   494
            Top             =   660
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   -70140
            TabIndex        =   493
            Top             =   660
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   4
            Left            =   -71040
            TabIndex        =   492
            Top             =   660
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   -71940
            TabIndex        =   491
            Top             =   660
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   -72840
            TabIndex        =   490
            Top             =   660
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   -73740
            TabIndex        =   489
            Top             =   660
            Width           =   705
         End
         Begin VB.Label lblUserUnknowns 
            AutoSize        =   -1  'True
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   -74640
            TabIndex        =   488
            Top             =   660
            Width           =   705
         End
         Begin VB.Label Label76 
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
            Height          =   315
            Index           =   11
            Left            =   -72600
            TabIndex        =   373
            Top             =   1140
            Width           =   135
         End
         Begin VB.Label Label76 
            AutoSize        =   -1  'True
            Caption         =   "Copper"
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
            Index           =   10
            Left            =   -70020
            TabIndex        =   389
            Top             =   3780
            Width           =   615
         End
         Begin VB.Label Label76 
            AutoSize        =   -1  'True
            Caption         =   "Silver"
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
            Index           =   9
            Left            =   -70020
            TabIndex        =   387
            Top             =   3420
            Width           =   495
         End
         Begin VB.Label Label76 
            AutoSize        =   -1  'True
            Caption         =   "Gold"
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
            Index           =   8
            Left            =   -70020
            TabIndex        =   385
            Top             =   3060
            Width           =   405
         End
         Begin VB.Label Label76 
            AutoSize        =   -1  'True
            Caption         =   "Platinum"
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
            Index           =   7
            Left            =   -70020
            TabIndex        =   383
            Top             =   2700
            Width           =   735
         End
         Begin VB.Label Label76 
            AutoSize        =   -1  'True
            Caption         =   "Runic"
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
            Index           =   6
            Left            =   -70020
            TabIndex        =   381
            Top             =   2340
            Width           =   510
         End
         Begin VB.Label Label76 
            AutoSize        =   -1  'True
            Caption         =   "Broadcast:"
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
            Index           =   5
            Left            =   -71340
            TabIndex        =   379
            Top             =   1200
            Width           =   930
         End
         Begin VB.Label Label76 
            AutoSize        =   -1  'True
            Caption         =   "Evil Points:"
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
            Left            =   -71340
            TabIndex        =   377
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label76 
            AutoSize        =   -1  'True
            Caption         =   "Suicide:"
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
            Left            =   -71340
            TabIndex        =   375
            Top             =   480
            Width           =   705
         End
         Begin VB.Label Label76 
            AutoSize        =   -1  'True
            Caption         =   "Encumberance:"
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
            Left            =   -74820
            TabIndex        =   371
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label76 
            AutoSize        =   -1  'True
            Caption         =   "Gang:"
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
            Left            =   -74820
            TabIndex        =   369
            Top             =   840
            Width           =   525
         End
         Begin VB.Label Label76 
            AutoSize        =   -1  'True
            Caption         =   "Title:"
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
            Left            =   -74820
            TabIndex        =   367
            Top             =   480
            Width           =   450
         End
         Begin VB.Label Label75 
            AutoSize        =   -1  'True
            Caption         =   "Weapon:"
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
            Left            =   -74280
            TabIndex        =   302
            Top             =   1200
            Width           =   780
         End
         Begin VB.Label Label74 
            AutoSize        =   -1  'True
            Caption         =   "Current Room:"
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
            Left            =   -74280
            TabIndex        =   212
            Top             =   600
            Width           =   1230
         End
      End
      Begin VB.CommandButton cmdDiscard 
         Caption         =   "Dis&card"
         Height          =   255
         Left            =   5940
         TabIndex        =   8
         Top             =   0
         Width           =   915
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   255
         Left            =   4860
         TabIndex        =   7
         Top             =   0
         Width           =   1035
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Duplicate"
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   915
      End
      Begin VB.CheckBox chkAutoSave 
         Caption         =   "Auto-Save"
         Height          =   195
         Left            =   3720
         TabIndex        =   6
         ToolTipText     =   "Turn this on to auto-save users when switching records (carefule if using live dats!)"
         Top             =   15
         Width           =   1095
      End
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   60
      TabIndex        =   1
      Top             =   180
      Width           =   3135
   End
   Begin MSComctlLib.ListView lvDatabase 
      Height          =   4815
      Left            =   60
      TabIndex        =   2
      Top             =   540
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   8493
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
      Left            =   2640
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
   End
   Begin VB.Label Label72 
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
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim UserItem(0 To 99) As Long
Dim UserItemUses(0 To 99) As Integer
Dim UserKey(0 To 49) As Long
Dim UserKeyUses(0 To 49) As Integer
Dim UserSpell(0 To 99) As Long
Dim bLoaded As Boolean
'Dim bDontSetup As Boolean
Dim sCurrentRecord As String * 30

Private Sub Check1_Click()

End Sub

'Private Declare Function CalcExpNeeded Lib "lltmmudxp" (ByVal Level As Integer, ByVal Chart As Integer) As Currency


Private Sub chkNoRoomNames_Click()

On Error GoTo error:
Dim x As Integer

For x = 0 To 19
    Call txtMapTrail_Change(x)
Next x

out:
Exit Sub
error:
Call HandleError("chkNoRoomNames_Click")
Resume out:
End Sub

Private Sub cmdAbilsClear_Click()
Dim x As Integer
On Error GoTo error:

For x = 0 To 29
    txtAbilityA(x).Text = 0
    txtAbilityB(x).Text = 0
Next x

out:
Exit Sub
error:
Call HandleError("cmdAbilsClear_Click")
Resume out:

End Sub

Private Sub cmdCalcExp_Click()
On Error GoTo error:

frmExpCalc.Show
frmExpCalc.SetFocus

Call frmExpCalc.CalcBy(cmbClasses.ListIndex, cmbRaces.ListIndex, Val(txtLevel.Text))

out:
Exit Sub
error:
Call HandleError("cmdCalcExp_Click")
Resume out:
End Sub

Private Sub cmdImportGo_Click(Index As Integer)
txtSearch.Enabled = True
lvDatabase.Enabled = True
framImport.Visible = False
If Index = 0 Then Call ImportFromClipboard
End Sub

Private Sub cmdPasteChar_Click()
On Error GoTo error:

Call PasteCharacter

Exit Sub

error:
Call HandleError
End Sub

Private Sub PasteCharacter()
On Error GoTo error:
Dim nStatus As Integer, sSearch As String, x As Long, y As Integer, x2 As Integer
Dim sRaceName As String, sClassName As String, sChar As String
Dim sStr As String
'x = current position in string
'y = length of next possible (current) string match
'x2 = starting point of field match

Me.Enabled = False
Load frmUserPaste
frmUserPaste.Caption = "Paste Character Class/Race/Level/Exp"
frmUserPaste.txtText = ""
frmUserPaste.Tag = "-1"

sSearch = Clipboard.GetText
If Not sSearch = "" Then frmUserPaste.txtText = sSearch

frmUserPaste.Show vbModal, frmMain
If frmUserPaste.Tag = "-1" Then GoTo canceled:

sSearch = frmUserPaste.txtText.Text

Unload frmUserPaste

If Len(sSearch) < 10 Then GoTo canceled:

If Not InStr(1, sSearch, "Race: ") = 0 Then
    x = InStr(1, sSearch, "Race: ") + 6 '6=len("race: ")
    y = InStr(x, sSearch, "Exp:") 'exp is the next thing in the string for stats
    If y > x + 15 Then y = 0 'just incase "exp:" is somewhere way down in the paste
    If y > 0 Then
        If InStr(1, LTrim(RTrim(Mid(sSearch, x, y - x))), Chr(10)) > 0 Then y = 0
    End If
    If y = 0 Then y = InStr(x, sSearch, Chr(13))
    If y = 0 Then y = InStr(x, sSearch, Chr(10))
    If y > x Then sRaceName = LTrim(RTrim(Mid(sSearch, x, y - x)))
End If

If Not InStr(1, sSearch, "Class: ") = 0 Then
    x = InStr(1, sSearch, "Class: ") + 7
    y = InStr(x, sSearch, "Level:")
    If y > x + 15 Then y = 0
    If y > 0 Then
        If InStr(1, LTrim(RTrim(Mid(sSearch, x, y - x))), Chr(10)) > 0 Then y = 0
    End If
    If y = 0 Then y = InStr(x, sSearch, Chr(13))
    If y = 0 Then y = InStr(x, sSearch, Chr(10))
    If y > x Then sClassName = LTrim(RTrim(Mid(sSearch, x, y - x)))
End If


sStr = "Intellect: "
x = InStr(1, sSearch, sStr)
If x > 0 Then
    x = x + Len(sStr)
    sStr = GetNextNumbers(x, sSearch)
    If Not Val(sStr) < 1 Then
        txtStat(0).Text = sStr
        txtStat(6).Text = sStr
    End If
End If
sStr = "Willpower: "
x = InStr(1, sSearch, sStr)
If x > 0 Then
    x = x + Len(sStr)
    sStr = GetNextNumbers(x, sSearch)
    If Not Val(sStr) < 1 Then
        txtStat(1).Text = sStr
        txtStat(7).Text = sStr
    End If
End If
sStr = "Strength: "
x = InStr(1, sSearch, sStr)
If x > 0 Then
    x = x + Len(sStr)
    sStr = GetNextNumbers(x, sSearch)
    If Not Val(sStr) < 1 Then
        txtStat(2).Text = sStr
        txtStat(8).Text = sStr
    End If
End If
sStr = "Health: "
x = InStr(1, sSearch, sStr)
If x > 0 Then
    x = x + Len(sStr)
    sStr = GetNextNumbers(x, sSearch)
    If Not Val(sStr) < 1 Then
        txtStat(3).Text = sStr
        txtStat(9).Text = sStr
    End If
End If
sStr = "Agility: "
x = InStr(1, sSearch, sStr)
If x > 0 Then
    x = x + Len(sStr)
    sStr = GetNextNumbers(x, sSearch)
    If Not Val(sStr) < 1 Then
        txtStat(4).Text = sStr
        txtStat(10).Text = sStr
    End If
End If
sStr = "Charm: "
x = InStr(1, sSearch, sStr)
If x > 0 Then
    x = x + Len(sStr)
    sStr = GetNextNumbers(x, sSearch)
    If Not Val(sStr) < 1 Then
        txtStat(5).Text = sStr
        txtStat(11).Text = sStr
    End If
End If

sStr = "Exp: "
x = InStr(1, sSearch, sStr)
If x > 0 Then
    x = x + Len(sStr)
    sStr = GetNextNumbers(x, sSearch)
    If Not Val(sStr) < 0 Then txtExperience.Text = sStr
End If

sStr = "Level: "
x = InStr(1, sSearch, sStr)
If x > 0 Then
    x = x + Len(sStr)
    sStr = GetNextNumbers(x, sSearch)
    If Not Val(sStr) < 1 Then txtLevel.Text = sStr
End If

sStr = "Hits: "
x = InStr(1, sSearch, sStr)
If x > 0 Then
    x = x + Len(sStr)
    sStr = GetNextNumbers(x, sSearch)
    If Not Val(sStr) < 1 Then txtCurrentHP.Text = sStr
End If

sStr = "Mana: "
x = InStr(1, sSearch, sStr)
If x > 0 Then
    x = x + Len(sStr)
    sStr = GetNextNumbers(x, sSearch)
    If Not Val(sStr) < 1 Then txtCurrentMana.Text = sStr
End If

sStr = "Kai: "
x = InStr(1, sSearch, sStr)
If x > 0 Then
    x = x + Len(sStr)
    sStr = GetNextNumbers(x, sSearch)
    If Not Val(sStr) < 1 Then txtCurrentMana.Text = sStr
End If

sStr = "Lives/CP:"
x = InStr(1, sSearch, sStr)
If x > 0 Then
    x = x + Len(sStr)
    sStr = GetNextNumbers(x, sSearch)
    If Not Val(sStr) < 1 Then txtLives.Text = sStr
End If

sStr = "CP:"
x = InStr(1, sSearch, sStr)
If x > 0 Then
    x = x + Len(sStr)
    
    Do Until x > Len(sSearch)
        sChar = Mid(sSearch, x, 1)
        Select Case sChar
            Case "/": Exit Do
            Case Else:
        End Select
        x = x + 1
    Loop
    x = x + 1
    
    If x <= Len(sSearch) Then
        txtCp.Text = GetNextNumbers(x, sSearch)
    End If
End If

If Not sRaceName = "" Then
    If cmbRaces.ListCount > 0 Then
        For x = 0 To cmbRaces.ListCount - 1
            If cmbRaces.List(x) = sRaceName Then
                cmbRaces.ListIndex = x
            End If
        Next
    End If
End If

If Not sClassName = "" Then
    If cmbClasses.ListCount > 0 Then
        For x = 0 To cmbClasses.ListCount - 1
            If cmbClasses.List(x) = sClassName Then
                cmbClasses.ListIndex = x
            End If
        Next
    End If
End If

canceled:
Me.Enabled = True
Exit Sub
error:
Call HandleError
Me.Enabled = True
End Sub

Private Function GetNextNumbers(ByVal nStart As Long, sSearchString As String) As String
Dim y As Long, sChar As String
On Error GoTo error:

y = nStart
Do Until y > Len(sSearchString)
    sChar = Mid(sSearchString, y, 1)
    Select Case sChar
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", " "
        Case "*":
            nStart = y + 1
        Case Else: Exit Do
    End Select
    y = y + 1
Loop
If y > nStart Then
    GetNextNumbers = Val(Mid(sSearchString, nStart, y - nStart))
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetNextNumbers")
Resume out:
End Function
Private Sub PasteItems()
On Error GoTo error:
Dim sSearch As String, sText As String, sChar As String, x As Integer, y As Integer, x2 As Integer
Dim nStatus As Integer, sAmount As String, bResult As Boolean, nMaxSItems As Integer
Dim nCurKey As Integer, nCurItem As Integer, nCurWorn As Integer
Dim sEquipLoc(0 To 15) As String, sItems(0 To 199) As String, nOrigMatch As Long
Dim sFoundItems(0 To 199, 1 To 2) As String, sItemAmounts(0 To 199) As String

'x = current position in string
'y = length of next possible (current) string match
'x2 = starting point of field match

Me.Enabled = False
Load frmUserPaste
frmUserPaste.Caption = "Paste Items/Keys/Worn/$$"
frmUserPaste.txtText = ""
frmUserPaste.Tag = "-1"

sSearch = Clipboard.GetText
If Not sSearch = "" Then frmUserPaste.txtText = sSearch

frmUserPaste.Show vbModal, frmMain
If frmUserPaste.Tag = "-1" Then GoTo canceled:

sSearch = frmUserPaste.txtText.Text

Unload frmUserPaste

If Len(sSearch) < 10 Then GoTo canceled:

Me.MousePointer = vbHourglass

x = 1
y = 1
x2 = -1
sAmount = ""
Do Until x + y > Len(sSearch) + 1
    
    sChar = Mid(sSearch, x + y - 1, 1)
    
    bResult = TestPasteChar(sChar)
    If bResult = False Then GoTo next_y:
    
    sText = RemoveCharacter(sText & sChar, " ")
    'If Right(sText, 2) = "  " Then sText = Left(sText, Len(sText) - 1)
    
    If Not InStr(1, LCase(sText), "isequippedwith:") = 0 Then
        sText = Left(sText, Len(sText) - Len("isequippedwith:"))
        GoTo store_item:
    ElseIf Not InStr(1, LCase(sText), "arecarrying") = 0 Then
        sText = Left(sText, Len(sText) - Len("arecarrying"))
        GoTo store_item:
    ElseIf Not InStr(1, LCase(sText), "followingkeys:") = 0 Then
        sText = Left(sText, Len(sText) - Len("followingkeys:"))
        GoTo store_item:
    ElseIf Not InStr(1, LCase(sText), "wealth:") = 0 Then
        sText = Left(sText, Len(sText) - Len("wealth:"))
        GoTo store_item:
    End If
    
    If IsNumeric(sChar) Then
        If Len(sText) = Len(sAmount) + 1 Then
            'this IF is to ensure that the amount only comes before the item name and not within
            sAmount = sAmount & sChar
        End If
    End If
    
    Select Case sChar
        Case ",", ".":
            sText = Left(sText, Len(sText) - 1)
            GoTo store_item:
            
        Case "(":
            x2 = Len(sText)
            
        Case ")":
            If x2 = -1 Then GoTo clear:
            
            Select Case UCase(Mid(sText, x2 + 1, Len(sText) - x2 - 1))
                Case "HEAD": sEquipLoc(0) = Left(sText, x2 - 1)
                Case "EARS": sEquipLoc(1) = Left(sText, x2 - 1)
                Case "NECK": sEquipLoc(2) = Left(sText, x2 - 1)
                Case "BACK": sEquipLoc(3) = Left(sText, x2 - 1)
                Case "TORSO": sEquipLoc(4) = Left(sText, x2 - 1)
                Case "ARMS": sEquipLoc(5) = Left(sText, x2 - 1)
                Case "WRIST": sEquipLoc(6) = Left(sText, x2 - 1)
                Case "WAIST": sEquipLoc(7) = Left(sText, x2 - 1)
                Case "FINGER":
                    If sEquipLoc(9) = "" Then
                        sEquipLoc(9) = Left(sText, x2 - 1)
                    Else
                        sEquipLoc(10) = Left(sText, x2 - 1)
                    End If
                Case "HANDS": sEquipLoc(8) = Left(sText, x2 - 1)
                Case "LEGS": sEquipLoc(11) = Left(sText, x2 - 1)
                Case "FEET": sEquipLoc(12) = Left(sText, x2 - 1)
                Case "WORN": sEquipLoc(13) = Left(sText, x2 - 1)
                Case "OFF-HAND": sEquipLoc(14) = Left(sText, x2 - 1)
                Case "WEAPONHAND": sEquipLoc(15) = Left(sText, x2 - 1)
                Case "TWOHANDED": sEquipLoc(15) = Left(sText, x2 - 1)
            End Select
            
            GoTo clear:
    End Select
    
    If (x + y + 1) > (Len(sSearch) + 1) Then GoTo store_item:
    
GoTo next_y:

store_item:
If Right(sText, Len(sText) - Len(sAmount)) = "platinumpieces" Then
    txtPlatinum.Text = sAmount
    GoTo clear:
ElseIf Right(sText, Len(sText) - Len(sAmount)) = "goldcrowns" Then
    txtGold.Text = sAmount
    GoTo clear:
ElseIf Right(sText, Len(sText) - Len(sAmount)) = "silvernobles" Then
    txtSilver.Text = sAmount
    GoTo clear:
ElseIf Right(sText, Len(sText) - Len(sAmount)) = "copperfarthings" Then
    txtCopper.Text = sAmount
    GoTo clear:
ElseIf Left(Right(sText, Len(sText) - Len(sAmount)), Len("runiccoin")) = "runiccoin" Then
    txtRunic.Text = sAmount
    GoTo clear:
End If

If Not sAmount = "" And Not sText = "" Then
    sItems(nCurItem) = Right(sText, Len(sText) - Len(sAmount))
    sItemAmounts(nCurItem) = sAmount
Else
    sItems(nCurItem) = sText
End If
If nCurItem < 199 Then nCurItem = nCurItem + 1


clear:
sAmount = ""
sText = ""
x = x + y
y = 0
x2 = -1

next_y:
    y = y + 1
Loop

'RESET STATS:
x = MsgBox("Clear item and key lists before loading?  If you choose not to," & vbCrLf _
        & "then anything found will be /added/ to the current inventory.  If any" & vbCrLf _
        & "worn items are found, that list will be cleared and reloaded either way." _
        , vbYesNoCancel + vbQuestion + vbDefaultButton1, "Clear Items/Keys?")
If x = vbCancel Then
    GoTo canceled:
ElseIf x = vbYes Then
    Call cmdClearAllItem_Click
    Call cmdClearAllKey_Click
    Call cmdClearWorn_Click
Else
    For x = 0 To 15
        If Not sEquipLoc(x) = "" Then
            Call cmdClearWorn_Click
            Exit For
        End If
    Next
End If

'DISTRO DATA COLLECTED:
nMaxSItems = nCurItem
nCurItem = 0

'add to inven
nStatus = BTRCALL(BGETFIRST, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error getting first item -- " & BtrieveErrorCode(nStatus)
    Exit Sub
End If
Do While nStatus = 0
    Call ItemRowToStruct(Itemdatabuf.buf)
    
    sText = RemoveCharacter(ClipNull(Itemrec.Name), " ")
    If sText = "" Then GoTo skip_item:
    
    For x = 0 To 15
        If sText = sEquipLoc(x) Then
            If x = 15 Then 'weapon
                txtWeaponNumber.Text = Itemrec.Number
            Else
                txtWornItem(nCurWorn).Text = Itemrec.Number
                nCurWorn = nCurWorn + 1
            End If
retry_worn_item:
            If nCurItem < 100 Then
                If Not UserItem(nCurItem) = 0 Then nCurItem = nCurItem + 1: GoTo retry_worn_item:
                UserItem(nCurItem) = Itemrec.Number
                If Itemrec.Uses > 0 Then UserItemUses(nCurItem) = Itemrec.Uses
                nCurItem = nCurItem + 1
            End If
            sEquipLoc(x) = ""
        End If
    Next x
    
    For x = 0 To nMaxSItems
        If sText = sItems(x) Then
            For y = 0 To IIf(Val(sItemAmounts(x)) = 0, 0, Val(sItemAmounts(x)) - 1)
                If Itemrec.Type = 7 Then 'key
retry_key:
                    If nCurKey < 50 Then
                        If Not UserKey(nCurKey) = 0 Then nCurKey = nCurKey + 1: GoTo retry_key:
                        UserKey(nCurKey) = Itemrec.Number
                        UserKeyUses(nCurKey) = Itemrec.Uses
                        nCurKey = nCurKey + 1
                    End If
                Else
retry_item:
                    If nCurItem < 100 Then
                        If Not UserItem(nCurItem) = 0 Then nCurItem = nCurItem + 1: GoTo retry_item:
                        UserItem(nCurItem) = Itemrec.Number
                        If Itemrec.Uses > 0 Then UserItemUses(nCurItem) = Itemrec.Uses
                        nCurItem = nCurItem + 1
                    End If
                End If
                'sItems(x) = Itemrec.Number
            Next y
            
            sFoundItems(x, 1) = sText
            sFoundItems(x, 2) = Itemrec.Number
            sItems(x) = ""
            
        End If
    Next x
    
    For x = 0 To nMaxSItems
        If sFoundItems(x, 1) = sText And Not sFoundItems(x, 2) = CStr(Itemrec.Number) Then
            sFoundItems(x, 2) = sFoundItems(x, 2) & "," & Itemrec.Number
            Exit For
        End If
    Next
skip_item:
    nStatus = BTRCALL(BGETNEXT, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
Loop

'check for dups
For x = 0 To nMaxSItems
    nOrigMatch = 0
    If Not InStr(1, sFoundItems(x, 2), ",") = 0 Then
        
        'load match form
        Load frmUserSelectItem
        frmUserSelectItem.Tag = "-1"
        frmUserSelectItem.lstItems.clear
        frmUserSelectItem.FormJump = 0 'items
        
        'find numbers and load list
        sText = ""
        For y = 1 To Len(sFoundItems(x, 2))
            sChar = Mid(sFoundItems(x, 2), y, 1)
            If IsNumeric(sChar) And Not y = Len(sFoundItems(x, 2)) Then
                sText = sText & sChar
            Else
                If y = Len(sFoundItems(x, 2)) Then sText = sText & sChar
                If nOrigMatch = 0 Then nOrigMatch = Val(sText)
                frmUserSelectItem.lstItems.AddItem Val(sText) & vbTab & GetItemName(Val(sText))
                frmUserSelectItem.lstItems.ItemData(frmUserSelectItem.lstItems.NewIndex) = Val(sText)
                sText = ""
            End If
        Next y
        frmUserSelectItem.lstItems.ListIndex = 0
        'frmUserSelectItem.Left = Me.Left + 500
        'frmUserSelectItem.Top = Me.Top + 500
        frmUserSelectItem.Show
        frmUserSelectItem.SetFocus
        
        Do While frmUserSelectItem.Tag = "-1"
            DoEvents
        Loop
        
        If frmUserSelectItem.lstItems.ListIndex < 0 Then GoTo canceled:
        
        'change numbers
        If Not nOrigMatch = frmUserSelectItem.lstItems.ItemData(frmUserSelectItem.lstItems.ListIndex) Then
            For y = 0 To 49
                If UserItem(y) = nOrigMatch Then
                    UserItem(y) = frmUserSelectItem.lstItems.ItemData(frmUserSelectItem.lstItems.ListIndex)
                    UserItemUses(y) = GetItemUses(UserItem(y))
                End If
                If UserKey(y) = nOrigMatch Then
                    UserKey(y) = frmUserSelectItem.lstItems.ItemData(frmUserSelectItem.lstItems.ListIndex)
                    UserKeyUses(y) = GetItemUses(UserKey(y))
                End If
            Next y
            For y = 50 To 99
                If UserItem(y) = nOrigMatch Then
                    UserItem(y) = frmUserSelectItem.lstItems.ItemData(frmUserSelectItem.lstItems.ListIndex)
                    UserItemUses(y) = GetItemUses(UserItem(y))
                End If
            Next y
            For y = 0 To 19
                If Val(txtWornItem(y).Text) = nOrigMatch Then
                    txtWornItem(y).Text = frmUserSelectItem.lstItems.ItemData(frmUserSelectItem.lstItems.ListIndex)
                End If
            Next y
            If Val(txtWeaponNumber.Text) = nOrigMatch Then
                txtWeaponNumber.Text = frmUserSelectItem.lstItems.ItemData(frmUserSelectItem.lstItems.ListIndex)
            End If
        End If
    End If
Next x

Call LoadUserItems(True)
Call CheckWornItems

canceled:
Me.MousePointer = vbDefault
Me.Enabled = True
Unload frmUserSelectItem
Exit Sub
error:
Call HandleError
Resume canceled:
End Sub

Private Sub PasteSpells()
On Error GoTo error:
Dim sSearch As String, sText As String, sChar As String, x As Integer, y As Integer, x2 As Integer
Dim nStatus As Integer, bResult As Boolean, nMaxSSpells As Integer, nCurSpell As Integer
Dim sSpells(0 To 199) As String, nOrigMatch As Long
Dim sFoundSpells(0 To 199, 1 To 2) As String

'x = current position in string
'y = length of next possible (current) string match
'x2 = starting point of field match

Me.Enabled = False
Load frmUserPaste
frmUserPaste.Caption = "Paste Spellbook"
frmUserPaste.txtText = ""
frmUserPaste.Tag = "-1"

sSearch = Clipboard.GetText
If Not sSearch = "" Then frmUserPaste.txtText = sSearch

frmUserPaste.Show vbModal, frmMain
If frmUserPaste.Tag = "-1" Then GoTo canceled:

sSearch = frmUserPaste.txtText.Text

Unload frmUserPaste

If Len(sSearch) < 10 Then GoTo canceled:

Me.MousePointer = vbHourglass

x = 1
y = 1
x2 = -1
Do Until x + y > Len(sSearch) + 1
    
    sChar = Mid(sSearch, x + y - 1, 1)
    
    If Asc(sChar) = 10 Or Asc(sChar) = 13 Then GoTo store_spell:
    'bResult = TestPasteChar(sChar)
    'If bResult = False Then GoTo next_y:
    
    'if IsNumeric(sChar) Then GoTo store_spell:
    
    'sText = RemoveCharacter(sText & sChar, " ")
    sText = sText & sChar
    'If Right(sText, 2) = "  " Then sText = Left(sText, Len(sText) - 1)
    
    If Not InStr(1, LCase(sText), "you have the following spells:") = 0 Then
        GoTo clear:
    ElseIf Not InStr(1, LCase(sText), "level mana short spell name") = 0 Then
        GoTo clear:
    End If
    
    If (x + y + 1) > (Len(sSearch) + 1) Then GoTo store_spell:
    
GoTo next_y:

store_spell:
If Len(sText) > 17 Then
    sSpells(nCurSpell) = Right(sText, Len(sText) - 17)
    If nCurSpell < 199 Then nCurSpell = nCurSpell + 1
End If

clear:
sText = ""
x = x + y
y = 0
x2 = -1

next_y:
    y = y + 1
Loop

'RESET STATS:
x = MsgBox("Clear spellbook before loading?  If you choose not to," & vbCrLf _
        & "then anything found will be /added/ to the current spellbook." _
        , vbYesNoCancel + vbQuestion + vbDefaultButton1, "Clear Spellbook?")
If x = vbCancel Then
    GoTo canceled:
ElseIf x = vbYes Then
    Call cmdClearAllSpell_Click
End If

'DISTRO DATA COLLECTED:
nMaxSSpells = nCurSpell
nCurSpell = 0

nStatus = BTRCALL(BGETFIRST, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error getting first spell -- " & BtrieveErrorCode(nStatus)
    GoTo canceled:
End If
Do While nStatus = 0
    Call SpellRowToStruct(Spelldatabuf.buf)
    
    'sText = RemoveCharacter(ClipNull(Spellrec.Name), " ")
    sText = ClipNull(Spellrec.Name)
    
    If sText = "" Then GoTo skip_spell:
    
    For x = 0 To nMaxSSpells
        If sText = sSpells(x) Then
retry_spell:
            If nCurSpell < 100 Then
                If Not UserSpell(nCurSpell) = 0 Then nCurSpell = nCurSpell + 1: GoTo retry_spell:
                UserSpell(nCurSpell) = Spellrec.Number
                nCurSpell = nCurSpell + 1
                
                sFoundSpells(x, 1) = sText
                sFoundSpells(x, 2) = Spellrec.Number
                sSpells(x) = ""
            End If
        End If
    Next x
    
    For x = 0 To nMaxSSpells
        If sFoundSpells(x, 1) = sText And Not sFoundSpells(x, 2) = CStr(Spellrec.Number) Then
            sFoundSpells(x, 2) = sFoundSpells(x, 2) & "," & Spellrec.Number
            Exit For
        End If
    Next
skip_spell:
    nStatus = BTRCALL(BGETNEXT, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), ByVal SpellKeyBuffer, KEY_BUF_LEN, 0)
Loop

'check for dups
For x = 0 To nMaxSSpells
    nOrigMatch = 0
    If Not InStr(1, sFoundSpells(x, 2), ",") = 0 Then
        
        'load match form
        Load frmUserSelectItem
        frmUserSelectItem.Tag = "-1"
        frmUserSelectItem.lstItems.clear
        frmUserSelectItem.FormJump = 1 'spells
        
        'find numbers and load list
        sText = ""
        For y = 1 To Len(sFoundSpells(x, 2))
            sChar = Mid(sFoundSpells(x, 2), y, 1)
            If IsNumeric(sChar) And Not y = Len(sFoundSpells(x, 2)) Then
                sText = sText & sChar
            Else
                If y = Len(sFoundSpells(x, 2)) Then sText = sText & sChar
                If nOrigMatch = 0 Then nOrigMatch = Val(sText)
                frmUserSelectItem.lstItems.AddItem Val(sText) & vbTab & GetShortSpellName(Val(sText)) & vbTab & GetSpellName(Val(sText))
                frmUserSelectItem.lstItems.ItemData(frmUserSelectItem.lstItems.NewIndex) = Val(sText)
                sText = ""
            End If
        Next y
        frmUserSelectItem.lstItems.ListIndex = 0
        'frmUserSelectItem.Left = Me.Left + 500
        'frmUserSelectItem.Top = Me.Top + 500
        frmUserSelectItem.Show
        frmUserSelectItem.SetFocus
        
        Do While frmUserSelectItem.Tag = "-1"
            DoEvents
        Loop
        
        If frmUserSelectItem.lstItems.ListIndex < 0 Then GoTo canceled:
        
        'change numbers
        If Not nOrigMatch = frmUserSelectItem.lstItems.ItemData(frmUserSelectItem.lstItems.ListIndex) Then
            For y = 0 To 99
                If UserSpell(y) = nOrigMatch Then
                    UserSpell(y) = frmUserSelectItem.lstItems.ItemData(frmUserSelectItem.lstItems.ListIndex)
                End If
            Next
        End If
    End If
Next x

Call LoadUserItems(True)

canceled:
Me.MousePointer = vbDefault
Me.Enabled = True
Unload frmUserSelectItem
Exit Sub
error:
Call HandleError
Resume canceled:
End Sub

Private Sub cmdPasteItems_Click(Index As Integer)
Call PasteItems
End Sub

Private Sub cmdPasteSpells_Click()
Call PasteSpells
End Sub


Private Sub cmdPasteStatQ_Click()
MsgBox "Revive will set HP and Mana to full and then present choices for other stuff." _
    & vbCrLf & vbCrLf & "Paste: Paste a capture of a character's ""stat"" output.  Class, Race, Level, Exp, " _
    & "Lives, CP, HP, Mana, and the six stats will be pasted.  NOTE: You can also setup a character " _
    & "in MMUD Explorer and click ""Copy Only Stats"" and paste that here as well.", vbInformation
End Sub

Private Sub cmdReviveChar_Click()
On Error GoTo error:
Dim x As Integer, y As Long, sTemp As String, MapRoom As RoomExitType

txtCurrentHP.Text = Val(txtMaxHP.Text)
txtCurrentMana.Text = Val(txtMaxMana.Text)

If Val(txtLives.Text) < 9 Then
    x = MsgBox("Reset Lives?", vbYesNo + vbDefaultButton2 + vbQuestion)
    If x = vbYes Then txtLives.Text = 9
End If

y = Val(txtSomeSortOfFlag.Text)
If y = 0 Then
    For x = 0 To 9
        If txtSpellNumber(x).Text > 0 Then y = 1: Exit For
    Next x
End If
If y > 0 Then
    SSTab3.Tab = 6
    DoEvents
    x = MsgBox("Reset Auras?", vbYesNo + vbDefaultButton2 + vbQuestion)
    If x = vbYes Then Call cmdClearSpellsCasted_Click
End If

If Val(txtCurrentMap.Text) <> 1 Or Val(txtCurrentRoom.Text) <> 2189 Then 'halls
    SSTab3.Tab = 4
    DoEvents
    sTemp = InputBox("Move user?" _
        & vbCrLf & "-Enter 1-19: move them that many rooms back" _
        & vbCrLf & "-Enter 99: teleport them to the halls of the dead" _
        & vbCrLf & "-Enter Map/Room in the format of 1/2189" _
        & vbCrLf & vbCrLf & "Enter 0 or cancel to leave them alone.", , 0)
    If InStr(1, sTemp, "/", vbTextCompare) Then
        MapRoom = ExtractMapRoom(sTemp)
        txtCurrentMap.Text = MapRoom.Map
        txtCurrentRoom.Text = MapRoom.Room
    ElseIf Val(sTemp) = 99 Then
        txtCurrentMap.Text = 1
        txtCurrentRoom.Text = 2189
    ElseIf Val(sTemp) > 0 Then
        txtCurrentMap.Text = txtMapTrail(Val(sTemp)).Text
        txtCurrentRoom.Text = txtRoomTrail(Val(sTemp)).Text
    End If
End If

SSTab3.Tab = 0
cmdReviveChar.SetFocus
DoEvents

MsgBox "Done. You still need to save.", vbInformation

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("cmdReviveChar_Click")
Resume out:

End Sub

Private Sub cmdSomeFlagQ_Click()
MsgBox "Observed a value of 78 here that indicated the the character was poisoned.  It caused the character to take damage and not be able to rest.  It could be a bitmask or something.", vbInformation
End Sub

Private Sub cmdSpellEditor_GotFocus()
'Call SelectAll(cmdSpellEditor)

End Sub

Private Sub cmdUserExport_Click()
On Error GoTo error:
Dim x As Integer, sTemp(1 To 3) As String, nStatus As Integer
Dim sExportText As String

nStatus = BTRCALL(BGETEQUAL, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal sCurrentRecord, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error on save, BGETQUAL: " & BtrieveErrorCode(nStatus)
    Exit Sub
Else
    UserRowToStruct Userdatabuf.buf
End If

sExportText = AutoAppendString(sExportText, "BBSName:" & ClipNull(Userrec.BBSName), vbCrLf)
sExportText = AutoAppendString(sExportText, "FirstName:" & ClipNull(Userrec.FirstName), vbCrLf)
sExportText = AutoAppendString(sExportText, "LastName:" & ClipNull(Userrec.LastName), vbCrLf)
sExportText = AutoAppendString(sExportText, "Race:" & Userrec.Race, vbCrLf)
sExportText = AutoAppendString(sExportText, "Class:" & Userrec.Class, vbCrLf)
sExportText = AutoAppendString(sExportText, "Experience:" & (SLong2ULong(Userrec.BillionsOfExperience) * 1000000000#) + SLong2ULong(Userrec.MillionsOfExperience), vbCrLf)
sExportText = AutoAppendString(sExportText, "Level:" & SInt2UInt(Userrec.Level), vbCrLf)
sExportText = AutoAppendString(sExportText, "Title:" & ClipNull(Userrec.Title), vbCrLf)
sExportText = AutoAppendString(sExportText, "EvilPoints:" & Userrec.EvilPoints, vbCrLf)
sExportText = AutoAppendString(sExportText, "GangName:" & ClipNull(Userrec.GangName), vbCrLf)
sExportText = AutoAppendString(sExportText, "SuicidePassword:" & ClipNull(Userrec.SuicidePassword), vbCrLf)
sExportText = AutoAppendString(sExportText, "BroadcastChan:" & Userrec.BroadcastChan, vbCrLf)
sExportText = AutoAppendString(sExportText, "bEDITED:" & Userrec.bEDITED, vbCrLf)
sExportText = AutoAppendString(sExportText, "Runic:" & SLong2ULong(Userrec.Runic), vbCrLf)
sExportText = AutoAppendString(sExportText, "Platinum:" & SLong2ULong(Userrec.Platinum), vbCrLf)
sExportText = AutoAppendString(sExportText, "Gold:" & SLong2ULong(Userrec.Gold), vbCrLf)
sExportText = AutoAppendString(sExportText, "Silver:" & SLong2ULong(Userrec.Silver), vbCrLf)
sExportText = AutoAppendString(sExportText, "Copper:" & SLong2ULong(Userrec.Copper), vbCrLf)

sExportText = AutoAppendString(sExportText, "MaxHP:" & Userrec.MaxHP, vbCrLf)
sExportText = AutoAppendString(sExportText, "CurrentHP:" & Userrec.CurrentHP, vbCrLf)
sExportText = AutoAppendString(sExportText, "HPRolls:" & Userrec.HitPointRolls, vbCrLf)
sExportText = AutoAppendString(sExportText, "MaxMana:" & Userrec.MaxMana, vbCrLf)
sExportText = AutoAppendString(sExportText, "CurrentMana:" & Userrec.CurrentMana, vbCrLf)
sExportText = AutoAppendString(sExportText, "SpellCasting:" & Userrec.SpellCasting, vbCrLf)
sExportText = AutoAppendString(sExportText, "LivesRemaining:" & Userrec.LivesRemaining, vbCrLf)
sExportText = AutoAppendString(sExportText, "CPRemaining:" & Userrec.CPRemaining, vbCrLf)
sExportText = AutoAppendString(sExportText, "Perception:" & Userrec.Perception, vbCrLf)
sExportText = AutoAppendString(sExportText, "Thievery:" & Userrec.Thievery, vbCrLf)
sExportText = AutoAppendString(sExportText, "Traps:" & Userrec.Traps, vbCrLf)
sExportText = AutoAppendString(sExportText, "Picklocks:" & Userrec.Picklocks, vbCrLf)
sExportText = AutoAppendString(sExportText, "Tracking:" & Userrec.Tracking, vbCrLf)
sExportText = AutoAppendString(sExportText, "MartialArts:" & Userrec.MartialArts, vbCrLf)
sExportText = AutoAppendString(sExportText, "MagicRes:" & Userrec.MagicRes, vbCrLf)
sExportText = AutoAppendString(sExportText, "Stealth:" & Userrec.Stealth, vbCrLf)

sTemp(1) = ""
For x = 0 To 11
    sTemp(1) = AutoAppendString(sTemp(1), Userrec.Stat(x))
Next
sExportText = AutoAppendString(sExportText, "STATS:" & sTemp(1), vbCrLf)

sTemp(1) = ""
For x = 0 To 29
    sTemp(1) = AutoAppendString(sTemp(1), Userrec.Ability(x) & "|" & Userrec.AbilityModifier(x))
Next x
sExportText = AutoAppendString(sExportText, "ABILS:" & sTemp(1), vbCrLf)


sExportText = AutoAppendString(sExportText, "CurrentENC:" & SInt2UInt(Userrec.CurrentENC), vbCrLf)
sExportText = AutoAppendString(sExportText, "MaxENC:" & SInt2UInt(Userrec.MaxENC), vbCrLf)
sExportText = AutoAppendString(sExportText, "WeaponHand:" & Userrec.WeaponHand, vbCrLf)

sTemp(1) = ""
For x = 0 To 19
    sTemp(1) = AutoAppendString(sTemp(1), Userrec.WornItem(x))
Next x
sExportText = AutoAppendString(sExportText, "WORN:" & sTemp(1), vbCrLf)

sTemp(1) = ""
sTemp(2) = ""
sTemp(3) = ""
For x = 0 To 99
    sTemp(1) = AutoAppendString(sTemp(1), Userrec.Item(x) & "|" & Userrec.ItemUses(x))
    If x < 50 Then sTemp(2) = AutoAppendString(sTemp(2), Userrec.Key(x) & "|" & Userrec.KeyUses(x))
    sTemp(3) = AutoAppendString(sTemp(3), Userrec.Spell(x))
Next x

sExportText = AutoAppendString(sExportText, "ITEMS:" & sTemp(1), vbCrLf)
sExportText = AutoAppendString(sExportText, "KEYS:" & sTemp(2), vbCrLf)
sExportText = AutoAppendString(sExportText, "SPELLS:" & sTemp(3), vbCrLf)

sTemp(1) = ""
For x = 0 To 9
    sTemp(1) = AutoAppendString(sTemp(1), SInt2UInt(Userrec.SpellCasted(x)) & "|" & SInt2UInt(Userrec.SpellValue(x)) & "|" & SInt2UInt(Userrec.SpellRoundsLeft(x)))
Next x
sExportText = AutoAppendString(sExportText, "AURAS:" & sTemp(1), vbCrLf)

sExportText = AutoAppendString(sExportText, "RoomNum:" & Userrec.RoomNum, vbCrLf)
sExportText = AutoAppendString(sExportText, "MapNumber:" & Userrec.MapNumber, vbCrLf)
sTemp(1) = ""
For x = 0 To 19
    sTemp(1) = AutoAppendString(sTemp(1), Userrec.LastMap(x) & "|" & Userrec.LastRoom(x))
Next x
sExportText = AutoAppendString(sExportText, "ROOMS:" & sTemp(1), vbCrLf)

Clipboard.clear
Clipboard.SetText sExportText

MsgBox "Data copied to clipboard.", vbOKOnly + vbInformation, "Export User"

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("cmdUserExport_Click")
Resume out:
End Sub

Private Sub cmdUserHitPointMax_Click()
On Error GoTo error:

'Max = [(Class Max Range minus Class Min Range) x Level]
txtHitPointRolls.Text = GetClassMaxHP(cmbClasses.ItemData(cmbClasses.ListIndex)) * Val(txtLevel.Text)

If Val(txtHitPointRolls.Text) > 255 Then txtHitPointRolls.Text = 255

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("cmdUserHitPointMax_Click")
Resume out:

End Sub

Private Sub cmdUserHitPointRollQ_Click()
MsgBox "This is the cummulative value of the random rolls of additional hitpoints when leveling, added to hitpoint calculation for maxmimum HP. " _
    & "MMUD makes sure this value is in range of what's allowed. Max = [(Class Max Range minus Class Min Range) x Level].  Max value = 255." _
    & vbCrLf & vbCrLf & "Note that a value of 255 will rollover upon levling and reset back to the minimum.", vbInformation + vbOKOnly
End Sub

Private Sub cmdUserImport_Click()
txtSearch.Enabled = False
lvDatabase.Enabled = False
framImport.Visible = True
End Sub

Private Sub ImportFromClipboard()
On Error GoTo error:
Dim x As Integer, y As Integer, sSubMatches() As String, sSubValues() As String
Dim sImportText As String, iMatch As Integer, iSubMatch As Integer, nValue As Long
Dim tMatches() As RegexMatches, sRegexPattern As String
Dim bWriteName As Integer, bWriteGang As Integer, bImportRooms As Boolean
Dim bImportDeets As Boolean

sImportText = Trim(Clipboard.GetText)
If Len(sImportText) = 0 Then Exit Sub

sRegexPattern = "^([^\r\n:]+):([^\r\n]+|)$"
tMatches() = RegExpFindv2(sImportText, sRegexPattern, False, True, True)
If UBound(tMatches()) = 0 And Len(tMatches(0).sFullMatch) = 0 Then Exit Sub

If chkImportRooms.Value = 1 Then bImportRooms = True
If chkImportDeets.Value = 1 Then bImportDeets = True
If chkImportName.Value = 1 Then bWriteName = True
If chkImportName.Value = 1 Then bWriteGang = True
If chkImportClearItems.Value = 1 Then
    Call cmdClearAllItem_Click
    Call cmdClearWorn_Click
End If
If chkImportClearKeys.Value = 1 Then Call cmdClearAllKey_Click
If chkImportClearSpells.Value = 1 Then
    Call cmdClearAllSpell_Click
    Call cmdClearSpellsCasted_Click
End If
If chkImportClearAbils.Value = 1 Then Call cmdAbilsClear_Click

For iMatch = 0 To UBound(tMatches())
    If UBound(tMatches(iMatch).sSubMatches()) = 0 Then GoTo skip_match
    
    Select Case tMatches(iMatch).sSubMatches(0)
        Case "FirstName": If bWriteName Then txtFirstName.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "LastName": If bWriteName Then txtLastName.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "Experience": txtExperience.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "Level": txtLevel.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "Title": txtTitle.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "GangName": If bWriteGang Then txtGang.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "SuicidePassword": If bImportDeets Then txtSuicide.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "EvilPoints": txtEvilPoints.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "BroadcastChan": If bImportDeets Then txtBroadcastChan.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "MaxHP": txtMaxHP.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "CurrentHP": txtCurrentHP.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "MaxMana": txtMaxMana.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "HPRolls": txtHitPointRolls.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "CurrentMana": txtCurrentMana.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "SpellCasting": txtSpellcasting.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "LivesRemaining": If bImportDeets Then txtLives.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "CPRemaining": txtCp.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "Perception": txtPerception.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "Thievery": txtThievery.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "Traps": txtTraps.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "Picklocks": txtPicklocks.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "Tracking": txtTracking.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "MartialArts": txtMartialArts.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "MagicRes": txtMagicResistance.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "Stealth": txtStealth.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "WeaponHand": txtWeaponNumber.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "RoomNum": If bImportRooms Then txtCurrentRoom.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "MapNumber": If bImportRooms Then txtCurrentMap.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "CurrentENC": txtCurrentEncum.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "MaxENC": txtMaxEncum.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "Runic": If bImportDeets Then txtRunic.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "Platinum": If bImportDeets Then txtPlatinum.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "Gold": If bImportDeets Then txtGold.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "Silver": If bImportDeets Then txtSilver.Text = Trim(tMatches(iMatch).sSubMatches(1))
        Case "Copper": If bImportDeets Then txtCopper.Text = Trim(tMatches(iMatch).sSubMatches(1))
        
        Case "bEDITED":
            nValue = Val(Trim(tMatches(iMatch).sSubMatches(1)))
            If nValue > 0 Then
                chkEdited.Value = 1
            Else
                chkEdited.Value = 0
            End If
            
        Case "Race":
            nValue = Val(Trim(tMatches(iMatch).sSubMatches(1)))
            If nValue > cmbRaces.ListCount - 1 Then
                MsgBox "Race not found.", vbOKOnly + vbInformation
                Call Add2RaceArray(nValue)
                cmbRaces.clear
                For x = 0 To UBound(Races)
                    cmbRaces.AddItem Races(x).Name
                Next x
            End If
            cmbRaces.ListIndex = nValue
            
        Case "Class":
            nValue = Val(Trim(tMatches(iMatch).sSubMatches(1)))
            If nValue > cmbClasses.ListCount - 1 Then
                MsgBox "Class not found.", vbOKOnly + vbInformation
                Call Add2ClassArray(nValue)
                cmbClasses.clear
                For x = 0 To UBound(Classes)
                    cmbClasses.AddItem Classes(x).Name
                Next x
            End If
            cmbClasses.ListIndex = nValue
            
        Case "STATS", "WORN", "SPELLS":
            sSubMatches() = Split(Trim(tMatches(iMatch).sSubMatches(1)), ",")
            For x = 0 To UBound(sSubMatches())
                Select Case tMatches(iMatch).sSubMatches(0)
                    Case "STATS": txtStat(x).Text = sSubMatches(x)
                    Case "WORN": txtWornItem(x).Text = sSubMatches(x)
                    Case "SPELLS": UserSpell(x) = Val(sSubMatches(x))
                End Select
            Next x
            
        Case "ABILS", "ITEMS", "KEYS", "ROOMS":
            sSubMatches() = Split(Trim(tMatches(iMatch).sSubMatches(1)), ",")
            For x = 0 To UBound(sSubMatches())
                sSubValues() = Split(sSubMatches(x), "|")
                If UBound(sSubValues()) = 1 Then
                    Select Case tMatches(iMatch).sSubMatches(0)
                        Case "ABILS":
                            txtAbilityA(x).Text = sSubValues(0)
                            txtAbilityB(x).Text = sSubValues(1)
                        Case "ITEMS":
                            UserItem(x) = sSubValues(0)
                            UserItemUses(x) = sSubValues(1)
                         Case "KEYS":
                            UserKey(x) = sSubValues(0)
                            UserKeyUses(x) = sSubValues(1)
                        Case "ROOMS":
                            If bImportRooms Then
                                txtMapTrail(x).Text = sSubValues(0)
                                txtRoomTrail(x).Text = sSubValues(1)
                            End If
                    End Select
                End If
            Next x
        
        Case "AURAS":
            sSubMatches() = Split(Trim(tMatches(iMatch).sSubMatches(1)), ",")
            For x = 0 To UBound(sSubMatches())
                sSubValues() = Split(sSubMatches(x), "|")
                If UBound(sSubValues()) = 2 Then
                    txtSpellNumber(x).Text = sSubValues(0)
                    txtSpellValue(x).Text = sSubValues(1)
                    txtSpellRounds(x).Text = sSubValues(2)
                End If
            Next x
        
    End Select
    
skip_match:
Next iMatch

Call LoadUserItems(True)
Call CheckWornItems

MsgBox "Done. Save if happy, discard if not.", vbOKOnly + vbInformation, "Import User"

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("ImportFromClipboard")
Resume out:
End Sub

Private Sub cmdUserUnknownsNote_Click()
MsgBox "Three things..." _
    & vbCrLf & vbCrLf & "#1) These values are not saved." _
    & vbCrLf & vbCrLf & "#2) Clicking a different set will reload the selected character and lose any changes to the other tabs if auto-save is off." _
    & vbCrLf & vbCrLf & "#3) The values themselves could be incorrect as assumptions have been made to whether values are 1-byte, 2-byte, or 4-byte values. " _
    & "i.e. FFFF in the DB could be read as two 1-byte values of 255(FF) or one 2-byte value of 65535(FFFF)." _
    , vbInformation + vbOKOnly
End Sub

Private Sub Form_Load()
On Error GoTo error:
Dim nStatus As Integer
bLoaded = False

With EL1
    .FormInQuestion = Me
    .MINHEIGHT = 385 + (TITLEBAR_OFFSET / 10) + 12
    .MINWIDTH = 690
    .CenterOnLoad = False
    .EnableLimiter = True
End With

Me.Top = ReadINI("Windows", "UserTop")
Me.Left = ReadINI("Windows", "UserLeft")
Me.Width = ReadINI("Windows", "UserWidth")
Me.Height = ReadINI("Windows", "UserHeight")
chkNoRoomNames.Value = ReadINI("Settings", "UserNoLookupRooms")

cmbRaces.clear
Dim i&
For i = 0 To UBound(Races)
    cmbRaces.AddItem Races(i).Name
    cmbRaces.ItemData(cmbRaces.NewIndex) = i
Next i

cmbClasses.clear
For i = 0 To UBound(Classes)
    cmbClasses.AddItem Classes(i).Name
    cmbClasses.ItemData(cmbClasses.NewIndex) = i
Next i

Call LoadUsers

Me.Show
Me.SetFocus
txtSearch.SetFocus
If ReadINI("Windows", "UserMaxed") = "1" Then Me.WindowState = vbMaximized

Exit Sub
error:
Call HandleError
Resume Next
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadUsers()
On Error GoTo error:
Dim oLI As ListItem, x As Integer
Dim nStatus As Integer

lvDatabase.ColumnHeaders.clear
lvDatabase.ColumnHeaders.add 1, "MudName", "Mud Name", 1400, lvwColumnLeft
lvDatabase.ColumnHeaders.add 2, "BBSID", "BBS ID", 1400, lvwColumnLeft
If Not bOnlyNames Then
    lvDatabase.ColumnHeaders.add 3, "Class", "Class", 1100, lvwColumnLeft
    lvDatabase.ColumnHeaders.add 4, "Race", "Race", 1100, lvwColumnLeft
    lvDatabase.ColumnHeaders.add 5, "EXP", "EXP", 1300, lvwColumnLeft
    lvDatabase.ColumnHeaders.add 6, "LVL", "LVL", 600, lvwColumnLeft
    lvDatabase.ColumnHeaders.add 7, "Lives", "Lives", 600, lvwColumnLeft
    lvDatabase.ColumnHeaders.add 8, "CP", "CP", 600, lvwColumnLeft
End If

lvDatabase.ListItems.clear

nStatus = BTRCALL(BGETFIRST, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "LoadUser, BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

DoEvents
Do While nStatus = 0
    UserRowToStruct Userdatabuf.buf
    
    Call AddUser2LV(lvDatabase)
    
    nStatus = BTRCALL(BGETNEXT, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0)
Loop

'If chkJumpLast.value = 1 Then
'    Set lvDatabase.SelectedItem = lvDatabase.ListItems(lvDatabase.ListItems.Count)
'    lvDatabase.SelectedItem.EnsureVisible
'    Call lvDatabase_ItemClick(lvDatabase.SelectedItem)
'Else
    Set lvDatabase.SelectedItem = lvDatabase.ListItems(1)
    lvDatabase.SelectedItem.EnsureVisible
    Call lvDatabase_ItemClick(lvDatabase.SelectedItem)
'End If

bLoaded = True

SortListView lvDatabase, lvDatabase.ColumnHeaders(1).Index, ldtString, True

lvDatabase.refresh

Set oLI = Nothing

Exit Sub
error:
Call HandleError("LoadUsers")
Set oLI = Nothing

End Sub
Private Sub AddUser2LV(lv As ListView)
Dim oLI As ListItem
    
On Error GoTo error:

    Set oLI = lv.ListItems.add()
    oLI.Text = ClipNull(Userrec.FirstName, Len(Userrec.FirstName))
    
    oLI.ListSubItems.add (1), "BBSID", ClipNull(Userrec.BBSName, Len(Userrec.BBSName))
    If Not bOnlyNames Then
        oLI.ListSubItems.add (2), "Class", GetClassName(Userrec.Class)
        oLI.ListSubItems.add (3), "Race", GetRaceName(Userrec.Race)
        oLI.ListSubItems.add (4), "EXP", (SLong2ULong(Userrec.BillionsOfExperience) * 1000000000#) + SLong2ULong(Userrec.MillionsOfExperience)
        oLI.ListSubItems.add (5), "LVL", Userrec.Level
        oLI.ListSubItems.add (6), "Lives", Userrec.LivesRemaining
        oLI.ListSubItems.add (7), "CP", Userrec.CPRemaining
    End If
    
Set oLI = Nothing

Exit Sub
error:
Call HandleError("AddUser2LV")
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClearAllItem_Click()
Dim x As Integer

On Error GoTo error:

For x = 0 To 99
    UserItem(x) = 0
    UserItemUses(x) = -1
Next

Call LoadUserItems(True)
Call CheckWornItems

Exit Sub
error:
Call HandleError
End Sub
Private Sub cmdClearAllKey_Click()
Dim x As Integer

On Error GoTo error:

For x = 0 To 49
    UserKey(x) = 0
    UserKeyUses(x) = -2
Next

Call LoadUserItems(True)

Exit Sub
error:
Call HandleError
End Sub

Private Sub cmdClearAllSpell_Click()
Dim x As Integer

On Error GoTo error:

For x = 0 To 99
    UserSpell(x) = 0
Next

Call LoadUserItems(True)

Exit Sub
error:
Call HandleError
End Sub

Private Sub cmdClearSpellsCasted_Click()
Dim x As Integer

On Error GoTo error:

For x = 0 To 9
    txtSpellNumber(x).Text = 0
    txtSpellName(x).Text = "none"
    txtSpellValue(x).Text = 0
    txtSpellRounds(x).Text = 0
Next

txtSomeSortOfFlag.Text = 0

Exit Sub
error:
Call HandleError
End Sub

Private Sub cmdClearWorn_Click()
Dim x As Integer

On Error GoTo error:

txtWeaponNumber.Text = 0

For x = 0 To 19
    txtWornItem(x).Text = 0
Next


Exit Sub
error:
Call HandleError

End Sub

Private Sub cmdCopy_Click()
On Error GoTo error:
Dim x As Integer, nStatus As Integer, BBSName As String, FirstName As String, temp As String

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

temp = MsgBox("Warning, this will have VERY unpredictable results" & vbCrLf _
& "with limited items, gangs, etc and should really only" & vbCrLf _
& "be used to transfer characters between BBS accounts (if that is even safe)." & vbCrLf _
& vbCrLf & "Are you sure you want to continue?", vbYesNo, "Copy User")

If temp <> 6 Then Exit Sub

BBSName = InputBox("Enter *BBS Name* to copy to (New or Existing)" & vbCrLf & vbCrLf & "NOTE: Enter it EXACTLY, Case Sensitive!", "Copy user", "")
If BBSName = "" Then Exit Sub

If bLoaded = True And chkAutoSave.Value = 1 Then saverecord (sCurrentRecord)

BBSName = Trim(BBSName) & String(30 - Len(BBSName), Chr(0))

nStatus = BTRCALL(BGETEQUAL, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal BBSName, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    If nStatus = 4 Then ''' record doesn't exist
        If bLoaded = True Then
            nStatus = BTRCALL(BGETEQUAL, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal sCurrentRecord, KEY_BUF_LEN, 0)
            UserRowToStruct Userdatabuf.buf
            FormValuesToRecord
        Else
            For x = 1 To UBound(Userdatabuf.buf)
                Userdatabuf.buf(x) = &H0
            Next x
            UserRowToStruct Userdatabuf.buf
        End If
        
        Userrec.BBSName = BBSName
        
        FirstName = InputBox("Enter New First Name (must be unique)", "Enter New First Name")
        If FirstName = "" Then
            Form_Load
            Exit Sub
        End If
        
        Userrec.FirstName = RTrim(RemoveCharacter(FirstName, " ") & Chr(0))
        UserStructToRow Userdatabuf.buf
        
        nStatus = BTRCALL(BINSERT, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            MsgBox "Insert error: " & BtrieveErrorCode(nStatus)
            Form_Load
            Exit Sub
        End If
    Else
        MsgBox "BGETEQUAL Error: " & BtrieveErrorCode(nStatus)
        Form_Load
        Exit Sub
    End If
Else
    UserRowToStruct Userdatabuf.buf
    temp = Userrec.FirstName
    FormValuesToRecord
    Userrec.FirstName = LTrim(RTrim(RemoveCharacter(temp, " ")) & Chr(0))
    Userrec.BBSName = BBSName
    nStatus = UpdateUser
    If Not nStatus = 0 Then
        MsgBox "Record was found but BUPDATE had an error: " & BtrieveErrorCode(nStatus)
        Form_Load
        Exit Sub
    End If
End If

Form_Load

Exit Sub
error:
Call HandleError
End Sub

Private Sub cmdDiscard_Click()
Dim nStatus As Integer, BBSName As String

On Error GoTo error:

If lvDatabase.ListItems.Count = 0 Then Exit Sub

BBSName = lvDatabase.SelectedItem.SubItems(1)
BBSName = BBSName & String(30 - Len(BBSName), vbNullChar)

nStatus = BTRCALL(BGETEQUAL, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal BBSName, KEY_BUF_LEN, 0)

If Not nStatus = 0 Then
    MsgBox "Error on BGETGE: " & BtrieveErrorCode(nStatus)
Else
    sCurrentRecord = BBSName
    DispUserInfo Userdatabuf.buf
End If

Exit Sub
error:
Call HandleError
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = vbMinimized Then Exit Sub
framNav.Left = Me.Width - framNav.Width - 200
framImport.Left = Me.Width - framImport.Width - 250
lvDatabase.Width = framNav.Left - 175
lvDatabase.Height = Me.Height - 925 - TITLEBAR_OFFSET - 200
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

If bLoaded = True And chkAutoSave.Value = 1 Then Call saverecord(sCurrentRecord)

Call WriteINI("Settings", "UserNoLookupRooms", chkNoRoomNames.Value)
If Me.WindowState = vbMinimized Then Exit Sub

If Me.WindowState = vbMaximized Then
    Call WriteINI("Windows", "UserMaxed", 1)
Else
    Call WriteINI("Windows", "UserMaxed", 0)
    Call WriteINI("Windows", "UserTop", Me.Top)
    Call WriteINI("Windows", "UserLeft", Me.Left)
    Call WriteINI("Windows", "UserWidth", Me.Width)
    Call WriteINI("Windows", "UserHeight", Me.Height)
End If

End Sub

Private Sub lstItems_DblClick()
On Error GoTo error:

cmdEditItem_Click

Exit Sub
error:
Call HandleError
End Sub

Private Sub lstKeys_DblClick()
On Error GoTo error:

cmdEditKey_Click

Exit Sub
error:
Call HandleError
End Sub

Private Sub lstSpells_DblClick()
On Error GoTo error:

cmdEditSpell_Click

Exit Sub
error:
Call HandleError
End Sub

Private Sub lvDatabase_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo error:

Dim nSort As ListDataType
Select Case ColumnHeader.Index
    Case 5 To 8: nSort = ldtNumber
    Case Else: nSort = ldtString
End Select
SortListView lvDatabase, ColumnHeader.Index, nSort, lvDatabase.SortOrder

Exit Sub
error:
Call HandleError
End Sub



Private Sub optUserUnknowns_Click(Index As Integer)
Call lvDatabase_ItemClick(lvDatabase.SelectedItem)
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

Public Sub lvDatabase_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim nStatus As Integer, BBSName As String

On Error GoTo error:

If bLoaded = True And chkAutoSave.Value = 1 Then saverecord (sCurrentRecord)

BBSName = Item.SubItems(1)
BBSName = BBSName & String(30 - Len(BBSName), vbNullChar)

nStatus = BTRCALL(BGETEQUAL, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal BBSName, KEY_BUF_LEN, 0)

If Not nStatus = 0 Then
    MsgBox "Error on BGETGEQUAL: " & BtrieveErrorCode(nStatus)
    bLoaded = False
Else
    sCurrentRecord = BBSName
    DispUserInfo Userdatabuf.buf
    bLoaded = True
End If

Exit Sub
error:
Call HandleError
End Sub

Private Sub DispUserInfo(row() As Byte)
On Error GoTo error:
Dim x As Integer, counter As Integer, temp As String

Me.MousePointer = 11

UserRowToStruct row()

'bDontSetup = True
'txtUserUnknowns = Userrec.CharLife 'And &H80
''Text2 = Userrec.TestFlag2 'And &H80
''Text3 = Userrec.TestFlag3 'And &H80
'
'Text4 = Userrec.Bitmask1 'And &H80
'Text5 = Userrec.Bitmask2
'
'Text6 = Userrec.unknown12(13)
'
'List1.Clear
'List2.Clear
'For x = 0 To 8
'    List1.AddItem Userrec.unknown13(x)
'Next
'For x = 0 To 3
'    List2.AddItem Userrec.unknown15(x)
'Next

Me.Caption = "User Editor -- " & ClipNull(Userrec.FirstName) & " " & ClipNull(Userrec.LastName)

txtBBSName.Text = Userrec.BBSName
txtFirstName.Text = Userrec.FirstName
txtLastName.Text = Userrec.LastName

If Userrec.Race > cmbRaces.ListCount - 1 Then
    Call Add2RaceArray(Userrec.Race)
    cmbRaces.clear
    For x = 0 To UBound(Races)
        cmbRaces.AddItem Races(x).Name
    Next x
End If

If Userrec.Class > cmbClasses.ListCount - 1 Then
    Call Add2ClassArray(Userrec.Class)
    cmbClasses.clear
    For x = 0 To UBound(Classes)
        cmbClasses.AddItem Classes(x).Name
    Next x
End If

cmbRaces.ListIndex = Userrec.Race
cmbClasses.ListIndex = Userrec.Class
txtLevel.Text = SInt2UInt(Userrec.Level)

If Userrec.bEDITED > 0 Then
    chkEdited.Value = 1
Else
    chkEdited.Value = 0
End If

txtExperience.Text = (SLong2ULong(Userrec.BillionsOfExperience) * 1000000000#) + SLong2ULong(Userrec.MillionsOfExperience)

txtMaxHP.Text = Userrec.MaxHP
txtCurrentHP.Text = Userrec.CurrentHP
txtMaxMana.Text = Userrec.MaxMana
txtCurrentMana.Text = Userrec.CurrentMana
txtSpellcasting.Text = Userrec.SpellCasting
txtLives.Text = Userrec.LivesRemaining
txtCp.Text = Userrec.CPRemaining
txtPerception.Text = Userrec.Perception
txtStealth.Text = Userrec.Stealth
txtThievery.Text = Userrec.Thievery
txtTraps.Text = Userrec.Traps
txtPicklocks.Text = Userrec.Picklocks
txtTracking.Text = Userrec.Tracking
txtMartialArts.Text = Userrec.MartialArts
txtMagicResistance.Text = Userrec.MagicRes
txtBroadcastChan = Userrec.BroadcastChan
txtRunic.Text = SLong2ULong(Userrec.Runic)
txtPlatinum.Text = SLong2ULong(Userrec.Platinum)
txtGold.Text = SLong2ULong(Userrec.Gold)
txtSilver.Text = SLong2ULong(Userrec.Silver)
txtCopper.Text = SLong2ULong(Userrec.Copper)
txtCurrentEncum.Text = SInt2UInt(Userrec.CurrentENC)
txtMaxEncum.Text = SInt2UInt(Userrec.MaxENC)
txtEvilPoints.Text = Userrec.EvilPoints
txtGang.Text = Userrec.GangName
txtSuicide.Text = Userrec.SuicidePassword
txtTitle.Text = Userrec.Title
txtCurrentRoom.Text = Userrec.RoomNum
txtCurrentMap.Text = Userrec.MapNumber
txtWeaponNumber.Text = Userrec.WeaponHand
'txtWeaponName.Text = GetItemName(Userrec.WeaponHand)

txtHitPointRolls.Text = Userrec.HitPointRolls

For x = 0 To 11
    txtStat(x).Text = Userrec.Stat(x)
Next

For x = 0 To 29
    txtAbilityA(x).Text = Userrec.Ability(x)
    txtAbilityB(x).Text = Userrec.AbilityModifier(x)
Next x

For x = 0 To 9
    txtSpellNumber(x).Text = SInt2UInt(Userrec.SpellCasted(x))
    txtSpellValue(x).Text = SInt2UInt(Userrec.SpellValue(x))
    txtSpellRounds(x).Text = SInt2UInt(Userrec.SpellRoundsLeft(x))
Next

txtSomeSortOfFlag.Text = Userrec.SomeSortOfFlag

For x = 0 To 19
    txtMapTrail(x).Text = Userrec.LastMap(x)
    txtRoomTrail(x).Text = Userrec.LastRoom(x)
    txtWornItem(x).Text = Userrec.WornItem(x)
Next

'49 total (base 0 = 48 max index)
'x = 0: lblUserUnknowns(x).Caption = "AAAAA": txtUserUnknowns(x).Text = Userrec.AAAAA

If optUserUnknowns(0).Value = True Then
    x = 0:  lblUserUnknowns(x).Caption = "unkwn1": txtUserUnknowns(x).Text = SInt2UInt(Userrec.unknown1)
    x = 1:  lblUserUnknowns(x).Caption = "Bitmask1": txtUserUnknowns(x).Text = Userrec.Bitmask1
    x = 2:  lblUserUnknowns(x).Caption = "unkwn2(0)": txtUserUnknowns(x).Text = SInt2UInt(Userrec.unknown2(0))
    x = 3:  lblUserUnknowns(x).Caption = "unkwn2(1)": txtUserUnknowns(x).Text = SInt2UInt(Userrec.unknown2(1))
    x = 4:  lblUserUnknowns(x).Caption = "unkwn3(0)": txtUserUnknowns(x).Text = Userrec.unknown3(0)
    x = 5:  lblUserUnknowns(x).Caption = "unkwn3(1)": txtUserUnknowns(x).Text = Userrec.unknown3(1)
    x = 6:  lblUserUnknowns(x).Caption = "unkwn4(0)": txtUserUnknowns(x).Text = SLong2ULong(Userrec.unknown4(0))
    x = 7:  lblUserUnknowns(x).Caption = "unkwn4(1)": txtUserUnknowns(x).Text = SLong2ULong(Userrec.unknown4(1))
    x = 8:  lblUserUnknowns(x).Caption = "unkwn4(2)": txtUserUnknowns(x).Text = SLong2ULong(Userrec.unknown4(2))
    x = 9:  lblUserUnknowns(x).Caption = "unkwn4(3)": txtUserUnknowns(x).Text = SLong2ULong(Userrec.unknown4(3))
    x = 10: lblUserUnknowns(x).Caption = "unkwn5": txtUserUnknowns(x).Text = SLong2ULong(Userrec.unknown5)
    x = 11: lblUserUnknowns(x).Caption = "unkwn6": txtUserUnknowns(x).Text = SInt2UInt(Userrec.unknown6)
    x = 12: lblUserUnknowns(x).Caption = "unkwn8": txtUserUnknowns(x).Text = SInt2UInt(Userrec.unknown8)
    x = 13: lblUserUnknowns(x).Caption = "unkwn12c": txtUserUnknowns(x).Text = Userrec.unknown12c
    x = 14: lblUserUnknowns(x).Caption = "unkwn13a": txtUserUnknowns(x).Text = SInt2UInt(Userrec.unknown13a)
    x = 15: lblUserUnknowns(x).Caption = "A58-Crits": txtUserUnknowns(x).Text = SInt2UInt(Userrec.unknown13b)
    x = 16: lblUserUnknowns(x).Caption = "unkwn13c": txtUserUnknowns(x).Text = SInt2UInt(Userrec.unknown13c)
    x = 17: lblUserUnknowns(x).Caption = "unkwn13d": txtUserUnknowns(x).Text = SInt2UInt(Userrec.unknown13d)
    x = 18: lblUserUnknowns(x).Caption = "unkwn13e": txtUserUnknowns(x).Text = SInt2UInt(Userrec.unknown13e)
    x = 19: lblUserUnknowns(x).Caption = "A10-ACBlur": txtUserUnknowns(x).Text = SInt2UInt(Userrec.unknown13f)
    x = 20: lblUserUnknowns(x).Caption = "unkwn13g": txtUserUnknowns(x).Text = SInt2UInt(Userrec.unknown13g)
    x = 21: lblUserUnknowns(x).Caption = "unkwn14": txtUserUnknowns(x).Text = SInt2UInt(Userrec.unknown14)
    x = 22: lblUserUnknowns(x).Caption = "nothn2": txtUserUnknowns(x).Text = SInt2UInt(Userrec.nothing2)
    x = 23: lblUserUnknowns(x).Caption = "nothn3": txtUserUnknowns(x).Text = SInt2UInt(Userrec.nothing3)
    x = 24: lblUserUnknowns(x).Caption = "nothn4": txtUserUnknowns(x).Text = SInt2UInt(Userrec.nothing4)
    x = 25: lblUserUnknowns(x).Caption = "nothn5": txtUserUnknowns(x).Text = Userrec.nothing5
    x = 26: lblUserUnknowns(x).Caption = "nothn6": txtUserUnknowns(x).Text = SInt2UInt(Userrec.Nothing6)
    x = 27: lblUserUnknowns(x).Caption = "nothn7(0)": txtUserUnknowns(x).Text = SLong2ULong(Userrec.nothing7(0))
    x = 28: lblUserUnknowns(x).Caption = "nothn7(1)": txtUserUnknowns(x).Text = SLong2ULong(Userrec.nothing7(1))
    x = 29: lblUserUnknowns(x).Caption = "nothn7(2)": txtUserUnknowns(x).Text = SLong2ULong(Userrec.nothing7(2))
    x = 30: lblUserUnknowns(x).Caption = "nothn8": txtUserUnknowns(x).Text = SInt2UInt(Userrec.nothing8)
    x = 31: lblUserUnknowns(x).Caption = "nothn9": txtUserUnknowns(x).Text = SInt2UInt(Userrec.nothing9)
    x = 32: lblUserUnknowns(x).Caption = "nothn10": txtUserUnknowns(x).Text = SLong2ULong(Userrec.nothing10)
    
    For x = 0 To 12
        lblUserUnknowns(x + 33).Caption = "unkwn9(" & x & ")": txtUserUnknowns(x + 33).Text = SInt2UInt(Userrec.unknown9(x))
    Next x
    x = 46:  lblUserUnknowns(x).Caption = "Bitmask2": txtUserUnknowns(x).Text = Userrec.Bitmask2
    x = 47:  lblUserUnknowns(x).Caption = "TestFlag1": txtUserUnknowns(x).Text = Userrec.TestFlag1
    x = 48:  lblUserUnknowns(x).Caption = "TestFlag2": txtUserUnknowns(x).Text = Userrec.TestFlag2
    
ElseIf optUserUnknowns(1).Value = True Then 'set2
    counter = 0
    For x = 0 To 19
        lblUserUnknowns(counter).Caption = "unkwn7(" & x & ")": txtUserUnknowns(counter).Text = SInt2UInt(Userrec.unknown7(x))
        counter = counter + 1
    Next x
    For x = 0 To 5
        lblUserUnknowns(counter).Caption = "unkwn11(" & x & ")": txtUserUnknowns(counter).Text = Userrec.unknown11(x)
        counter = counter + 1
    Next x
    For x = 0 To 3
        lblUserUnknowns(counter).Caption = "unkwn12a(" & x & ")": txtUserUnknowns(counter).Text = SInt2UInt(Userrec.unknown12a(x))
        counter = counter + 1
    Next x
    For x = 0 To 8
        lblUserUnknowns(counter).Caption = "unkwn13(" & x & ")": txtUserUnknowns(counter).Text = SInt2UInt(Userrec.unknown13(x))
        If x = 2 Then lblUserUnknowns(counter).Caption = "movemnt?"
        counter = counter + 1
    Next x
    For x = 0 To 3
        lblUserUnknowns(counter).Caption = "unkwn15(" & x & ")": txtUserUnknowns(counter).Text = SLong2ULong(Userrec.unknown15(x))
        counter = counter + 1
    Next x
    For x = 0 To 5
        lblUserUnknowns(counter).Caption = "unkwn9a(" & x & ")": txtUserUnknowns(counter).Text = SLong2ULong(Userrec.unknown9a(x))
        If x = 4 Then lblUserUnknowns(counter).Caption = "partyrank"
        counter = counter + 1
    Next x
    
    For counter = counter To 48
        lblUserUnknowns(counter).Caption = "": txtUserUnknowns(counter).Text = ""
    Next counter
    
ElseIf optUserUnknowns(2).Value = True Then 'set3
    counter = 0
    For x = 0 To 18
        lblUserUnknowns(counter).Caption = "unkwn12d(" & x & ")": txtUserUnknowns(counter).Text = SInt2UInt(Userrec.unknown12d(x))
        If x = 5 Then lblUserUnknowns(counter).Caption = "Encum %"
        If x = 6 Then lblUserUnknowns(counter).Caption = "AccyAbl22"
        counter = counter + 1
    Next x
    
    lblUserUnknowns(counter).Caption = "unkwn12e": txtUserUnknowns(counter).Text = Userrec.unknown12e
    counter = counter + 1
    
    For x = 0 To 9
        lblUserUnknowns(counter).Caption = "unkwn12f(" & x & ")": txtUserUnknowns(counter).Text = SInt2UInt(Userrec.unknown12f(x))
        counter = counter + 1
    Next x
    
    For x = 0 To 7
        lblUserUnknowns(counter).Caption = "UFlags(" & x & ")": txtUserUnknowns(counter).Text = SInt2UInt(Userrec.SomeUserFlags(x))
        If x = 1 Then lblUserUnknowns(counter).Caption = "StatusMask"
        counter = counter + 1
    Next x
    
    For x = 0 To 2
        lblUserUnknowns(counter).Caption = "Energy(" & x & ")": txtUserUnknowns(counter).Text = SInt2UInt(Userrec.Energy(x))
        counter = counter + 1
    Next x
    
    
    For counter = counter To 48
        lblUserUnknowns(counter).Caption = "": txtUserUnknowns(counter).Text = ""
    Next counter
    
End If

Call LoadUserItems(False)
Call CheckWornItems

'bDontSetup = False
Me.MousePointer = 0
Exit Sub
error:
Me.MousePointer = 0
Call HandleError
MsgBox "Warning, record was not completely displayed." & vbCrLf _
    & "Previous records stats may still be in memory.  Select 'Disable DB Writing'" & vbCrLf _
    & "from the file menu and then reload the editor.", vbExclamation
End Sub

Private Sub txtAbilityB_GotFocus(Index As Integer)
Call SelectAll(txtAbilityB(Index))

End Sub

Private Sub txtBBSName_GotFocus()
Call SelectAll(txtBBSName)

End Sub

Private Sub txtBroadcastChan_GotFocus()
Call SelectAll(txtBroadcastChan)

End Sub

Private Sub txtCopper_GotFocus()
Call SelectAll(txtCopper)

End Sub

Private Sub txtCP_GotFocus()
Call SelectAll(txtCp)

End Sub

Private Sub txtCurrentEncum_GotFocus()
Call SelectAll(txtCurrentEncum)

End Sub

Private Sub txtCurrentHP_GotFocus()
Call SelectAll(txtCurrentHP)
End Sub

Private Sub txtCurrentMana_GotFocus()
Call SelectAll(txtCurrentMana)
End Sub

Private Sub txtCurrentMap_Change()

On Error GoTo error:

If Val(txtCurrentMap.Text) > 0 And Val(txtCurrentRoom.Text) > 0 Then
    txtCurrRoomDisp.Text = GetRoomName(Val(txtCurrentMap.Text), _
        Val(txtCurrentRoom.Text))
Else
    txtCurrRoomDisp.Text = ""
End If

out:
Exit Sub
error:
Call HandleError("txtCurrentMap_Change")
Resume out:

End Sub

Private Sub txtCurrentMap_GotFocus()
Call SelectAll(txtCurrentMap)

End Sub

Private Sub txtCurrentRoom_Change()

If Val(txtCurrentMap.Text) > 0 And Val(txtCurrentRoom.Text) > 0 Then
    txtCurrRoomDisp.Text = GetRoomName(Val(txtCurrentMap.Text), _
        Val(txtCurrentRoom.Text))
Else
    txtCurrRoomDisp.Text = ""
End If
End Sub

Private Sub txtCurrentRoom_GotFocus()
Call SelectAll(txtCurrentRoom)

End Sub

Private Sub txtEvilPoints_GotFocus()
Call SelectAll(txtEvilPoints)

End Sub

Private Sub txtExperience_GotFocus()
Call SelectAll(txtExperience)

End Sub

Private Sub txtFirstName_GotFocus()
Call SelectAll(txtFirstName)

End Sub

Private Sub txtGang_GotFocus()
Call SelectAll(txtGang)

End Sub

Private Sub txtGold_GotFocus()
Call SelectAll(txtGold)

End Sub

Private Sub txtHitPointRolls_GotFocus()
Call SelectAll(txtHitPointRolls)
End Sub

Private Sub txtLastName_GotFocus()
Call SelectAll(txtLastName)

End Sub

Private Sub txtLevel_Change()
'Dim nClassExp As Integer, nRaceExp As Integer, nStatus As Integer, nRecord As Integer, nExp As Currency
'On Error GoTo error:
'
'If chkSetupChar.Value = 1 And bDontSetup = False And Val(txtLevel.Text) > 0 Then
'    If cmbClasses.ListIndex > 0 Then
'        If cmbClasses.ItemData(cmbClasses.ListIndex) <= 0 Then GoTo no_exp:
'        nRecord = cmbClasses.ItemData(cmbClasses.ListIndex)
'        nStatus = BTRCALL(BGETEQUAL, ClassPosBlock, Classdatabuf, Len(Classdatabuf), nRecord, KEY_BUF_LEN, 0)
'        Call ClassRowToStruct(Classdatabuf.buf)
'        If nStatus = 0 Then
'            nClassExp = Classrec.Exp + 100
'        Else
'            MsgBox "Setup Char -- Error getting class: " & BtrieveErrorCode(nStatus)
'            Exit Sub
'        End If
'    End If
'
'    If cmbRaces.ListIndex > 0 Then
'        If cmbRaces.ItemData(cmbRaces.ListIndex) <= 0 Then GoTo no_exp:
'        nRecord = cmbRaces.ItemData(cmbRaces.ListIndex)
'        nStatus = BTRCALL(BGETEQUAL, RacePosBlock, Racedatabuf, Len(Racedatabuf), nRecord, KEY_BUF_LEN, 0)
'        Call RaceRowToStruct(Racedatabuf.buf)
'        If nStatus = 0 Then
'            nRaceExp = Racerec.ExpChart
'        Else
'            MsgBox "Setup Char -- Error getting Race: " & BtrieveErrorCode(nStatus)
'            Exit Sub
'        End If
'    End If
'
'    nExp = CalcExpNeeded(CLng(Val(txtLevel.Text)), CInt(nClassExp + nRaceExp))
'    If nExp > 0 Then txtExperience.Text = CStr(nExp + 1)
'End If
'no_exp:
'
'out:
'On Error Resume Next
'Exit Sub
'error:
'Call HandleError("txtLevel_Change")
'Resume out:
End Sub

Private Sub txtLevel_GotFocus()
Call SelectAll(txtLevel)

End Sub

Private Sub txtLives_GotFocus()
Call SelectAll(txtLives)

End Sub

Private Sub txtMapTrail_Change(Index As Integer)
On Error GoTo error:

If Val(txtMapTrail(Index)) > 0 And Val(txtRoomTrail(Index)) > 0 _
    And chkNoRoomNames.Value = 0 Then
    txtRoomTrailDisp(Index).Text = GetRoomName(Val(txtMapTrail(Index).Text), _
        Val(txtRoomTrail(Index).Text))
Else
    txtRoomTrailDisp(Index).Text = ""
End If

out:
Exit Sub
error:
Call HandleError("txtMapTrail_Change")
Resume out:
End Sub

Private Sub txtMapTrail_GotFocus(Index As Integer)
Call SelectAll(txtMapTrail(Index))

End Sub

Private Sub txtMaxEncum_GotFocus()
Call SelectAll(txtMaxEncum)

End Sub

Private Sub txtMaxHP_GotFocus()
Call SelectAll(txtMaxHP)
End Sub

Private Sub txtMaxMana_GotFocus()
Call SelectAll(txtMaxMana)
End Sub

Private Sub txtPlatinum_GotFocus()
Call SelectAll(txtPlatinum)

End Sub

Private Sub txtRoomTrail_Change(Index As Integer)
On Error GoTo error:

If Val(txtMapTrail(Index)) > 0 And Val(txtRoomTrail(Index)) > 0 _
    And chkNoRoomNames.Value = 0 Then
    txtRoomTrailDisp(Index).Text = GetRoomName(Val(txtMapTrail(Index).Text), _
        Val(txtRoomTrail(Index).Text))
Else
    txtRoomTrailDisp(Index).Text = ""
End If

out:
Exit Sub
error:
Call HandleError("txtRoomTrail_Change")
Resume out:
End Sub

Private Sub txtRoomTrail_GotFocus(Index As Integer)
Call SelectAll(txtRoomTrail(Index))

End Sub

Private Sub txtRunic_GotFocus()
Call SelectAll(txtRunic)

End Sub

Private Sub txtSearch_GotFocus()
Call SelectAll(txtSearch)

End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Long, SearchStart As Long, SearchAgain As Boolean, SelectText As String, temp As String

On Error GoTo error:

If KeyCode = vbKeyUp Then Exit Sub
If KeyCode = vbKeyLeft Then Exit Sub
If KeyCode = vbKeyBack Then Exit Sub
If KeyCode = vbKeyDown Then lvDatabase.SetFocus
If KeyCode = vbKeyRight Then SearchAgain = True
If KeyCode = vbKeyControl Then Exit Sub 'control
If KeyCode = 18 Then Exit Sub 'alt
If KeyCode = vbKeyTab Then Exit Sub 'tab
If KeyCode = vbKeyShift Then Exit Sub

SelectText = txtSearch.Text

If SearchAgain = True Then
    SearchStart = lvDatabase.SelectedItem.Index + 1
Else
    SearchStart = 1
End If

If Not SearchStart + 1 <= lvDatabase.ListItems.Count Then Exit Sub

For i = SearchStart To lvDatabase.ListItems.Count
    
    temp = lvDatabase.ListItems(i).Text 'mud name
    If Not InStr(1, UCase(temp), UCase(SelectText)) = 0 Then
        Set lvDatabase.SelectedItem = lvDatabase.ListItems(i)
        lvDatabase.SelectedItem.EnsureVisible
        Call lvDatabase_ItemClick(lvDatabase.SelectedItem)
        Exit For
    End If
    
    temp = lvDatabase.ListItems(i).ListSubItems(1).Text 'bbs name
    If Not InStr(1, UCase(temp), UCase(SelectText)) = 0 Then
        Set lvDatabase.SelectedItem = lvDatabase.ListItems(i)
        lvDatabase.SelectedItem.EnsureVisible
        Call lvDatabase_ItemClick(lvDatabase.SelectedItem)
        Exit For
    End If

Next

Exit Sub
error:
Call HandleError

End Sub
Private Sub cmdEditItem_Click()
On Error GoTo error:
Dim ItemNum As Long, ItemUses As Integer, temp As String, nTmp As Integer

temp = InputBox("Enter new Item Number (0 for none)", "Changing inventory item #" & lstItems.ListIndex, UserItem(lstItems.ListIndex))
If temp = "" Then Exit Sub
ItemNum = ULong2SLong(Val(temp))

temp = InputBox("Enter the number of uses remaining, enter -1 if N/A", "Changing inventory item #" & lstItems.ListIndex, UserItemUses(lstItems.ListIndex))
If temp = "" Then Exit Sub
ItemUses = Val(temp)

nTmp = lstItems.ListIndex

UserItem(lstItems.ListIndex) = ItemNum
UserItemUses(lstItems.ListIndex) = ItemUses

Call LoadUserItems(True)
Call CheckWornItems

lstItems.ListIndex = nTmp

Exit Sub
error:
Call HandleError
End Sub

Private Sub CheckWornItems()
On Error GoTo error:
Dim x As Integer, y As Integer, bItemFound As Boolean

For x = 0 To 19
    If Val(txtWornItem(x).Text) > 0 Then
        bItemFound = False
        For y = 0 To 99
            If UserItem(y) = Val(txtWornItem(x).Text) Then bItemFound = True
        Next y
        If bItemFound = False Then txtWornItem(x).Text = 0
    End If
Next x

If Val(txtWeaponNumber.Text) > 0 Then
    bItemFound = False
    For y = 0 To 99
        If UserItem(y) = Val(txtWeaponNumber.Text) Then bItemFound = True
    Next y
    If bItemFound = False Then txtWeaponNumber.Text = 0
End If

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("CheckWornItems")
Resume out:
End Sub

Private Sub LoadUserItems(RefreshOnly As Boolean)
On Error GoTo error:
Dim x As Integer

If RefreshOnly Then GoTo refresh:

For x = 0 To 49
    UserKey(x) = Userrec.Key(x)
    UserKeyUses(x) = Userrec.KeyUses(x)

    UserItem(x) = Userrec.Item(x)
    UserItemUses(x) = Userrec.ItemUses(x)
    
    UserSpell(x) = Userrec.Spell(x)
Next x

For x = 50 To 99
    UserItem(x) = Userrec.Item(x)
    UserItemUses(x) = Userrec.ItemUses(x)
    
    UserSpell(x) = Userrec.Spell(x)
Next x

refresh:
lstItems.clear
lstKeys.clear
lstSpells.clear

For x = 0 To 49
    lstKeys.AddItem x & vbTab & UserKey(x) & vbTab & UserKeyUses(x) & vbTab & GetItemName(UserKey(x))
    lstItems.AddItem x & vbTab & UserItem(x) & vbTab & UserItemUses(x) & vbTab & GetItemName(UserItem(x))
    lstSpells.AddItem x & vbTab & UserSpell(x) & vbTab & GetShortSpellName(UserSpell(x)) & vbTab & GetSpellName(UserSpell(x))
Next x

For x = 50 To 99
    lstItems.AddItem x & vbTab & UserItem(x) & vbTab & UserItemUses(x) & vbTab & GetItemName(UserItem(x))
    lstSpells.AddItem x & vbTab & UserSpell(x) & vbTab & GetShortSpellName(UserSpell(x)) & vbTab & GetSpellName(UserSpell(x))
Next x

lstItems.ListIndex = 0
lstSpells.ListIndex = 0
lstKeys.ListIndex = 0

Exit Sub
error:
Call HandleError
End Sub

Private Sub cmdEditKey_Click()
On Error GoTo error:
Dim KeyNum As Long, KeyUses As Integer, temp As String, nTmp As Integer

temp = InputBox("Enter new Key number (0 for none)", "Changing inventory Key #" & lstKeys.ListIndex, UserKey(lstKeys.ListIndex))
If temp = "" Then Exit Sub
KeyNum = ULong2SLong(Val(temp))

temp = InputBox("Enter the number of uses remaining", "Changing inventory Key #" & lstKeys.ListIndex, UserKeyUses(lstKeys.ListIndex))
If temp = "" Then Exit Sub
KeyUses = Val(temp)

nTmp = lstKeys.ListIndex

UserKey(lstKeys.ListIndex) = KeyNum
UserKeyUses(lstKeys.ListIndex) = KeyUses

Call LoadUserItems(True)

lstKeys.ListIndex = nTmp
Exit Sub
error:
Call HandleError
End Sub
Private Sub cmdEditSpell_Click()
On Error GoTo error:
Dim SpellNum As Integer, temp As String, nTmp As Integer

temp = InputBox("Enter new Spell number (0 for none)", "Changing spellbook #" & lstSpells.ListIndex, UserSpell(lstSpells.ListIndex))
If temp = "" Then Exit Sub
SpellNum = UInt2SInt(Val(temp))

nTmp = lstSpells.ListIndex

UserSpell(lstSpells.ListIndex) = SpellNum

Call LoadUserItems(True)

lstSpells.ListIndex = nTmp
Exit Sub
error:
Call HandleError
End Sub

Private Sub saverecord(ByVal sRecord As String)
On Error GoTo error:
Dim x As Integer, nStatus As Integer

nStatus = BTRCALL(BGETEQUAL, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal sRecord, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error on save, BGETQUAL: " & BtrieveErrorCode(nStatus)
    Exit Sub
Else
    UserRowToStruct Userdatabuf.buf
End If

Call FormValuesToRecord

nStatus = UpdateUser
If Not nStatus = 0 Then
    MsgBox "SaveRecord, BUPDATE: " & BtrieveErrorCode(nStatus)
Else
    DispUserInfo Userdatabuf.buf
End If

Exit Sub
error:
Call HandleError
End Sub
Private Sub cmdSave_Click()

On Error GoTo error:

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
Call saverecord(sCurrentRecord)

Exit Sub
error:
Call HandleError
End Sub
Private Sub cmdDelete_Click()
Dim nStatus As Integer
Dim nDelete As Integer, BBSName As String, temp As Long

On Error GoTo error:

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

If lvDatabase.ListItems.Count = 0 Then Exit Sub

nDelete = MsgBox("Are you sure you want to Delete this user?", vbYesNo, "Delete User?")
If Not nDelete = vbYes Then Exit Sub

If bLoaded = True And chkAutoSave.Value = 1 Then Call saverecord(sCurrentRecord)

temp = lvDatabase.SelectedItem.Index

BBSName = lvDatabase.SelectedItem.SubItems(1)
BBSName = BBSName & String(30 - Len(BBSName), vbNullChar)

nStatus = BTRCALL(BGETEQUAL, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal BBSName, KEY_BUF_LEN, 0)
If nStatus = 0 Then
    nStatus = BTRCALL(BDELETE, UserPosBlock, Userdatabuf, Len(Userdatabuf), ByVal UserKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "cmdDelete, BDELETE, Error: " & BtrieveErrorCode(nStatus)
    Else
        lvDatabase.ListItems.Remove temp
        sCurrentRecord = String(30, Chr(0))
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
    MsgBox "Couldn't Get Record,Error: " & BtrieveErrorCode(nStatus)
End If


Exit Sub
error:
Call HandleError
End Sub
Private Sub FormValuesToRecord()
On Error GoTo error:
Dim x As Integer

'Userrec.Bitmask1 = Val(Text4.Text)
'Userrec.unknown12(13) = Val(Text6.Text)
'Userrec.unknown13(x) = 0

'DoEvents
Call CheckWornItems

Userrec.FirstName = Trim(RemoveCharacter(txtFirstName.Text, " ")) & Chr(0)
Userrec.LastName = Trim(RemoveCharacter(txtLastName.Text, " ")) & Chr(0)
Userrec.Race = cmbRaces.ListIndex
Userrec.Class = cmbClasses.ListIndex
Userrec.Level = UInt2SInt(Val(txtLevel.Text))

If Val(txtExperience.Text) > 1000000000 Then
    Userrec.BillionsOfExperience = ULong2SLong((Val(txtExperience.Text) - Val(Right(txtExperience.Text, 9))) / 1000000000)
    Userrec.MillionsOfExperience = ULong2SLong(Val(Right(txtExperience.Text, 9)))
Else
    Userrec.BillionsOfExperience = 0
    Userrec.MillionsOfExperience = ULong2SLong(Val(txtExperience.Text))
End If

Userrec.MaxHP = Val(txtMaxHP.Text)
Userrec.CurrentHP = Val(txtCurrentHP.Text)
Userrec.MaxMana = Val(txtMaxMana.Text)
Userrec.CurrentMana = Val(txtCurrentMana.Text)
Userrec.SpellCasting = Val(txtSpellcasting.Text)
Userrec.LivesRemaining = Val(txtLives.Text)
Userrec.CPRemaining = Val(txtCp.Text)
Userrec.Perception = Val(txtPerception.Text)
Userrec.Stealth = Val(txtStealth.Text)
Userrec.Thievery = Val(txtThievery.Text)
Userrec.Traps = Val(txtTraps.Text)
Userrec.Picklocks = Val(txtPicklocks.Text)
Userrec.Tracking = Val(txtTracking.Text)
Userrec.MartialArts = Val(txtMartialArts.Text)
Userrec.MagicRes = Val(txtMagicResistance.Text)
Userrec.MagicRes2 = Val(txtMagicResistance.Text)
Userrec.BroadcastChan = Val(txtBroadcastChan.Text)
Userrec.Runic = ULong2SLong(Val(txtRunic.Text))
Userrec.Platinum = ULong2SLong(Val(txtPlatinum.Text))
Userrec.Gold = ULong2SLong(Val(txtGold.Text))
Userrec.Silver = ULong2SLong(Val(txtSilver.Text))
Userrec.Copper = ULong2SLong(Val(txtCopper.Text))
Userrec.CurrentENC = UInt2SInt(Val(txtCurrentEncum.Text))
Userrec.MaxENC = UInt2SInt(Val(txtMaxEncum.Text))
Userrec.EvilPoints = Val(txtEvilPoints.Text)
Userrec.GangName = LTrim(RTrim(txtGang.Text) & Chr(0))
Userrec.SuicidePassword = LTrim(RTrim(txtSuicide.Text) & Chr(0))
Userrec.Title = LTrim(RTrim(txtTitle.Text) & Chr(0))
Userrec.RoomNum = ULong2SLong(Val(txtCurrentRoom.Text))
Userrec.MapNumber = ULong2SLong(Val(txtCurrentMap.Text))
Userrec.WeaponHand = ULong2SLong(Val(txtWeaponNumber.Text))

If Val(txtHitPointRolls.Text) > 255 Then txtHitPointRolls.Text = 255
Userrec.HitPointRolls = Val(txtHitPointRolls.Text)

If chkEdited.Value = 0 Then
    Userrec.bEDITED = 0
Else
    Userrec.bEDITED = 1
End If

For x = 0 To 11
    Userrec.Stat(x) = Val(txtStat(x).Text)
Next

For x = 0 To 29
    Userrec.Ability(x) = Val(txtAbilityA(x).Text)
    Userrec.AbilityModifier(x) = Val(txtAbilityB(x).Text)
Next x

For x = 0 To 9
    Userrec.SpellCasted(x) = UInt2SInt(Val(txtSpellNumber(x).Text))
    Userrec.SpellValue(x) = UInt2SInt(Val(txtSpellValue(x).Text))
    Userrec.SpellRoundsLeft(x) = UInt2SInt(Val(txtSpellRounds(x).Text))
Next

Userrec.SomeSortOfFlag = Val(txtSomeSortOfFlag.Text)

For x = 0 To 19
    Userrec.LastMap(x) = ULong2SLong(Val(txtMapTrail(x).Text))
    Userrec.LastRoom(x) = ULong2SLong(Val(txtRoomTrail(x).Text))
    Userrec.WornItem(x) = ULong2SLong(Val(txtWornItem(x).Text))
Next

For x = 0 To 99
    Userrec.Item(x) = UserItem(x)
    Userrec.ItemUses(x) = UserItemUses(x)
    Userrec.Spell(x) = UserSpell(x)
Next

For x = 0 To 49
    Userrec.Key(x) = UserKey(x)
    Userrec.KeyUses(x) = UserKeyUses(x)
Next

Exit Sub
error:
Call HandleError
End Sub
Private Sub cmdEditCurrentRoom_Click()
On Error GoTo error:

    Call frmRoom.GotoRoom(Val(txtCurrentMap.Text), Val(txtCurrentRoom.Text))

Exit Sub
error:
Call HandleError
End Sub

Private Sub cmdEditSpellCasted_Click(Index As Integer)
On Error GoTo error:

    Call frmSpell.GotoSpell(Val(txtSpellNumber(Index).Text))

Exit Sub
error:
Call HandleError
End Sub

Private Sub cmdEditWeapon_Click()
On Error GoTo error:

    Call frmItem.GotoItem(Val(txtWeaponNumber.Text))

Exit Sub
error:
Call HandleError
End Sub

Private Sub cmdEditWornItem_Click(Index As Integer)
On Error GoTo error:

    Call frmItem.GotoItem(Val(txtWornItem(Index).Text))

Exit Sub
error:
Call HandleError
End Sub

Private Sub cmdGotoRoom_Click(Index As Integer)
On Error GoTo error:

    Call frmRoom.GotoRoom(Val(txtMapTrail(Index).Text), Val(txtRoomTrail(Index).Text))

Exit Sub
error:
Call HandleError
End Sub

Private Sub cmdItemEditor_Click()
On Error GoTo error:

    Call frmItem.GotoItem(UserItem(lstItems.ListIndex))

Exit Sub
error:
Call HandleError
End Sub

Private Sub cmdItemEditorKey_Click()
On Error GoTo error:

    Call frmItem.GotoItem(UserKey(lstKeys.ListIndex))

Exit Sub
error:
Call HandleError
End Sub

Private Sub cmdSpellEditor_Click()
On Error GoTo error:

    Call frmSpell.GotoSpell(UserSpell(lstSpells.ListIndex))

Exit Sub
error:
Call HandleError
End Sub

Private Sub txtSilver_GotFocus()
Call SelectAll(txtSilver)

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

Private Sub txtSpellRounds_GotFocus(Index As Integer)
Call SelectAll(txtSpellRounds(Index))

End Sub

Private Sub txtSpellValue_GotFocus(Index As Integer)
Call SelectAll(txtSpellValue(Index))

End Sub

Private Sub txtSuicide_GotFocus()
Call SelectAll(txtSuicide)

End Sub

Private Sub txtTitle_GotFocus()
Call SelectAll(txtTitle)

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

Private Sub txtWornItem_Change(Index As Integer)
On Error GoTo error:

txtWornItemName(Index).Text = GetItemName(Val(txtWornItem(Index).Text))

out:
Exit Sub
error:
Call HandleError("txtWornItem_Change")
Resume out:
End Sub

Private Sub txtWornItem_GotFocus(Index As Integer)
Call SelectAll(txtWornItem(Index))

End Sub
