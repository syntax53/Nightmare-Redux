VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMassRoomEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mass Room Editor"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   Icon            =   "frmMassRoomEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   7395
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Cancel / Close"
      Height          =   555
      Left            =   6480
      TabIndex        =   10
      Top             =   5340
      Width           =   855
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
      Height          =   555
      Left            =   6480
      TabIndex        =   9
      Top             =   4560
      Width           =   855
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   7455
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.CommandButton cmdSelectNone 
         Caption         =   "Select None"
         Height          =   315
         Left            =   5220
         TabIndex        =   13
         Top             =   60
         Width           =   1095
      End
      Begin VB.TextBox txtEndRoom 
         Height          =   285
         Left            =   6420
         TabIndex        =   8
         Text            =   "2"
         Top             =   3780
         Width           =   795
      End
      Begin VB.TextBox txtStartRoom 
         Height          =   285
         Left            =   6420
         TabIndex        =   5
         Text            =   "1"
         Top             =   2460
         Width           =   795
      End
      Begin VB.TextBox txtStartMap 
         Height          =   285
         Left            =   6420
         TabIndex        =   3
         Text            =   "1"
         Top             =   1860
         Width           =   795
      End
      Begin VB.TextBox txtName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1020
         MaxLength       =   53
         TabIndex        =   16
         Text            =   "Name"
         Top             =   120
         Width           =   2535
      End
      Begin VB.CheckBox chkName 
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
         Left            =   60
         TabIndex        =   15
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdDeleteRange 
         Caption         =   "Delete Range"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   6480
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   6360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "Select All"
         Height          =   315
         Left            =   4260
         TabIndex        =   14
         Top             =   60
         Width           =   975
      End
      Begin VB.CommandButton cmdLog 
         Caption         =   "Log"
         Height          =   435
         Left            =   6660
         TabIndex        =   12
         Top             =   60
         Width           =   555
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6570
         Left            =   60
         TabIndex        =   17
         Top             =   480
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   11589
         _Version        =   393216
         Style           =   1
         Tabs            =   6
         TabsPerRow      =   6
         TabHeight       =   520
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "frmMassRoomEditor.frx":08CA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "chkDescription"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frmDescription"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Exits"
         TabPicture(1)   =   "frmMassRoomEditor.frx":08E6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "label(15)"
         Tab(1).Control(1)=   "label(14)"
         Tab(1).Control(2)=   "Para1(16)"
         Tab(1).Control(3)=   "label(17)"
         Tab(1).Control(4)=   "label(18)"
         Tab(1).Control(5)=   "label(19)"
         Tab(1).Control(6)=   "cmbRoomType(9)"
         Tab(1).Control(7)=   "cmbRoomType(8)"
         Tab(1).Control(8)=   "cmbRoomType(7)"
         Tab(1).Control(9)=   "cmbRoomType(6)"
         Tab(1).Control(10)=   "cmbRoomType(5)"
         Tab(1).Control(11)=   "cmbRoomType(4)"
         Tab(1).Control(12)=   "cmbRoomType(3)"
         Tab(1).Control(13)=   "cmbRoomType(2)"
         Tab(1).Control(14)=   "cmbRoomType(1)"
         Tab(1).Control(15)=   "cmbRoomType(0)"
         Tab(1).Control(16)=   "txtRoomExit(0)"
         Tab(1).Control(17)=   "txtRoomExit(1)"
         Tab(1).Control(18)=   "txtRoomExit(2)"
         Tab(1).Control(19)=   "txtRoomExit(3)"
         Tab(1).Control(20)=   "txtRoomExit(4)"
         Tab(1).Control(21)=   "txtRoomExit(5)"
         Tab(1).Control(22)=   "txtRoomExit(6)"
         Tab(1).Control(23)=   "txtRoomExit(7)"
         Tab(1).Control(24)=   "txtRoomExit(8)"
         Tab(1).Control(25)=   "txtRoomExit(9)"
         Tab(1).Control(26)=   "txtRoomPara(0)"
         Tab(1).Control(27)=   "txtRoomPara(1)"
         Tab(1).Control(28)=   "txtRoomPara(2)"
         Tab(1).Control(29)=   "txtRoomPara(3)"
         Tab(1).Control(30)=   "txtRoomPara(4)"
         Tab(1).Control(31)=   "txtRoomPara(5)"
         Tab(1).Control(32)=   "txtRoomPara(6)"
         Tab(1).Control(33)=   "txtRoomPara(7)"
         Tab(1).Control(34)=   "txtRoomPara(8)"
         Tab(1).Control(35)=   "txtRoomPara(9)"
         Tab(1).Control(36)=   "txtRoomWPara(9)"
         Tab(1).Control(37)=   "txtRoomWPara(8)"
         Tab(1).Control(38)=   "txtRoomWPara(7)"
         Tab(1).Control(39)=   "txtRoomWPara(6)"
         Tab(1).Control(40)=   "txtRoomWPara(5)"
         Tab(1).Control(41)=   "txtRoomWPara(4)"
         Tab(1).Control(42)=   "txtRoomWPara(3)"
         Tab(1).Control(43)=   "txtRoomWPara(2)"
         Tab(1).Control(44)=   "txtRoomWPara(1)"
         Tab(1).Control(45)=   "txtRoomWPara(0)"
         Tab(1).Control(46)=   "txtRoomLPara1(9)"
         Tab(1).Control(47)=   "txtRoomLPara1(8)"
         Tab(1).Control(48)=   "txtRoomLPara1(7)"
         Tab(1).Control(49)=   "txtRoomLPara1(6)"
         Tab(1).Control(50)=   "txtRoomLPara1(5)"
         Tab(1).Control(51)=   "txtRoomLPara1(4)"
         Tab(1).Control(52)=   "txtRoomLPara1(3)"
         Tab(1).Control(53)=   "txtRoomLPara1(2)"
         Tab(1).Control(54)=   "txtRoomLPara1(1)"
         Tab(1).Control(55)=   "txtRoomLPara1(0)"
         Tab(1).Control(56)=   "txtRoomLPara2(9)"
         Tab(1).Control(57)=   "txtRoomLPara2(8)"
         Tab(1).Control(58)=   "txtRoomLPara2(7)"
         Tab(1).Control(59)=   "txtRoomLPara2(6)"
         Tab(1).Control(60)=   "txtRoomLPara2(5)"
         Tab(1).Control(61)=   "txtRoomLPara2(4)"
         Tab(1).Control(62)=   "txtRoomLPara2(3)"
         Tab(1).Control(63)=   "txtRoomLPara2(2)"
         Tab(1).Control(64)=   "txtRoomLPara2(1)"
         Tab(1).Control(65)=   "txtRoomLPara2(0)"
         Tab(1).Control(66)=   "chkExits(0)"
         Tab(1).Control(67)=   "chkExits(1)"
         Tab(1).Control(68)=   "chkExits(2)"
         Tab(1).Control(69)=   "chkExits(3)"
         Tab(1).Control(70)=   "chkExits(4)"
         Tab(1).Control(71)=   "chkExits(5)"
         Tab(1).Control(72)=   "chkExits(6)"
         Tab(1).Control(73)=   "chkExits(7)"
         Tab(1).Control(74)=   "chkExits(8)"
         Tab(1).Control(75)=   "chkExits(9)"
         Tab(1).ControlCount=   76
         TabCaption(2)   =   "Placed Items/Monster"
         TabPicture(2)   =   "frmMassRoomEditor.frx":0902
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "chkInvisMonies"
         Tab(2).Control(1)=   "frmHiddenCoins"
         Tab(2).Control(2)=   "txtPermNPC"
         Tab(2).Control(3)=   "txtPermNPCName"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "txtPlacedItems(9)"
         Tab(2).Control(5)=   "txtPlacedItems(8)"
         Tab(2).Control(6)=   "txtPlacedItems(7)"
         Tab(2).Control(7)=   "txtPlacedItems(6)"
         Tab(2).Control(8)=   "txtPlacedItems(5)"
         Tab(2).Control(9)=   "txtPlacedItems(4)"
         Tab(2).Control(10)=   "txtPlacedItems(3)"
         Tab(2).Control(11)=   "txtPlacedItems(2)"
         Tab(2).Control(12)=   "txtPlacedItems(1)"
         Tab(2).Control(13)=   "txtPlacedItems(0)"
         Tab(2).Control(14)=   "txtPlacedItemsName(0)"
         Tab(2).Control(14).Enabled=   0   'False
         Tab(2).Control(15)=   "txtPlacedItemsName(1)"
         Tab(2).Control(15).Enabled=   0   'False
         Tab(2).Control(16)=   "txtPlacedItemsName(2)"
         Tab(2).Control(16).Enabled=   0   'False
         Tab(2).Control(17)=   "txtPlacedItemsName(3)"
         Tab(2).Control(17).Enabled=   0   'False
         Tab(2).Control(18)=   "txtPlacedItemsName(4)"
         Tab(2).Control(18).Enabled=   0   'False
         Tab(2).Control(19)=   "txtPlacedItemsName(5)"
         Tab(2).Control(19).Enabled=   0   'False
         Tab(2).Control(20)=   "txtPlacedItemsName(6)"
         Tab(2).Control(20).Enabled=   0   'False
         Tab(2).Control(21)=   "txtPlacedItemsName(7)"
         Tab(2).Control(21).Enabled=   0   'False
         Tab(2).Control(22)=   "txtPlacedItemsName(8)"
         Tab(2).Control(22).Enabled=   0   'False
         Tab(2).Control(23)=   "txtPlacedItemsName(9)"
         Tab(2).Control(23).Enabled=   0   'False
         Tab(2).Control(24)=   "ChkPermNPC"
         Tab(2).Control(25)=   "chkPlacedItems"
         Tab(2).Control(26)=   "chkMonies"
         Tab(2).Control(27)=   "frmVisibleCoins"
         Tab(2).ControlCount=   28
         TabCaption(3)   =   "Visible Items"
         TabPicture(3)   =   "frmMassRoomEditor.frx":091E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "label(49)"
         Tab(3).Control(1)=   "Label12"
         Tab(3).Control(2)=   "txtVisibleUses(0)"
         Tab(3).Control(3)=   "txtVisibleUses(1)"
         Tab(3).Control(4)=   "txtVisibleUses(2)"
         Tab(3).Control(5)=   "txtVisibleUses(3)"
         Tab(3).Control(6)=   "txtVisibleUses(4)"
         Tab(3).Control(7)=   "txtVisibleUses(5)"
         Tab(3).Control(8)=   "txtVisibleUses(6)"
         Tab(3).Control(9)=   "txtVisibleUses(7)"
         Tab(3).Control(10)=   "txtVisibleUses(8)"
         Tab(3).Control(11)=   "txtVisibleUses(9)"
         Tab(3).Control(12)=   "txtVisibleUses(10)"
         Tab(3).Control(13)=   "txtVisibleUses(11)"
         Tab(3).Control(14)=   "txtVisibleUses(12)"
         Tab(3).Control(15)=   "txtVisibleUses(13)"
         Tab(3).Control(16)=   "txtVisibleUses(14)"
         Tab(3).Control(17)=   "txtVisibleUses(15)"
         Tab(3).Control(18)=   "txtVisibleUses(16)"
         Tab(3).Control(19)=   "txtVisibleName(0)"
         Tab(3).Control(19).Enabled=   0   'False
         Tab(3).Control(20)=   "txtVisibleName(1)"
         Tab(3).Control(20).Enabled=   0   'False
         Tab(3).Control(21)=   "txtVisibleName(2)"
         Tab(3).Control(21).Enabled=   0   'False
         Tab(3).Control(22)=   "txtVisibleName(3)"
         Tab(3).Control(22).Enabled=   0   'False
         Tab(3).Control(23)=   "txtVisibleName(4)"
         Tab(3).Control(23).Enabled=   0   'False
         Tab(3).Control(24)=   "txtVisibleName(5)"
         Tab(3).Control(24).Enabled=   0   'False
         Tab(3).Control(25)=   "txtVisibleName(6)"
         Tab(3).Control(25).Enabled=   0   'False
         Tab(3).Control(26)=   "txtVisibleName(7)"
         Tab(3).Control(26).Enabled=   0   'False
         Tab(3).Control(27)=   "txtVisibleName(8)"
         Tab(3).Control(27).Enabled=   0   'False
         Tab(3).Control(28)=   "txtVisibleName(9)"
         Tab(3).Control(28).Enabled=   0   'False
         Tab(3).Control(29)=   "txtVisibleName(10)"
         Tab(3).Control(29).Enabled=   0   'False
         Tab(3).Control(30)=   "txtVisibleName(11)"
         Tab(3).Control(30).Enabled=   0   'False
         Tab(3).Control(31)=   "txtVisibleName(12)"
         Tab(3).Control(31).Enabled=   0   'False
         Tab(3).Control(32)=   "txtVisibleName(13)"
         Tab(3).Control(32).Enabled=   0   'False
         Tab(3).Control(33)=   "txtVisibleName(14)"
         Tab(3).Control(33).Enabled=   0   'False
         Tab(3).Control(34)=   "txtVisibleName(15)"
         Tab(3).Control(34).Enabled=   0   'False
         Tab(3).Control(35)=   "txtVisibleName(16)"
         Tab(3).Control(35).Enabled=   0   'False
         Tab(3).Control(36)=   "chkItemsInRoom"
         Tab(3).Control(37)=   "txtVisibleNumber(9)"
         Tab(3).Control(38)=   "txtVisibleNumber(8)"
         Tab(3).Control(39)=   "txtVisibleNumber(7)"
         Tab(3).Control(40)=   "txtVisibleNumber(6)"
         Tab(3).Control(41)=   "txtVisibleNumber(5)"
         Tab(3).Control(42)=   "txtVisibleNumber(4)"
         Tab(3).Control(43)=   "txtVisibleNumber(3)"
         Tab(3).Control(44)=   "txtVisibleNumber(2)"
         Tab(3).Control(45)=   "txtVisibleNumber(1)"
         Tab(3).Control(46)=   "txtVisibleNumber(0)"
         Tab(3).Control(47)=   "txtVisibleNumber(11)"
         Tab(3).Control(48)=   "txtVisibleNumber(12)"
         Tab(3).Control(49)=   "txtVisibleNumber(13)"
         Tab(3).Control(50)=   "txtVisibleNumber(14)"
         Tab(3).Control(51)=   "txtVisibleNumber(15)"
         Tab(3).Control(52)=   "txtVisibleNumber(16)"
         Tab(3).Control(53)=   "txtVisibleNumber(10)"
         Tab(3).Control(54)=   "txtVisibleQty(0)"
         Tab(3).Control(55)=   "txtVisibleQty(1)"
         Tab(3).Control(56)=   "txtVisibleQty(2)"
         Tab(3).Control(57)=   "txtVisibleQty(3)"
         Tab(3).Control(58)=   "txtVisibleQty(4)"
         Tab(3).Control(59)=   "txtVisibleQty(5)"
         Tab(3).Control(60)=   "txtVisibleQty(6)"
         Tab(3).Control(61)=   "txtVisibleQty(7)"
         Tab(3).Control(62)=   "txtVisibleQty(8)"
         Tab(3).Control(63)=   "txtVisibleQty(9)"
         Tab(3).Control(64)=   "txtVisibleQty(10)"
         Tab(3).Control(65)=   "txtVisibleQty(11)"
         Tab(3).Control(66)=   "txtVisibleQty(12)"
         Tab(3).Control(67)=   "txtVisibleQty(13)"
         Tab(3).Control(68)=   "txtVisibleQty(14)"
         Tab(3).Control(69)=   "txtVisibleQty(15)"
         Tab(3).Control(70)=   "txtVisibleQty(16)"
         Tab(3).ControlCount=   71
         TabCaption(4)   =   "Hidden Items"
         TabPicture(4)   =   "frmMassRoomEditor.frx":093A
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "label(50)"
         Tab(4).Control(1)=   "Label11"
         Tab(4).Control(2)=   "Label13"
         Tab(4).Control(3)=   "txtHiddenUses(0)"
         Tab(4).Control(4)=   "txtHiddenUses(1)"
         Tab(4).Control(5)=   "txtHiddenUses(2)"
         Tab(4).Control(6)=   "txtHiddenUses(3)"
         Tab(4).Control(7)=   "txtHiddenUses(4)"
         Tab(4).Control(8)=   "txtHiddenUses(5)"
         Tab(4).Control(9)=   "txtHiddenUses(6)"
         Tab(4).Control(10)=   "txtHiddenUses(7)"
         Tab(4).Control(11)=   "txtHiddenUses(8)"
         Tab(4).Control(12)=   "txtHiddenUses(9)"
         Tab(4).Control(13)=   "txtHiddenUses(10)"
         Tab(4).Control(14)=   "txtHiddenUses(11)"
         Tab(4).Control(15)=   "txtHiddenUses(12)"
         Tab(4).Control(16)=   "txtHiddenUses(13)"
         Tab(4).Control(17)=   "txtHiddenUses(14)"
         Tab(4).Control(18)=   "txtHiddenQty(14)"
         Tab(4).Control(19)=   "txtHiddenQty(13)"
         Tab(4).Control(20)=   "txtHiddenQty(12)"
         Tab(4).Control(21)=   "txtHiddenQty(11)"
         Tab(4).Control(22)=   "txtHiddenQty(10)"
         Tab(4).Control(23)=   "txtHiddenQty(9)"
         Tab(4).Control(24)=   "txtHiddenQty(8)"
         Tab(4).Control(25)=   "txtHiddenQty(7)"
         Tab(4).Control(26)=   "txtHiddenQty(6)"
         Tab(4).Control(27)=   "txtHiddenQty(5)"
         Tab(4).Control(28)=   "txtHiddenQty(4)"
         Tab(4).Control(29)=   "txtHiddenQty(3)"
         Tab(4).Control(30)=   "txtHiddenQty(2)"
         Tab(4).Control(31)=   "txtHiddenQty(1)"
         Tab(4).Control(32)=   "txtHiddenQty(0)"
         Tab(4).Control(33)=   "txtHiddenNumber(10)"
         Tab(4).Control(34)=   "txtHiddenNumber(11)"
         Tab(4).Control(35)=   "txtHiddenNumber(12)"
         Tab(4).Control(36)=   "txtHiddenNumber(13)"
         Tab(4).Control(37)=   "txtHiddenNumber(14)"
         Tab(4).Control(38)=   "txtHiddenNumber(9)"
         Tab(4).Control(39)=   "txtHiddenNumber(8)"
         Tab(4).Control(40)=   "txtHiddenNumber(7)"
         Tab(4).Control(41)=   "txtHiddenNumber(6)"
         Tab(4).Control(42)=   "txtHiddenNumber(5)"
         Tab(4).Control(43)=   "txtHiddenNumber(4)"
         Tab(4).Control(44)=   "txtHiddenNumber(3)"
         Tab(4).Control(45)=   "txtHiddenNumber(2)"
         Tab(4).Control(46)=   "txtHiddenNumber(1)"
         Tab(4).Control(47)=   "txtHiddenNumber(0)"
         Tab(4).Control(48)=   "chkHiddenItems"
         Tab(4).Control(49)=   "txtHiddenName(0)"
         Tab(4).Control(49).Enabled=   0   'False
         Tab(4).Control(50)=   "txtHiddenName(1)"
         Tab(4).Control(50).Enabled=   0   'False
         Tab(4).Control(51)=   "txtHiddenName(2)"
         Tab(4).Control(51).Enabled=   0   'False
         Tab(4).Control(52)=   "txtHiddenName(3)"
         Tab(4).Control(52).Enabled=   0   'False
         Tab(4).Control(53)=   "txtHiddenName(4)"
         Tab(4).Control(53).Enabled=   0   'False
         Tab(4).Control(54)=   "txtHiddenName(5)"
         Tab(4).Control(54).Enabled=   0   'False
         Tab(4).Control(55)=   "txtHiddenName(6)"
         Tab(4).Control(55).Enabled=   0   'False
         Tab(4).Control(56)=   "txtHiddenName(7)"
         Tab(4).Control(56).Enabled=   0   'False
         Tab(4).Control(57)=   "txtHiddenName(8)"
         Tab(4).Control(57).Enabled=   0   'False
         Tab(4).Control(58)=   "txtHiddenName(9)"
         Tab(4).Control(58).Enabled=   0   'False
         Tab(4).Control(59)=   "txtHiddenName(10)"
         Tab(4).Control(59).Enabled=   0   'False
         Tab(4).Control(60)=   "txtHiddenName(11)"
         Tab(4).Control(60).Enabled=   0   'False
         Tab(4).Control(61)=   "txtHiddenName(12)"
         Tab(4).Control(61).Enabled=   0   'False
         Tab(4).Control(62)=   "txtHiddenName(13)"
         Tab(4).Control(62).Enabled=   0   'False
         Tab(4).Control(63)=   "txtHiddenName(14)"
         Tab(4).Control(63).Enabled=   0   'False
         Tab(4).ControlCount=   64
         TabCaption(5)   =   "Other"
         TabPicture(5)   =   "frmMassRoomEditor.frx":0956
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "txtMonNote"
         Tab(5).Control(1)=   "txtCurrentRoomMon(14)"
         Tab(5).Control(2)=   "txtCurrentRoomMon(0)"
         Tab(5).Control(3)=   "txtCurrentRoomMon(13)"
         Tab(5).Control(4)=   "txtCurrentRoomMon(12)"
         Tab(5).Control(5)=   "txtCurrentRoomMon(11)"
         Tab(5).Control(6)=   "txtCurrentRoomMon(10)"
         Tab(5).Control(7)=   "txtCurrentRoomMon(9)"
         Tab(5).Control(8)=   "txtCurrentRoomMon(8)"
         Tab(5).Control(9)=   "txtCurrentRoomMon(7)"
         Tab(5).Control(10)=   "txtCurrentRoomMon(6)"
         Tab(5).Control(11)=   "txtCurrentRoomMon(5)"
         Tab(5).Control(12)=   "txtCurrentRoomMon(4)"
         Tab(5).Control(13)=   "txtCurrentRoomMon(3)"
         Tab(5).Control(14)=   "txtCurrentRoomMon(2)"
         Tab(5).Control(15)=   "txtCurrentRoomMon(1)"
         Tab(5).Control(16)=   "chkCurrentRoomMons"
         Tab(5).ControlCount=   17
         Begin VB.CheckBox chkInvisMonies 
            Caption         =   "Hidden Coins"
            Height          =   195
            Left            =   -71280
            TabIndex        =   344
            Top             =   3120
            Width           =   1275
         End
         Begin VB.Frame frmHiddenCoins 
            Caption         =   "      "
            Height          =   2295
            Left            =   -71520
            TabIndex        =   333
            Top             =   3120
            Width           =   2175
            Begin VB.TextBox txtInvisRunic 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   338
               Text            =   "0"
               Top             =   360
               Width           =   735
            End
            Begin VB.TextBox txtInvisPlatinum 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   337
               Text            =   "0"
               Top             =   720
               Width           =   735
            End
            Begin VB.TextBox txtInvisSilver 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   336
               Text            =   "0"
               Top             =   1440
               Width           =   735
            End
            Begin VB.TextBox txtInvisCopper 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   335
               Text            =   "0"
               Top             =   1800
               Width           =   735
            End
            Begin VB.TextBox txtInvisGold 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   334
               Text            =   "0"
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label Label8 
               Caption         =   "Runic"
               Height          =   255
               Left            =   240
               TabIndex        =   343
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label7 
               Caption         =   "Platinum"
               Height          =   255
               Left            =   240
               TabIndex        =   342
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label6 
               Caption         =   "Gold"
               Height          =   255
               Left            =   240
               TabIndex        =   341
               Top             =   1080
               Width           =   855
            End
            Begin VB.Label Label9 
               Caption         =   "Silver"
               Height          =   255
               Left            =   240
               TabIndex        =   340
               Top             =   1440
               Width           =   855
            End
            Begin VB.Label Label10 
               Caption         =   "Copper"
               Height          =   255
               Left            =   240
               TabIndex        =   339
               Top             =   1800
               Width           =   855
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Advanced"
            Height          =   3135
            Left            =   120
            TabIndex        =   29
            Top             =   3300
            Width           =   6015
            Begin VB.CommandButton cmdCopyPaste 
               Caption         =   "Pas&te"
               Height          =   315
               Index           =   1
               Left            =   4920
               TabIndex        =   65
               Top             =   600
               Width           =   975
            End
            Begin VB.CommandButton cmdCopyPaste 
               Caption         =   "Cop&y"
               Height          =   315
               Index           =   0
               Left            =   4920
               TabIndex        =   64
               Top             =   300
               Width           =   975
            End
            Begin VB.CheckBox chkGangHouseNumber 
               Caption         =   "Gang House #"
               Height          =   195
               Left            =   180
               TabIndex        =   32
               Top             =   660
               Width           =   1455
            End
            Begin VB.CheckBox chkCmdText 
               Caption         =   "CMD Text"
               Height          =   195
               Left            =   180
               TabIndex        =   44
               Top             =   2460
               Width           =   1395
            End
            Begin VB.CheckBox chkExitRoom 
               Caption         =   "Exit Room"
               Height          =   195
               Left            =   180
               TabIndex        =   42
               Top             =   2160
               Width           =   1395
            End
            Begin VB.CheckBox chkDeathRoom 
               Caption         =   "Death Room"
               Height          =   195
               Left            =   180
               TabIndex        =   40
               Top             =   1860
               Width           =   1395
            End
            Begin VB.CheckBox chkMaxRegen 
               Caption         =   "Max Regen"
               Height          =   195
               Left            =   2520
               TabIndex        =   50
               Top             =   900
               Width           =   1275
            End
            Begin VB.CheckBox chkSpell 
               Caption         =   "Room Spell"
               Height          =   195
               Left            =   180
               TabIndex        =   38
               Top             =   1560
               Width           =   1275
            End
            Begin VB.CheckBox chkMaxIndex 
               Caption         =   "Max Index"
               Height          =   195
               Left            =   2520
               TabIndex        =   48
               Top             =   600
               Width           =   1275
            End
            Begin VB.CheckBox chkMinIndex 
               Caption         =   "Min Index"
               Height          =   195
               Left            =   2520
               TabIndex        =   46
               Top             =   300
               Width           =   1275
            End
            Begin VB.CheckBox chkDelay 
               Caption         =   "Delay"
               Height          =   195
               Left            =   2520
               TabIndex        =   52
               Top             =   1200
               Width           =   1275
            End
            Begin VB.CheckBox chkMonsterType 
               Caption         =   "Monster Type"
               Height          =   195
               Left            =   2520
               TabIndex        =   58
               Top             =   2100
               Width           =   1455
            End
            Begin VB.CheckBox chkRoomType 
               Caption         =   "Room Type"
               Height          =   195
               Left            =   2520
               TabIndex        =   60
               Top             =   2415
               Width           =   1455
            End
            Begin VB.CheckBox chkControlRoom 
               Caption         =   "Control Room"
               Height          =   195
               Left            =   2520
               TabIndex        =   54
               Top             =   1500
               Width           =   1455
            End
            Begin VB.CheckBox chkMaxArea 
               Caption         =   "Max Area"
               Height          =   195
               Left            =   2520
               TabIndex        =   56
               Top             =   1800
               Width           =   1455
            End
            Begin VB.CheckBox chkAttributes 
               Caption         =   "Attributes"
               Height          =   195
               Left            =   180
               TabIndex        =   34
               Top             =   960
               Width           =   1455
            End
            Begin VB.CheckBox chkAnsiMap 
               Caption         =   "Ansi Map"
               Height          =   195
               Left            =   2520
               TabIndex        =   62
               Top             =   2760
               Width           =   1455
            End
            Begin VB.CheckBox chkLight 
               Caption         =   "Light"
               Height          =   195
               Left            =   180
               TabIndex        =   30
               Top             =   360
               Width           =   1455
            End
            Begin VB.CheckBox chkShop 
               Caption         =   "Shop #"
               Height          =   195
               Left            =   180
               TabIndex        =   36
               Top             =   1260
               Width           =   1455
            End
            Begin VB.ComboBox cmbType 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "frmMassRoomEditor.frx":0972
               Left            =   4140
               List            =   "frmMassRoomEditor.frx":098E
               Style           =   2  'Dropdown List
               TabIndex        =   61
               Top             =   2415
               Width           =   1755
            End
            Begin VB.ComboBox cmbMonsterType 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "frmMassRoomEditor.frx":09CE
               Left            =   4140
               List            =   "frmMassRoomEditor.frx":0A4A
               Style           =   2  'Dropdown List
               TabIndex        =   59
               Top             =   2100
               Width           =   1755
            End
            Begin VB.TextBox txtAttributes 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1680
               MaxLength       =   5
               TabIndex        =   35
               Text            =   "0"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtExitRoom 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1680
               MaxLength       =   5
               TabIndex        =   43
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtSpell 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1680
               MaxLength       =   5
               TabIndex        =   39
               Text            =   "0"
               Top             =   1560
               Width           =   615
            End
            Begin VB.TextBox txtDelay 
               Enabled         =   0   'False
               Height          =   285
               Left            =   4140
               MaxLength       =   5
               TabIndex        =   53
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtMaxArea 
               Enabled         =   0   'False
               Height          =   285
               Left            =   4140
               MaxLength       =   5
               TabIndex        =   57
               Text            =   "0"
               Top             =   1800
               Width           =   615
            End
            Begin VB.TextBox txtControlRoom 
               Enabled         =   0   'False
               Height          =   285
               Left            =   4140
               MaxLength       =   5
               TabIndex        =   55
               Text            =   "0"
               Top             =   1500
               Width           =   615
            End
            Begin VB.TextBox txtCmdText 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1680
               MaxLength       =   5
               TabIndex        =   45
               Text            =   "0"
               Top             =   2475
               Width           =   615
            End
            Begin VB.TextBox txtDeathRoom 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1680
               MaxLength       =   5
               TabIndex        =   41
               Text            =   "0"
               Top             =   1860
               Width           =   615
            End
            Begin VB.TextBox txtShopNum 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1680
               MaxLength       =   5
               TabIndex        =   37
               Text            =   "0"
               Top             =   1260
               Width           =   615
            End
            Begin VB.TextBox txtAnsiMap 
               Enabled         =   0   'False
               Height          =   285
               Left            =   4140
               MaxLength       =   12
               TabIndex        =   63
               Text            =   "WCCMAP01.ANS"
               Top             =   2760
               Width           =   1755
            End
            Begin VB.TextBox txtMaxIndex 
               Enabled         =   0   'False
               Height          =   285
               Left            =   4140
               MaxLength       =   5
               TabIndex        =   49
               Text            =   "0"
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox txtMinIndex 
               Enabled         =   0   'False
               Height          =   285
               Left            =   4140
               MaxLength       =   5
               TabIndex        =   47
               Text            =   "0"
               Top             =   300
               Width           =   615
            End
            Begin VB.TextBox txtMaxRegen 
               Enabled         =   0   'False
               Height          =   285
               Left            =   4140
               MaxLength       =   5
               TabIndex        =   51
               Text            =   "0"
               Top             =   900
               Width           =   615
            End
            Begin VB.TextBox txtLight 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1680
               MaxLength       =   5
               TabIndex        =   31
               Text            =   "0"
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox txtGangHouseNumber 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1680
               TabIndex        =   33
               Text            =   "0"
               Top             =   660
               Width           =   615
            End
         End
         Begin VB.CheckBox chkExits 
            Caption         =   "D"
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
            Left            =   -74820
            TabIndex        =   135
            Top             =   4680
            Width           =   555
         End
         Begin VB.CheckBox chkExits 
            Caption         =   "U"
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
            Left            =   -74820
            TabIndex        =   128
            Top             =   4260
            Width           =   555
         End
         Begin VB.CheckBox chkExits 
            Caption         =   "SW"
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
            Left            =   -74820
            TabIndex        =   121
            Top             =   3840
            Width           =   675
         End
         Begin VB.CheckBox chkExits 
            Caption         =   "SE"
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
            Left            =   -74820
            TabIndex        =   114
            Top             =   3420
            Width           =   675
         End
         Begin VB.CheckBox chkExits 
            Caption         =   "NW"
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
            Left            =   -74820
            TabIndex        =   107
            Top             =   3000
            Width           =   675
         End
         Begin VB.CheckBox chkExits 
            Caption         =   "NE"
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
            Left            =   -74820
            TabIndex        =   100
            Top             =   2580
            Width           =   675
         End
         Begin VB.CheckBox chkExits 
            Caption         =   "W"
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
            Left            =   -74820
            TabIndex        =   93
            Top             =   2160
            Width           =   555
         End
         Begin VB.CheckBox chkExits 
            Caption         =   "E"
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
            TabIndex        =   86
            Top             =   1740
            Width           =   555
         End
         Begin VB.CheckBox chkExits 
            Caption         =   "S"
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
            TabIndex        =   79
            Top             =   1320
            Width           =   555
         End
         Begin VB.CheckBox chkExits 
            Caption         =   "N"
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
            TabIndex        =   72
            Top             =   900
            Width           =   555
         End
         Begin VB.TextBox txtVisibleQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   16
            Left            =   -70575
            TabIndex        =   218
            Text            =   "0"
            Top             =   5940
            Width           =   555
         End
         Begin VB.TextBox txtVisibleQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   15
            Left            =   -70575
            TabIndex        =   217
            Text            =   "0"
            Top             =   5640
            Width           =   555
         End
         Begin VB.TextBox txtVisibleQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   14
            Left            =   -70575
            TabIndex        =   216
            Text            =   "0"
            Top             =   5325
            Width           =   555
         End
         Begin VB.TextBox txtVisibleQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   13
            Left            =   -70575
            TabIndex        =   215
            Text            =   "0"
            Top             =   5010
            Width           =   555
         End
         Begin VB.TextBox txtVisibleQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   12
            Left            =   -70575
            TabIndex        =   214
            Text            =   "0"
            Top             =   4710
            Width           =   555
         End
         Begin VB.TextBox txtVisibleQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   11
            Left            =   -70575
            TabIndex        =   213
            Text            =   "0"
            Top             =   4395
            Width           =   555
         End
         Begin VB.TextBox txtVisibleQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   10
            Left            =   -70575
            TabIndex        =   212
            Text            =   "0"
            Top             =   4095
            Width           =   555
         End
         Begin VB.TextBox txtVisibleQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   -70575
            TabIndex        =   211
            Text            =   "0"
            Top             =   3780
            Width           =   555
         End
         Begin VB.TextBox txtVisibleQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   -70575
            TabIndex        =   210
            Text            =   "0"
            Top             =   3480
            Width           =   555
         End
         Begin VB.TextBox txtVisibleQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   -70575
            TabIndex        =   209
            Text            =   "0"
            Top             =   3180
            Width           =   555
         End
         Begin VB.TextBox txtVisibleQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   -70575
            TabIndex        =   208
            Text            =   "0"
            Top             =   2865
            Width           =   555
         End
         Begin VB.TextBox txtVisibleQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   -70575
            TabIndex        =   207
            Text            =   "0"
            Top             =   2550
            Width           =   555
         End
         Begin VB.TextBox txtVisibleQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   -70575
            TabIndex        =   206
            Text            =   "0"
            Top             =   2250
            Width           =   555
         End
         Begin VB.TextBox txtVisibleQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   -70575
            TabIndex        =   205
            Text            =   "0"
            Top             =   1935
            Width           =   555
         End
         Begin VB.TextBox txtVisibleQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   -70575
            TabIndex        =   204
            Text            =   "0"
            Top             =   1635
            Width           =   555
         End
         Begin VB.TextBox txtVisibleQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   -70575
            TabIndex        =   203
            Text            =   "0"
            Top             =   1320
            Width           =   555
         End
         Begin VB.TextBox txtVisibleQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   -70575
            TabIndex        =   202
            Text            =   "0"
            Top             =   1020
            Width           =   555
         End
         Begin VB.TextBox txtPermNPC 
            Enabled         =   0   'False
            Height          =   285
            Left            =   -74835
            MaxLength       =   5
            TabIndex        =   164
            Text            =   "0"
            Top             =   4320
            Width           =   615
         End
         Begin VB.TextBox txtPermNPCName 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   -74205
            Locked          =   -1  'True
            TabIndex        =   165
            TabStop         =   0   'False
            Top             =   4320
            Width           =   2295
         End
         Begin VB.TextBox txtRoomLPara2 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   -69705
            MaxLength       =   5
            TabIndex        =   78
            Text            =   "0"
            Top             =   900
            Width           =   615
         End
         Begin VB.TextBox txtRoomLPara2 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   -69705
            MaxLength       =   5
            TabIndex        =   85
            Text            =   "0"
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox txtRoomLPara2 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   -69705
            MaxLength       =   5
            TabIndex        =   92
            Text            =   "0"
            Top             =   1740
            Width           =   615
         End
         Begin VB.TextBox txtRoomLPara2 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   -69705
            MaxLength       =   5
            TabIndex        =   99
            Text            =   "0"
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtRoomLPara2 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   -69705
            MaxLength       =   5
            TabIndex        =   106
            Text            =   "0"
            Top             =   2580
            Width           =   615
         End
         Begin VB.TextBox txtRoomLPara2 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   -69705
            MaxLength       =   5
            TabIndex        =   113
            Text            =   "0"
            Top             =   3000
            Width           =   615
         End
         Begin VB.TextBox txtRoomLPara2 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   -69705
            MaxLength       =   5
            TabIndex        =   120
            Text            =   "0"
            Top             =   3420
            Width           =   615
         End
         Begin VB.TextBox txtRoomLPara2 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   -69705
            MaxLength       =   5
            TabIndex        =   127
            Text            =   "0"
            Top             =   3840
            Width           =   615
         End
         Begin VB.TextBox txtRoomLPara2 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   -69705
            MaxLength       =   5
            TabIndex        =   134
            Text            =   "0"
            Top             =   4260
            Width           =   615
         End
         Begin VB.TextBox txtRoomLPara2 
            Enabled         =   0   'False
            Height          =   315
            Index           =   9
            Left            =   -69705
            MaxLength       =   5
            TabIndex        =   141
            Text            =   "0"
            Top             =   4680
            Width           =   615
         End
         Begin VB.TextBox txtRoomLPara1 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   -70425
            MaxLength       =   5
            TabIndex        =   77
            Text            =   "0"
            Top             =   900
            Width           =   615
         End
         Begin VB.TextBox txtRoomLPara1 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   -70425
            MaxLength       =   5
            TabIndex        =   84
            Text            =   "0"
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox txtRoomLPara1 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   -70425
            MaxLength       =   5
            TabIndex        =   91
            Text            =   "0"
            Top             =   1740
            Width           =   615
         End
         Begin VB.TextBox txtRoomLPara1 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   -70425
            MaxLength       =   5
            TabIndex        =   98
            Text            =   "0"
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtRoomLPara1 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   -70425
            MaxLength       =   5
            TabIndex        =   105
            Text            =   "0"
            Top             =   2580
            Width           =   615
         End
         Begin VB.TextBox txtRoomLPara1 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   -70425
            MaxLength       =   5
            TabIndex        =   112
            Text            =   "0"
            Top             =   3000
            Width           =   615
         End
         Begin VB.TextBox txtRoomLPara1 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   -70425
            MaxLength       =   5
            TabIndex        =   119
            Text            =   "0"
            Top             =   3420
            Width           =   615
         End
         Begin VB.TextBox txtRoomLPara1 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   -70425
            MaxLength       =   5
            TabIndex        =   126
            Text            =   "0"
            Top             =   3840
            Width           =   615
         End
         Begin VB.TextBox txtRoomLPara1 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   -70425
            MaxLength       =   5
            TabIndex        =   133
            Text            =   "0"
            Top             =   4260
            Width           =   615
         End
         Begin VB.TextBox txtRoomLPara1 
            Enabled         =   0   'False
            Height          =   315
            Index           =   9
            Left            =   -70425
            MaxLength       =   5
            TabIndex        =   140
            Text            =   "0"
            Top             =   4680
            Width           =   615
         End
         Begin VB.TextBox txtRoomWPara 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   -71145
            MaxLength       =   5
            TabIndex        =   76
            Text            =   "0"
            Top             =   900
            Width           =   615
         End
         Begin VB.TextBox txtRoomWPara 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   -71145
            MaxLength       =   5
            TabIndex        =   83
            Text            =   "0"
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox txtRoomWPara 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   -71145
            MaxLength       =   5
            TabIndex        =   90
            Text            =   "0"
            Top             =   1740
            Width           =   615
         End
         Begin VB.TextBox txtRoomWPara 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   -71145
            MaxLength       =   5
            TabIndex        =   97
            Text            =   "0"
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtRoomWPara 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   -71145
            MaxLength       =   5
            TabIndex        =   104
            Text            =   "0"
            Top             =   2580
            Width           =   615
         End
         Begin VB.TextBox txtRoomWPara 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   -71145
            MaxLength       =   5
            TabIndex        =   111
            Text            =   "0"
            Top             =   3000
            Width           =   615
         End
         Begin VB.TextBox txtRoomWPara 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   -71145
            MaxLength       =   5
            TabIndex        =   118
            Text            =   "0"
            Top             =   3420
            Width           =   615
         End
         Begin VB.TextBox txtRoomWPara 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   -71145
            MaxLength       =   5
            TabIndex        =   125
            Text            =   "0"
            Top             =   3840
            Width           =   615
         End
         Begin VB.TextBox txtRoomWPara 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   -71145
            MaxLength       =   5
            TabIndex        =   132
            Text            =   "0"
            Top             =   4260
            Width           =   615
         End
         Begin VB.TextBox txtRoomWPara 
            Enabled         =   0   'False
            Height          =   315
            Index           =   9
            Left            =   -71145
            MaxLength       =   5
            TabIndex        =   139
            Text            =   "0"
            Top             =   4680
            Width           =   615
         End
         Begin VB.TextBox txtRoomPara 
            Enabled         =   0   'False
            Height          =   315
            Index           =   9
            Left            =   -71880
            MaxLength       =   5
            TabIndex        =   138
            Text            =   "0"
            Top             =   4680
            Width           =   615
         End
         Begin VB.TextBox txtRoomPara 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   -71880
            MaxLength       =   5
            TabIndex        =   131
            Text            =   "0"
            Top             =   4260
            Width           =   615
         End
         Begin VB.TextBox txtRoomPara 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   -71880
            MaxLength       =   5
            TabIndex        =   124
            Text            =   "0"
            Top             =   3840
            Width           =   615
         End
         Begin VB.TextBox txtRoomPara 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   -71880
            MaxLength       =   5
            TabIndex        =   117
            Text            =   "0"
            Top             =   3420
            Width           =   615
         End
         Begin VB.TextBox txtRoomPara 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   -71880
            MaxLength       =   5
            TabIndex        =   110
            Text            =   "0"
            Top             =   3000
            Width           =   615
         End
         Begin VB.TextBox txtRoomPara 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   -71880
            MaxLength       =   5
            TabIndex        =   103
            Text            =   "0"
            Top             =   2580
            Width           =   615
         End
         Begin VB.TextBox txtRoomPara 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   -71880
            MaxLength       =   5
            TabIndex        =   96
            Text            =   "0"
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtRoomPara 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   -71880
            MaxLength       =   5
            TabIndex        =   89
            Text            =   "0"
            Top             =   1740
            Width           =   615
         End
         Begin VB.TextBox txtRoomPara 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   -71880
            MaxLength       =   5
            TabIndex        =   82
            Text            =   "0"
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox txtRoomPara 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   -71880
            MaxLength       =   5
            TabIndex        =   75
            Text            =   "0"
            Top             =   900
            Width           =   615
         End
         Begin VB.TextBox txtRoomExit 
            Enabled         =   0   'False
            Height          =   315
            Index           =   9
            Left            =   -74145
            MaxLength       =   5
            TabIndex        =   136
            Text            =   "0"
            Top             =   4680
            Width           =   615
         End
         Begin VB.TextBox txtRoomExit 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   -74145
            MaxLength       =   5
            TabIndex        =   129
            Text            =   "0"
            Top             =   4260
            Width           =   615
         End
         Begin VB.TextBox txtRoomExit 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   -74145
            MaxLength       =   5
            TabIndex        =   122
            Text            =   "0"
            Top             =   3840
            Width           =   615
         End
         Begin VB.TextBox txtRoomExit 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   -74145
            MaxLength       =   5
            TabIndex        =   115
            Text            =   "0"
            Top             =   3420
            Width           =   615
         End
         Begin VB.TextBox txtRoomExit 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   -74145
            MaxLength       =   5
            TabIndex        =   108
            Text            =   "0"
            Top             =   3000
            Width           =   615
         End
         Begin VB.TextBox txtRoomExit 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   -74145
            MaxLength       =   5
            TabIndex        =   101
            Text            =   "0"
            Top             =   2580
            Width           =   615
         End
         Begin VB.TextBox txtRoomExit 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   -74145
            MaxLength       =   5
            TabIndex        =   94
            Text            =   "0"
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtRoomExit 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   -74145
            MaxLength       =   5
            TabIndex        =   87
            Text            =   "0"
            Top             =   1740
            Width           =   615
         End
         Begin VB.TextBox txtRoomExit 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   -74145
            MaxLength       =   5
            TabIndex        =   80
            Text            =   "0"
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox txtRoomExit 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   -74145
            MaxLength       =   5
            TabIndex        =   73
            Text            =   "0"
            Top             =   900
            Width           =   615
         End
         Begin VB.ComboBox cmbRoomType 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            ItemData        =   "frmMassRoomEditor.frx":0BB9
            Left            =   -73455
            List            =   "frmMassRoomEditor.frx":0C08
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   900
            Width           =   1455
         End
         Begin VB.ComboBox cmbRoomType 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            ItemData        =   "frmMassRoomEditor.frx":0CD8
            Left            =   -73455
            List            =   "frmMassRoomEditor.frx":0D27
            Style           =   2  'Dropdown List
            TabIndex        =   81
            Top             =   1320
            Width           =   1455
         End
         Begin VB.ComboBox cmbRoomType 
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            ItemData        =   "frmMassRoomEditor.frx":0DF7
            Left            =   -73455
            List            =   "frmMassRoomEditor.frx":0E46
            Style           =   2  'Dropdown List
            TabIndex        =   88
            Top             =   1740
            Width           =   1455
         End
         Begin VB.ComboBox cmbRoomType 
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            ItemData        =   "frmMassRoomEditor.frx":0F16
            Left            =   -73455
            List            =   "frmMassRoomEditor.frx":0F65
            Style           =   2  'Dropdown List
            TabIndex        =   95
            Top             =   2160
            Width           =   1455
         End
         Begin VB.ComboBox cmbRoomType 
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            ItemData        =   "frmMassRoomEditor.frx":1035
            Left            =   -73455
            List            =   "frmMassRoomEditor.frx":1084
            Style           =   2  'Dropdown List
            TabIndex        =   102
            Top             =   2580
            Width           =   1455
         End
         Begin VB.ComboBox cmbRoomType 
            Enabled         =   0   'False
            Height          =   315
            Index           =   5
            ItemData        =   "frmMassRoomEditor.frx":1154
            Left            =   -73455
            List            =   "frmMassRoomEditor.frx":11A3
            Style           =   2  'Dropdown List
            TabIndex        =   109
            Top             =   3000
            Width           =   1455
         End
         Begin VB.ComboBox cmbRoomType 
            Enabled         =   0   'False
            Height          =   315
            Index           =   6
            ItemData        =   "frmMassRoomEditor.frx":1273
            Left            =   -73455
            List            =   "frmMassRoomEditor.frx":12C2
            Style           =   2  'Dropdown List
            TabIndex        =   116
            Top             =   3420
            Width           =   1455
         End
         Begin VB.ComboBox cmbRoomType 
            Enabled         =   0   'False
            Height          =   315
            Index           =   7
            ItemData        =   "frmMassRoomEditor.frx":1392
            Left            =   -73455
            List            =   "frmMassRoomEditor.frx":13E1
            Style           =   2  'Dropdown List
            TabIndex        =   123
            Top             =   3840
            Width           =   1455
         End
         Begin VB.ComboBox cmbRoomType 
            Enabled         =   0   'False
            Height          =   315
            Index           =   8
            ItemData        =   "frmMassRoomEditor.frx":14B1
            Left            =   -73455
            List            =   "frmMassRoomEditor.frx":1500
            Style           =   2  'Dropdown List
            TabIndex        =   130
            Top             =   4260
            Width           =   1455
         End
         Begin VB.ComboBox cmbRoomType 
            Enabled         =   0   'False
            Height          =   315
            Index           =   9
            ItemData        =   "frmMassRoomEditor.frx":15D0
            Left            =   -73455
            List            =   "frmMassRoomEditor.frx":161F
            Style           =   2  'Dropdown List
            TabIndex        =   137
            Top             =   4680
            Width           =   1455
         End
         Begin VB.TextBox txtVisibleNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   10
            Left            =   -73650
            MaxLength       =   5
            TabIndex        =   178
            Text            =   "0"
            Top             =   4095
            Width           =   615
         End
         Begin VB.TextBox txtVisibleNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   16
            Left            =   -73650
            MaxLength       =   5
            TabIndex        =   184
            Text            =   "0"
            Top             =   5940
            Width           =   615
         End
         Begin VB.TextBox txtVisibleNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   15
            Left            =   -73650
            MaxLength       =   5
            TabIndex        =   183
            Text            =   "0"
            Top             =   5625
            Width           =   615
         End
         Begin VB.TextBox txtVisibleNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   14
            Left            =   -73650
            MaxLength       =   5
            TabIndex        =   182
            Text            =   "0"
            Top             =   5325
            Width           =   615
         End
         Begin VB.TextBox txtVisibleNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   13
            Left            =   -73650
            MaxLength       =   5
            TabIndex        =   181
            Text            =   "0"
            Top             =   5010
            Width           =   615
         End
         Begin VB.TextBox txtVisibleNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   12
            Left            =   -73650
            MaxLength       =   5
            TabIndex        =   180
            Text            =   "0"
            Top             =   4710
            Width           =   615
         End
         Begin VB.TextBox txtVisibleNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   11
            Left            =   -73650
            MaxLength       =   5
            TabIndex        =   179
            Text            =   "0"
            Top             =   4395
            Width           =   615
         End
         Begin VB.TextBox txtVisibleNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   -73650
            MaxLength       =   5
            TabIndex        =   168
            Text            =   "0"
            Top             =   1020
            Width           =   615
         End
         Begin VB.TextBox txtVisibleNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   -73650
            MaxLength       =   5
            TabIndex        =   169
            Text            =   "0"
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox txtVisibleNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   -73650
            MaxLength       =   5
            TabIndex        =   170
            Text            =   "0"
            Top             =   1635
            Width           =   615
         End
         Begin VB.TextBox txtVisibleNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   -73650
            MaxLength       =   5
            TabIndex        =   171
            Text            =   "0"
            Top             =   1935
            Width           =   615
         End
         Begin VB.TextBox txtVisibleNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   -73650
            MaxLength       =   5
            TabIndex        =   172
            Text            =   "0"
            Top             =   2250
            Width           =   615
         End
         Begin VB.TextBox txtVisibleNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   -73650
            MaxLength       =   5
            TabIndex        =   173
            Text            =   "0"
            Top             =   2550
            Width           =   615
         End
         Begin VB.TextBox txtVisibleNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   -73650
            MaxLength       =   5
            TabIndex        =   174
            Text            =   "0"
            Top             =   2865
            Width           =   615
         End
         Begin VB.TextBox txtVisibleNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   -73650
            MaxLength       =   5
            TabIndex        =   175
            Text            =   "0"
            Top             =   3165
            Width           =   615
         End
         Begin VB.TextBox txtVisibleNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   -73650
            MaxLength       =   5
            TabIndex        =   176
            Text            =   "0"
            Top             =   3480
            Width           =   615
         End
         Begin VB.TextBox txtVisibleNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   -73650
            MaxLength       =   5
            TabIndex        =   177
            Text            =   "0"
            Top             =   3780
            Width           =   615
         End
         Begin VB.Frame frmDescription 
            Caption         =   "Description"
            Enabled         =   0   'False
            Height          =   2895
            Left            =   360
            TabIndex        =   19
            Top             =   360
            Width           =   5775
            Begin VB.CommandButton cmdCopyDesc 
               Caption         =   "Paste"
               Height          =   315
               Index           =   1
               Left            =   1020
               TabIndex        =   28
               Top             =   2460
               Width           =   975
            End
            Begin VB.CommandButton cmdCopyDesc 
               Caption         =   "Copy"
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   27
               Top             =   2460
               Width           =   915
            End
            Begin VB.TextBox txtDesc 
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   6
               Left            =   120
               MaxLength       =   70
               TabIndex        =   26
               Top             =   2040
               Width           =   5535
            End
            Begin VB.TextBox txtDesc 
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   5
               Left            =   120
               MaxLength       =   70
               TabIndex        =   25
               Top             =   1740
               Width           =   5535
            End
            Begin VB.TextBox txtDesc 
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   4
               Left            =   120
               MaxLength       =   70
               TabIndex        =   24
               Top             =   1440
               Width           =   5535
            End
            Begin VB.TextBox txtDesc 
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   3
               Left            =   120
               MaxLength       =   70
               TabIndex        =   23
               Top             =   1140
               Width           =   5535
            End
            Begin VB.TextBox txtDesc 
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   2
               Left            =   120
               MaxLength       =   70
               TabIndex        =   22
               Top             =   840
               Width           =   5535
            End
            Begin VB.TextBox txtDesc 
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   0
               Left            =   120
               MaxLength       =   70
               TabIndex        =   20
               Text            =   "Room Description"
               Top             =   240
               Width           =   5535
            End
            Begin VB.TextBox txtDesc 
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   1
               Left            =   120
               MaxLength       =   70
               TabIndex        =   21
               Top             =   540
               Width           =   5535
            End
         End
         Begin VB.TextBox txtPlacedItems 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   -74820
            MaxLength       =   5
            TabIndex        =   152
            Text            =   "0"
            Top             =   3660
            Width           =   615
         End
         Begin VB.TextBox txtPlacedItems 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   -74820
            MaxLength       =   5
            TabIndex        =   151
            Text            =   "0"
            Top             =   3360
            Width           =   615
         End
         Begin VB.TextBox txtPlacedItems 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   -74820
            MaxLength       =   5
            TabIndex        =   150
            Text            =   "0"
            Top             =   3060
            Width           =   615
         End
         Begin VB.TextBox txtPlacedItems 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   -74820
            MaxLength       =   5
            TabIndex        =   149
            Text            =   "0"
            Top             =   2760
            Width           =   615
         End
         Begin VB.TextBox txtPlacedItems 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   -74820
            MaxLength       =   5
            TabIndex        =   148
            Text            =   "0"
            Top             =   2460
            Width           =   615
         End
         Begin VB.TextBox txtPlacedItems 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   -74820
            MaxLength       =   5
            TabIndex        =   147
            Text            =   "0"
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtPlacedItems 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   -74820
            MaxLength       =   5
            TabIndex        =   146
            Text            =   "0"
            Top             =   1860
            Width           =   615
         End
         Begin VB.TextBox txtPlacedItems 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   -74820
            MaxLength       =   5
            TabIndex        =   145
            Text            =   "0"
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox txtPlacedItems 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   -74820
            MaxLength       =   5
            TabIndex        =   144
            Text            =   "0"
            Top             =   1260
            Width           =   615
         End
         Begin VB.TextBox txtPlacedItems 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   -74820
            MaxLength       =   5
            TabIndex        =   143
            Text            =   "0"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtPlacedItemsName 
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   0
            Left            =   -74220
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   153
            TabStop         =   0   'False
            Top             =   960
            Width           =   2295
         End
         Begin VB.TextBox txtPlacedItemsName 
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   1
            Left            =   -74220
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   154
            TabStop         =   0   'False
            Top             =   1260
            Width           =   2295
         End
         Begin VB.TextBox txtPlacedItemsName 
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   2
            Left            =   -74220
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   155
            TabStop         =   0   'False
            Top             =   1560
            Width           =   2295
         End
         Begin VB.TextBox txtPlacedItemsName 
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   3
            Left            =   -74220
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   156
            TabStop         =   0   'False
            Top             =   1860
            Width           =   2295
         End
         Begin VB.TextBox txtPlacedItemsName 
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   4
            Left            =   -74220
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   157
            TabStop         =   0   'False
            Top             =   2160
            Width           =   2295
         End
         Begin VB.TextBox txtPlacedItemsName 
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   5
            Left            =   -74220
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   158
            TabStop         =   0   'False
            Top             =   2460
            Width           =   2295
         End
         Begin VB.TextBox txtPlacedItemsName 
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   6
            Left            =   -74220
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   159
            TabStop         =   0   'False
            Top             =   2760
            Width           =   2295
         End
         Begin VB.TextBox txtPlacedItemsName 
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   7
            Left            =   -74220
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   160
            TabStop         =   0   'False
            Top             =   3060
            Width           =   2295
         End
         Begin VB.TextBox txtPlacedItemsName 
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   8
            Left            =   -74220
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   161
            TabStop         =   0   'False
            Top             =   3360
            Width           =   2295
         End
         Begin VB.TextBox txtPlacedItemsName 
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   9
            Left            =   -74220
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   162
            TabStop         =   0   'False
            Top             =   3660
            Width           =   2295
         End
         Begin VB.CheckBox chkDescription 
            Caption         =   "Check1"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   420
            Width           =   195
         End
         Begin VB.CheckBox chkItemsInRoom 
            Caption         =   "Current Items In Room"
            Height          =   195
            Left            =   -73950
            TabIndex        =   167
            Top             =   720
            Width           =   2115
         End
         Begin VB.TextBox txtVisibleName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   16
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   235
            TabStop         =   0   'False
            Top             =   5940
            Width           =   1815
         End
         Begin VB.TextBox txtVisibleName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   15
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   234
            TabStop         =   0   'False
            Top             =   5625
            Width           =   1815
         End
         Begin VB.TextBox txtVisibleName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   14
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   233
            TabStop         =   0   'False
            Top             =   5325
            Width           =   1815
         End
         Begin VB.TextBox txtVisibleName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   13
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   232
            TabStop         =   0   'False
            Top             =   5010
            Width           =   1815
         End
         Begin VB.TextBox txtVisibleName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   12
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   231
            TabStop         =   0   'False
            Top             =   4710
            Width           =   1815
         End
         Begin VB.TextBox txtVisibleName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   11
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   230
            TabStop         =   0   'False
            Top             =   4395
            Width           =   1815
         End
         Begin VB.TextBox txtVisibleName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   10
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   229
            TabStop         =   0   'False
            Top             =   4095
            Width           =   1815
         End
         Begin VB.TextBox txtVisibleName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   9
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   228
            TabStop         =   0   'False
            Top             =   3780
            Width           =   1815
         End
         Begin VB.TextBox txtVisibleName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   8
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   227
            TabStop         =   0   'False
            Top             =   3480
            Width           =   1815
         End
         Begin VB.TextBox txtVisibleName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   7
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   226
            TabStop         =   0   'False
            Top             =   3180
            Width           =   1815
         End
         Begin VB.TextBox txtVisibleName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   6
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   225
            TabStop         =   0   'False
            Top             =   2865
            Width           =   1815
         End
         Begin VB.TextBox txtVisibleName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   5
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   224
            TabStop         =   0   'False
            Top             =   2550
            Width           =   1815
         End
         Begin VB.TextBox txtVisibleName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   4
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   223
            TabStop         =   0   'False
            Top             =   2250
            Width           =   1815
         End
         Begin VB.TextBox txtVisibleName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   3
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   222
            TabStop         =   0   'False
            Top             =   1935
            Width           =   1815
         End
         Begin VB.TextBox txtVisibleName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   2
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   221
            TabStop         =   0   'False
            Top             =   1635
            Width           =   1815
         End
         Begin VB.TextBox txtVisibleName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   1
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   220
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox txtVisibleName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   0
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   219
            TabStop         =   0   'False
            Top             =   1020
            Width           =   1815
         End
         Begin VB.CheckBox ChkPermNPC 
            Caption         =   "Permanent NPC"
            Height          =   195
            Left            =   -74820
            TabIndex        =   163
            Top             =   4080
            Width           =   1635
         End
         Begin VB.CheckBox chkPlacedItems 
            Caption         =   "Placed Items"
            Height          =   195
            Left            =   -74820
            TabIndex        =   142
            Top             =   720
            Width           =   1275
         End
         Begin VB.CheckBox chkMonies 
            Caption         =   "Visible Coins"
            Height          =   195
            Left            =   -71280
            TabIndex        =   166
            Top             =   720
            Width           =   1275
         End
         Begin VB.CheckBox chkCurrentRoomMons 
            Caption         =   "Current Monsters In Room"
            Height          =   195
            Left            =   -74760
            TabIndex        =   297
            Top             =   840
            Width           =   3435
         End
         Begin VB.TextBox txtCurrentRoomMon 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   -74760
            MaxLength       =   5
            TabIndex        =   299
            Text            =   "0"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtCurrentRoomMon 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   -74760
            MaxLength       =   5
            TabIndex        =   300
            Text            =   "0"
            Top             =   1740
            Width           =   615
         End
         Begin VB.TextBox txtCurrentRoomMon 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   -74760
            MaxLength       =   5
            TabIndex        =   301
            Text            =   "0"
            Top             =   2040
            Width           =   615
         End
         Begin VB.TextBox txtCurrentRoomMon 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   -74760
            MaxLength       =   5
            TabIndex        =   302
            Text            =   "0"
            Top             =   2340
            Width           =   615
         End
         Begin VB.TextBox txtCurrentRoomMon 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   -74760
            MaxLength       =   5
            TabIndex        =   303
            Text            =   "0"
            Top             =   2640
            Width           =   615
         End
         Begin VB.TextBox txtCurrentRoomMon 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   -74760
            MaxLength       =   5
            TabIndex        =   304
            Text            =   "0"
            Top             =   2940
            Width           =   615
         End
         Begin VB.TextBox txtCurrentRoomMon 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   -74760
            MaxLength       =   5
            TabIndex        =   305
            Text            =   "0"
            Top             =   3240
            Width           =   615
         End
         Begin VB.TextBox txtCurrentRoomMon 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   -74760
            MaxLength       =   5
            TabIndex        =   306
            Text            =   "0"
            Top             =   3540
            Width           =   615
         End
         Begin VB.TextBox txtCurrentRoomMon 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   -74760
            MaxLength       =   5
            TabIndex        =   307
            Text            =   "0"
            Top             =   3840
            Width           =   615
         End
         Begin VB.TextBox txtCurrentRoomMon 
            Enabled         =   0   'False
            Height          =   285
            Index           =   10
            Left            =   -74760
            MaxLength       =   5
            TabIndex        =   308
            Text            =   "0"
            Top             =   4140
            Width           =   615
         End
         Begin VB.TextBox txtCurrentRoomMon 
            Enabled         =   0   'False
            Height          =   285
            Index           =   11
            Left            =   -74760
            MaxLength       =   5
            TabIndex        =   309
            Text            =   "0"
            Top             =   4440
            Width           =   615
         End
         Begin VB.TextBox txtCurrentRoomMon 
            Enabled         =   0   'False
            Height          =   285
            Index           =   12
            Left            =   -74760
            MaxLength       =   5
            TabIndex        =   310
            Text            =   "0"
            Top             =   4740
            Width           =   615
         End
         Begin VB.TextBox txtCurrentRoomMon 
            Enabled         =   0   'False
            Height          =   285
            Index           =   13
            Left            =   -74760
            MaxLength       =   5
            TabIndex        =   311
            Text            =   "0"
            Top             =   5040
            Width           =   615
         End
         Begin VB.TextBox txtCurrentRoomMon 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   -74760
            MaxLength       =   5
            TabIndex        =   298
            Text            =   "0"
            Top             =   1140
            Width           =   615
         End
         Begin VB.TextBox txtCurrentRoomMon 
            Enabled         =   0   'False
            Height          =   285
            Index           =   14
            Left            =   -74760
            MaxLength       =   5
            TabIndex        =   312
            Text            =   "0"
            Top             =   5340
            Width           =   615
         End
         Begin VB.TextBox txtHiddenName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   14
            Left            =   -73005
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   266
            TabStop         =   0   'False
            Top             =   5250
            Width           =   1815
         End
         Begin VB.TextBox txtHiddenName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   13
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   265
            TabStop         =   0   'False
            Top             =   4950
            Width           =   1815
         End
         Begin VB.TextBox txtHiddenName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   12
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   264
            TabStop         =   0   'False
            Top             =   4650
            Width           =   1815
         End
         Begin VB.TextBox txtHiddenName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   11
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   263
            TabStop         =   0   'False
            Top             =   4350
            Width           =   1815
         End
         Begin VB.TextBox txtHiddenName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   10
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   262
            TabStop         =   0   'False
            Top             =   4050
            Width           =   1815
         End
         Begin VB.TextBox txtHiddenName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   9
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   261
            TabStop         =   0   'False
            Top             =   3750
            Width           =   1815
         End
         Begin VB.TextBox txtHiddenName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   8
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   260
            TabStop         =   0   'False
            Top             =   3450
            Width           =   1815
         End
         Begin VB.TextBox txtHiddenName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   7
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   259
            TabStop         =   0   'False
            Top             =   3150
            Width           =   1815
         End
         Begin VB.TextBox txtHiddenName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   6
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   258
            TabStop         =   0   'False
            Top             =   2850
            Width           =   1815
         End
         Begin VB.TextBox txtHiddenName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   5
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   257
            TabStop         =   0   'False
            Top             =   2550
            Width           =   1815
         End
         Begin VB.TextBox txtHiddenName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   4
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   256
            TabStop         =   0   'False
            Top             =   2250
            Width           =   1815
         End
         Begin VB.TextBox txtHiddenName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   3
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   255
            TabStop         =   0   'False
            Top             =   1950
            Width           =   1815
         End
         Begin VB.TextBox txtHiddenName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   2
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   254
            TabStop         =   0   'False
            Top             =   1650
            Width           =   1815
         End
         Begin VB.TextBox txtHiddenName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   1
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   253
            TabStop         =   0   'False
            Top             =   1350
            Width           =   1815
         End
         Begin VB.TextBox txtHiddenName 
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   0
            Left            =   -73020
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   252
            TabStop         =   0   'False
            Top             =   1050
            Width           =   1815
         End
         Begin VB.CheckBox chkHiddenItems 
            Caption         =   "Hidden Items In Room"
            Height          =   195
            Left            =   -73920
            TabIndex        =   236
            Top             =   780
            Width           =   2175
         End
         Begin VB.TextBox txtHiddenNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   -73620
            MaxLength       =   5
            TabIndex        =   237
            Text            =   "0"
            Top             =   1050
            Width           =   615
         End
         Begin VB.TextBox txtHiddenNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   -73620
            MaxLength       =   5
            TabIndex        =   238
            Text            =   "0"
            Top             =   1350
            Width           =   615
         End
         Begin VB.TextBox txtHiddenNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   -73620
            MaxLength       =   5
            TabIndex        =   239
            Text            =   "0"
            Top             =   1650
            Width           =   615
         End
         Begin VB.TextBox txtHiddenNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   -73620
            MaxLength       =   5
            TabIndex        =   240
            Text            =   "0"
            Top             =   1950
            Width           =   615
         End
         Begin VB.TextBox txtHiddenNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   -73620
            MaxLength       =   5
            TabIndex        =   241
            Text            =   "0"
            Top             =   2250
            Width           =   615
         End
         Begin VB.TextBox txtHiddenNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   -73620
            MaxLength       =   5
            TabIndex        =   242
            Text            =   "0"
            Top             =   2550
            Width           =   615
         End
         Begin VB.TextBox txtHiddenNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   -73620
            MaxLength       =   5
            TabIndex        =   243
            Text            =   "0"
            Top             =   2850
            Width           =   615
         End
         Begin VB.TextBox txtHiddenNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   -73620
            MaxLength       =   5
            TabIndex        =   244
            Text            =   "0"
            Top             =   3150
            Width           =   615
         End
         Begin VB.TextBox txtHiddenNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   -73620
            MaxLength       =   5
            TabIndex        =   245
            Text            =   "0"
            Top             =   3450
            Width           =   615
         End
         Begin VB.TextBox txtHiddenNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   -73620
            MaxLength       =   5
            TabIndex        =   246
            Text            =   "0"
            Top             =   3750
            Width           =   615
         End
         Begin VB.TextBox txtHiddenNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   14
            Left            =   -73620
            MaxLength       =   5
            TabIndex        =   251
            Text            =   "0"
            Top             =   5250
            Width           =   615
         End
         Begin VB.TextBox txtHiddenNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   13
            Left            =   -73620
            MaxLength       =   5
            TabIndex        =   250
            Text            =   "0"
            Top             =   4950
            Width           =   615
         End
         Begin VB.TextBox txtHiddenNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   12
            Left            =   -73620
            MaxLength       =   5
            TabIndex        =   249
            Text            =   "0"
            Top             =   4650
            Width           =   615
         End
         Begin VB.TextBox txtHiddenNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   11
            Left            =   -73620
            MaxLength       =   5
            TabIndex        =   248
            Text            =   "0"
            Top             =   4350
            Width           =   615
         End
         Begin VB.TextBox txtHiddenNumber 
            Enabled         =   0   'False
            Height          =   285
            Index           =   10
            Left            =   -73620
            MaxLength       =   5
            TabIndex        =   247
            Text            =   "0"
            Top             =   4050
            Width           =   615
         End
         Begin VB.TextBox txtHiddenQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   -70545
            TabIndex        =   282
            Text            =   "0"
            Top             =   1050
            Width           =   555
         End
         Begin VB.TextBox txtHiddenQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   -70545
            TabIndex        =   283
            Text            =   "0"
            Top             =   1350
            Width           =   555
         End
         Begin VB.TextBox txtHiddenQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   -70545
            TabIndex        =   284
            Text            =   "0"
            Top             =   1650
            Width           =   555
         End
         Begin VB.TextBox txtHiddenQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   -70545
            TabIndex        =   285
            Text            =   "0"
            Top             =   1950
            Width           =   555
         End
         Begin VB.TextBox txtHiddenQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   -70545
            TabIndex        =   286
            Text            =   "0"
            Top             =   2250
            Width           =   555
         End
         Begin VB.TextBox txtHiddenQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   -70545
            TabIndex        =   287
            Text            =   "0"
            Top             =   2550
            Width           =   555
         End
         Begin VB.TextBox txtHiddenQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   -70545
            TabIndex        =   288
            Text            =   "0"
            Top             =   2850
            Width           =   555
         End
         Begin VB.TextBox txtHiddenQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   -70545
            TabIndex        =   289
            Text            =   "0"
            Top             =   3150
            Width           =   555
         End
         Begin VB.TextBox txtHiddenQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   -70545
            TabIndex        =   290
            Text            =   "0"
            Top             =   3450
            Width           =   555
         End
         Begin VB.TextBox txtHiddenQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   -70545
            TabIndex        =   291
            Text            =   "0"
            Top             =   3750
            Width           =   555
         End
         Begin VB.TextBox txtHiddenQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   10
            Left            =   -70545
            TabIndex        =   292
            Text            =   "0"
            Top             =   4050
            Width           =   555
         End
         Begin VB.TextBox txtHiddenQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   11
            Left            =   -70545
            TabIndex        =   293
            Text            =   "0"
            Top             =   4350
            Width           =   555
         End
         Begin VB.TextBox txtHiddenQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   12
            Left            =   -70545
            TabIndex        =   294
            Text            =   "0"
            Top             =   4650
            Width           =   555
         End
         Begin VB.TextBox txtHiddenQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   13
            Left            =   -70545
            TabIndex        =   295
            Text            =   "0"
            Top             =   4950
            Width           =   555
         End
         Begin VB.TextBox txtHiddenQty 
            Enabled         =   0   'False
            Height          =   285
            Index           =   14
            Left            =   -70545
            TabIndex        =   296
            Text            =   "0"
            Top             =   5250
            Width           =   555
         End
         Begin VB.TextBox txtVisibleUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   16
            Left            =   -71190
            TabIndex        =   201
            Text            =   "0"
            Top             =   5940
            Width           =   615
         End
         Begin VB.TextBox txtVisibleUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   15
            Left            =   -71190
            TabIndex        =   200
            Text            =   "0"
            Top             =   5625
            Width           =   615
         End
         Begin VB.TextBox txtVisibleUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   14
            Left            =   -71190
            TabIndex        =   199
            Text            =   "0"
            Top             =   5325
            Width           =   615
         End
         Begin VB.TextBox txtVisibleUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   13
            Left            =   -71190
            TabIndex        =   198
            Text            =   "0"
            Top             =   5010
            Width           =   615
         End
         Begin VB.TextBox txtVisibleUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   12
            Left            =   -71190
            TabIndex        =   197
            Text            =   "0"
            Top             =   4710
            Width           =   615
         End
         Begin VB.TextBox txtVisibleUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   11
            Left            =   -71190
            TabIndex        =   196
            Text            =   "0"
            Top             =   4395
            Width           =   615
         End
         Begin VB.TextBox txtVisibleUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   10
            Left            =   -71190
            TabIndex        =   195
            Text            =   "0"
            Top             =   4095
            Width           =   615
         End
         Begin VB.TextBox txtVisibleUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   -71190
            TabIndex        =   194
            Text            =   "0"
            Top             =   3780
            Width           =   615
         End
         Begin VB.TextBox txtVisibleUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   -71190
            TabIndex        =   193
            Text            =   "0"
            Top             =   3480
            Width           =   615
         End
         Begin VB.TextBox txtVisibleUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   -71190
            TabIndex        =   192
            Text            =   "0"
            Top             =   3165
            Width           =   615
         End
         Begin VB.TextBox txtVisibleUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   -71190
            TabIndex        =   191
            Text            =   "0"
            Top             =   2865
            Width           =   615
         End
         Begin VB.TextBox txtVisibleUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   -71205
            TabIndex        =   190
            Text            =   "0"
            Top             =   2550
            Width           =   615
         End
         Begin VB.TextBox txtVisibleUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   -71190
            TabIndex        =   189
            Text            =   "0"
            Top             =   2250
            Width           =   615
         End
         Begin VB.TextBox txtVisibleUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   -71190
            TabIndex        =   188
            Text            =   "0"
            Top             =   1935
            Width           =   615
         End
         Begin VB.TextBox txtVisibleUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   -71190
            TabIndex        =   187
            Text            =   "0"
            Top             =   1635
            Width           =   615
         End
         Begin VB.TextBox txtVisibleUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   -71190
            TabIndex        =   186
            Text            =   "0"
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox txtVisibleUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   -71190
            TabIndex        =   185
            Text            =   "0"
            Top             =   1020
            Width           =   615
         End
         Begin VB.TextBox txtHiddenUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   14
            Left            =   -71190
            TabIndex        =   281
            Text            =   "0"
            Top             =   5250
            Width           =   615
         End
         Begin VB.TextBox txtHiddenUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   13
            Left            =   -71190
            TabIndex        =   280
            Text            =   "0"
            Top             =   4950
            Width           =   615
         End
         Begin VB.TextBox txtHiddenUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   12
            Left            =   -71190
            TabIndex        =   279
            Text            =   "0"
            Top             =   4650
            Width           =   615
         End
         Begin VB.TextBox txtHiddenUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   11
            Left            =   -71190
            TabIndex        =   278
            Text            =   "0"
            Top             =   4350
            Width           =   615
         End
         Begin VB.TextBox txtHiddenUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   10
            Left            =   -71190
            TabIndex        =   277
            Text            =   "0"
            Top             =   4050
            Width           =   615
         End
         Begin VB.TextBox txtHiddenUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   -71190
            TabIndex        =   276
            Text            =   "0"
            Top             =   3750
            Width           =   615
         End
         Begin VB.TextBox txtHiddenUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   -71190
            TabIndex        =   275
            Text            =   "0"
            Top             =   3450
            Width           =   615
         End
         Begin VB.TextBox txtHiddenUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   -71190
            TabIndex        =   274
            Text            =   "0"
            Top             =   3150
            Width           =   615
         End
         Begin VB.TextBox txtHiddenUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   -71190
            TabIndex        =   273
            Text            =   "0"
            Top             =   2850
            Width           =   615
         End
         Begin VB.TextBox txtHiddenUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   -71190
            TabIndex        =   272
            Text            =   "0"
            Top             =   2550
            Width           =   615
         End
         Begin VB.TextBox txtHiddenUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   -71190
            TabIndex        =   271
            Text            =   "0"
            Top             =   2250
            Width           =   615
         End
         Begin VB.TextBox txtHiddenUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   -71190
            TabIndex        =   270
            Text            =   "0"
            Top             =   1950
            Width           =   615
         End
         Begin VB.TextBox txtHiddenUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   -71190
            TabIndex        =   269
            Text            =   "0"
            Top             =   1650
            Width           =   615
         End
         Begin VB.TextBox txtHiddenUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   -71190
            TabIndex        =   268
            Text            =   "0"
            Top             =   1350
            Width           =   615
         End
         Begin VB.TextBox txtHiddenUses 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   -71190
            TabIndex        =   267
            Text            =   "0"
            Top             =   1050
            Width           =   615
         End
         Begin VB.TextBox txtMonNote 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   2295
            Left            =   -73920
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   313
            Text            =   "frmMassRoomEditor.frx":16EF
            Top             =   1620
            Width           =   4935
         End
         Begin VB.Frame frmVisibleCoins 
            Caption         =   "        "
            Height          =   2295
            Left            =   -71520
            TabIndex        =   322
            Top             =   720
            Width           =   2175
            Begin VB.TextBox txtCopper 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   327
               Text            =   "0"
               Top             =   1800
               Width           =   735
            End
            Begin VB.TextBox txtSilver 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   326
               Text            =   "0"
               Top             =   1440
               Width           =   735
            End
            Begin VB.TextBox txtGold 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   325
               Text            =   "0"
               Top             =   1080
               Width           =   735
            End
            Begin VB.TextBox txtPlatinum 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   324
               Text            =   "0"
               Top             =   720
               Width           =   735
            End
            Begin VB.TextBox txtRunic 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   323
               Text            =   "0"
               Top             =   360
               Width           =   735
            End
            Begin VB.Label label 
               AutoSize        =   -1  'True
               Caption         =   "Copper"
               Height          =   195
               Index           =   40
               Left            =   240
               TabIndex        =   332
               Top             =   1800
               Width           =   510
            End
            Begin VB.Label label 
               AutoSize        =   -1  'True
               Caption         =   "Silver"
               Height          =   195
               Index           =   39
               Left            =   240
               TabIndex        =   331
               Top             =   1440
               Width           =   390
            End
            Begin VB.Label label 
               AutoSize        =   -1  'True
               Caption         =   "Gold"
               Height          =   195
               Index           =   38
               Left            =   240
               TabIndex        =   330
               Top             =   1080
               Width           =   330
            End
            Begin VB.Label label 
               AutoSize        =   -1  'True
               Caption         =   "Platinum"
               Height          =   195
               Index           =   37
               Left            =   240
               TabIndex        =   329
               Top             =   720
               Width           =   600
            End
            Begin VB.Label label 
               AutoSize        =   -1  'True
               Caption         =   "Runic"
               Height          =   195
               Index           =   44
               Left            =   240
               TabIndex        =   328
               Top             =   360
               Width           =   420
            End
         End
         Begin VB.Label label 
            Alignment       =   2  'Center
            Caption         =   "Para4"
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
            Index           =   19
            Left            =   -69675
            TabIndex        =   71
            Top             =   600
            Width           =   570
         End
         Begin VB.Label label 
            Alignment       =   2  'Center
            Caption         =   "Para3"
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
            Index           =   18
            Left            =   -70395
            TabIndex        =   70
            Top             =   600
            Width           =   570
         End
         Begin VB.Label label 
            Alignment       =   2  'Center
            Caption         =   "Para2"
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
            Index           =   17
            Left            =   -71115
            TabIndex        =   69
            Top             =   600
            Width           =   570
         End
         Begin VB.Label Para1 
            Alignment       =   2  'Center
            Caption         =   "Para1"
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
            Index           =   16
            Left            =   -71865
            TabIndex        =   68
            Top             =   600
            Width           =   570
         End
         Begin VB.Label label 
            Alignment       =   2  'Center
            Caption         =   "Room #"
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
            Index           =   14
            Left            =   -74100
            TabIndex        =   66
            Top             =   600
            Width           =   570
         End
         Begin VB.Label label 
            Alignment       =   2  'Center
            Caption         =   "Exit Type"
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
            Index           =   15
            Left            =   -73440
            TabIndex        =   67
            Top             =   600
            Width           =   1410
         End
         Begin VB.Label label 
            Caption         =   "Current Items In Room"
            Height          =   195
            Index           =   35
            Left            =   -74880
            TabIndex        =   321
            Top             =   480
            Width           =   2040
         End
         Begin VB.Label Label12 
            Caption         =   "Qty - 1"
            Height          =   255
            Left            =   -70560
            TabIndex        =   320
            Top             =   780
            Width           =   615
         End
         Begin VB.Label Label13 
            Caption         =   "Qty - 1"
            Height          =   255
            Left            =   -70530
            TabIndex        =   319
            Top             =   810
            Width           =   495
         End
         Begin VB.Label Label11 
            Caption         =   "The quantity totals here are the total minus one.  (EX: enter 0 for 1 ... 1 for 2 ... 3 for 4, etc.)"
            Height          =   615
            Left            =   -73605
            TabIndex        =   318
            Top             =   5670
            Width           =   2895
         End
         Begin VB.Label label 
            Alignment       =   2  'Center
            Caption         =   "Uses"
            Height          =   255
            Index           =   49
            Left            =   -71190
            TabIndex        =   317
            Top             =   780
            Width           =   615
         End
         Begin VB.Label label 
            Alignment       =   2  'Center
            Caption         =   "Uses"
            Height          =   255
            Index           =   50
            Left            =   -71190
            TabIndex        =   316
            Top             =   810
            Width           =   615
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   255
         Left            =   60
         TabIndex        =   314
         Top             =   7140
         Visible         =   0   'False
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Room"
         Height          =   195
         Left            =   6420
         TabIndex        =   4
         Top             =   2220
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Room"
         Height          =   195
         Left            =   6420
         TabIndex        =   7
         Top             =   3540
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Map"
         Height          =   195
         Left            =   6420
         TabIndex        =   2
         Top             =   1620
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6420
         TabIndex        =   6
         Top             =   3120
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6420
         TabIndex        =   1
         Top             =   1200
         Width           =   840
      End
   End
   Begin MSComctlLib.StatusBar stsStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   315
      Top             =   7485
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10425
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMassRoomEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim bStopProcess As Boolean
Public bErrors As Boolean
Public UpdateSuccess As Boolean

Private Sub chkAnsiMap_Click()
If chkAnsiMap.Value = 0 Then txtAnsiMap.Enabled = False
If chkAnsiMap.Value = 1 Then txtAnsiMap.Enabled = True
End Sub

Private Sub chkAttributes_Click()
If chkAttributes.Value = 0 Then txtAttributes.Enabled = False
If chkAttributes.Value = 1 Then txtAttributes.Enabled = True
End Sub

Private Sub chkCmdText_Click()
If chkCmdText.Value = 0 Then txtCmdText.Enabled = False
If chkCmdText.Value = 1 Then txtCmdText.Enabled = True
End Sub

Private Sub chkControlRoom_Click()
If chkControlRoom.Value = 0 Then txtControlRoom.Enabled = False
If chkControlRoom.Value = 1 Then txtControlRoom.Enabled = True
End Sub

Private Sub chkCurrentRoomMons_Click()
Dim x As Integer

If chkCurrentRoomMons = 0 Then
    For x = 0 To 14
        txtCurrentRoomMon(x).Enabled = False
    Next
End If
If chkCurrentRoomMons = 1 Then
    For x = 0 To 14
        txtCurrentRoomMon(x).Enabled = True
    Next
End If

End Sub

Private Sub chkDeathRoom_Click()
If chkDeathRoom.Value = 0 Then txtDeathRoom.Enabled = False
If chkDeathRoom.Value = 1 Then txtDeathRoom.Enabled = True
End Sub

Private Sub chkDelay_Click()
If chkDelay.Value = 0 Then txtDelay.Enabled = False
If chkDelay.Value = 1 Then txtDelay.Enabled = True
End Sub

Private Sub chkDescription_Click()
If chkDescription.Value = 0 Then frmDescription.Enabled = False
If chkDescription.Value = 1 Then frmDescription.Enabled = True
End Sub

Private Sub chkExitRoom_Click()
If chkExitRoom.Value = 0 Then txtExitRoom.Enabled = False
If chkExitRoom.Value = 1 Then txtExitRoom.Enabled = True
End Sub

Private Sub chkExits_Click(Index As Integer)
If chkExits(Index).Value = 0 Then
    txtRoomExit(Index).Enabled = False
    cmbRoomType(Index).Enabled = False
    txtRoomPara(Index).Enabled = False
    txtRoomWPara(Index).Enabled = False
    txtRoomLPara1(Index).Enabled = False
    txtRoomLPara2(Index).Enabled = False
End If

If chkExits(Index).Value = 1 Then
    txtRoomExit(Index).Enabled = True
    cmbRoomType(Index).Enabled = True
    txtRoomPara(Index).Enabled = True
    txtRoomWPara(Index).Enabled = True
    txtRoomLPara1(Index).Enabled = True
    txtRoomLPara2(Index).Enabled = True
End If

End Sub

Private Sub chkGangHouseNumber_Click()
If chkGangHouseNumber.Value = 0 Then txtGangHouseNumber.Enabled = False
If chkGangHouseNumber.Value = 1 Then txtGangHouseNumber.Enabled = True
End Sub

Private Sub chkHiddenItems_Click()
Dim x As Integer

If chkHiddenItems = 0 Then
    For x = 0 To 14
        txtHiddenNumber(x).Enabled = False
        txtHiddenUses(x).Enabled = False
        txtHiddenQty(x).Enabled = False
    Next
End If
If chkHiddenItems = 1 Then
    For x = 0 To 14
        txtHiddenNumber(x).Enabled = True
        txtHiddenUses(x).Enabled = True
        txtHiddenQty(x).Enabled = True
    Next
End If

End Sub

Private Sub chkItemsInRoom_Click()
Dim x As Integer

If chkItemsInRoom = 0 Then
    For x = 0 To 16
        txtVisibleNumber(x).Enabled = False
        txtVisibleUses(x).Enabled = False
        txtVisibleQty(x).Enabled = False
    Next
End If
If chkItemsInRoom = 1 Then
    For x = 0 To 16
        txtVisibleNumber(x).Enabled = True
        txtVisibleUses(x).Enabled = True
        txtVisibleQty(x).Enabled = True
    Next
End If

End Sub

Private Sub chkLight_Click()
If chkLight.Value = 0 Then txtLight.Enabled = False
If chkLight.Value = 1 Then txtLight.Enabled = True
End Sub

Private Sub chkMaxArea_Click()
If chkMaxArea.Value = 0 Then txtMaxArea.Enabled = False
If chkMaxArea.Value = 1 Then txtMaxArea.Enabled = True
End Sub

Private Sub chkMaxIndex_Click()
If chkMaxIndex.Value = 0 Then txtMaxIndex.Enabled = False
If chkMaxIndex.Value = 1 Then txtMaxIndex.Enabled = True
End Sub

Private Sub chkMaxRegen_Click()
If chkMaxRegen.Value = 0 Then txtMaxRegen.Enabled = False
If chkMaxRegen.Value = 1 Then txtMaxRegen.Enabled = True
End Sub

Private Sub chkMinIndex_Click()
If chkMinIndex.Value = 0 Then txtMinIndex.Enabled = False
If chkMinIndex.Value = 1 Then txtMinIndex.Enabled = True
End Sub

Private Sub chkMonies_Click()
If chkMonies.Value = 0 Then
    txtRunic.Enabled = False
    txtPlatinum.Enabled = False
    txtGold.Enabled = False
    txtSilver.Enabled = False
    txtCopper.Enabled = False
End If
If chkMonies.Value = 1 Then
    txtRunic.Enabled = True
    txtPlatinum.Enabled = True
    txtGold.Enabled = True
    txtSilver.Enabled = True
    txtCopper.Enabled = True
End If

End Sub

Private Sub chkInvisMonies_Click()
If chkMonies.Value = 0 Then
    txtInvisRunic.Enabled = False
    txtInvisPlatinum.Enabled = False
    txtInvisGold.Enabled = False
    txtInvisSilver.Enabled = False
    txtInvisCopper.Enabled = False
End If
If chkInvisMonies.Value = 1 Then
    txtInvisRunic.Enabled = True
    txtInvisPlatinum.Enabled = True
    txtInvisGold.Enabled = True
    txtInvisSilver.Enabled = True
    txtInvisCopper.Enabled = True
End If

End Sub

Private Sub chkMonsterType_Click()
If chkMonsterType.Value = 0 Then cmbMonsterType.Enabled = False
If chkMonsterType.Value = 1 Then cmbMonsterType.Enabled = True
End Sub

Private Sub chkName_Click()
If chkName.Value = 0 Then txtName.Enabled = False
If chkName.Value = 1 Then txtName.Enabled = True
End Sub

Private Sub ChkPermNPC_Click()
If ChkPermNPC = 0 Then txtPermNPC.Enabled = False
If ChkPermNPC = 1 Then txtPermNPC.Enabled = True
End Sub

Private Sub chkPlacedItems_Click()
Dim x As Integer

If chkPlacedItems = 0 Then
    For x = 0 To 9
        txtPlacedItems(x).Enabled = False
    Next
End If
If chkPlacedItems = 1 Then
    For x = 0 To 9
        txtPlacedItems(x).Enabled = True
    Next
End If
End Sub

Private Sub chkRoomType_Click()
If chkRoomType.Value = 0 Then cmbType.Enabled = False
If chkRoomType.Value = 1 Then cmbType.Enabled = True
End Sub

Private Sub chkShop_Click()
If chkShop.Value = 0 Then txtShopNum.Enabled = False
If chkShop.Value = 1 Then txtShopNum.Enabled = True
End Sub

Private Sub chkSpell_Click()
If chkSpell.Value = 0 Then txtSpell.Enabled = False
If chkSpell.Value = 1 Then txtSpell.Enabled = True
End Sub

Private Sub cmdClose_Click()
Dim nYesNo As Integer
If cmdGo.Enabled = False Then
    nYesNo = MsgBox("Are you sure you want to cancel?", vbYesNo + vbQuestion + vbDefaultButton2)
    If Not nYesNo = vbYes Then Exit Sub

    cmdClose.Enabled = False
    bStopProcess = True
Else
    Unload Me
End If
End Sub

Private Sub cmdCopyDesc_Click(Index As Integer)
On Error GoTo error:
Dim x As Integer

If Index = 0 Then
    For x = 0 To 6
        sRoomCopyDesc(x) = txtDesc(x).Text
    Next
ElseIf Index = 1 Then
    For x = 0 To 6
        txtDesc(x).Text = sRoomCopyDesc(x)
    Next
End If

Exit Sub
error:
Call HandleError
End Sub

Private Sub cmdCopyPaste_Click(Index As Integer)

If Index = 0 Then
    nRoomCopyPaste(0) = Val(txtLight.Text)
    nRoomCopyPaste(1) = Val(txtGangHouseNumber.Text)
    nRoomCopyPaste(2) = Val(txtAttributes.Text)
    nRoomCopyPaste(3) = Val(txtShopNum.Text)
    nRoomCopyPaste(4) = Val(txtSpell.Text)
    nRoomCopyPaste(5) = Val(txtDeathRoom.Text)
    nRoomCopyPaste(6) = Val(txtExitRoom.Text)
    nRoomCopyPaste(7) = Val(txtCmdText.Text)
    nRoomCopyPaste(8) = Val(txtMinIndex.Text)
    nRoomCopyPaste(9) = Val(txtMaxIndex.Text)
    nRoomCopyPaste(10) = Val(txtMaxRegen.Text)
    nRoomCopyPaste(11) = Val(txtDelay.Text)
    nRoomCopyPaste(12) = Val(txtControlRoom.Text)
    nRoomCopyPaste(13) = Val(txtMaxArea.Text)
    nRoomCopyPaste(14) = Val(cmbMonsterType.ListIndex)
    nRoomCopyPaste(15) = Val(cmbType.ListIndex)
    sRoomCopyPaste = txtAnsiMap.Text
ElseIf Index = 1 Then
    txtLight.Text = nRoomCopyPaste(0)
    txtGangHouseNumber.Text = nRoomCopyPaste(1)
    txtAttributes.Text = nRoomCopyPaste(2)
    txtShopNum.Text = nRoomCopyPaste(3)
    txtSpell.Text = nRoomCopyPaste(4)
    txtDeathRoom.Text = nRoomCopyPaste(5)
    txtExitRoom.Text = nRoomCopyPaste(6)
    txtCmdText.Text = nRoomCopyPaste(7)
    txtMinIndex.Text = nRoomCopyPaste(8)
    txtMaxIndex.Text = nRoomCopyPaste(9)
    txtMaxRegen.Text = nRoomCopyPaste(10)
    txtDelay.Text = nRoomCopyPaste(11)
    txtControlRoom.Text = nRoomCopyPaste(12)
    txtMaxArea.Text = nRoomCopyPaste(13)
    cmbMonsterType.ListIndex = nRoomCopyPaste(14)
    cmbType.ListIndex = nRoomCopyPaste(15)
    txtAnsiMap.Text = sRoomCopyPaste
End If

End Sub

Private Sub cmdDeleteRange_Click()
On Error GoTo error:
Dim nStatus As Integer, x As Integer, y As Integer

x = MsgBox("Are you sure you want to DELETE rooms " & Val(txtStartRoom.Text) & " to " & Val(txtEndRoom.Text) & " of map " & Val(txtStartMap.Text) & "?", vbYesNo + vbDefaultButton2)
If x <> 6 Then Exit Sub

If Val(txtStartRoom.Text) > Val(txtEndRoom.Text) Then
    MsgBox "Illegal Range Entered!", vbExclamation
    Exit Sub
End If

UnloadForms (Me.Name)

RoomKeyStruct.MapNum = Val(txtStartMap.Text)
RoomKeyStruct.RoomNum = Val(txtStartRoom.Text)

nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Couldn't get first room, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

stsStatus.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_MP
stsStatus.Panels(2).Text = RoomKeyStruct.RoomNum
ProgressBar.Value = 0
ProgressBar.Max = Val(txtEndRoom.Text) - Val(txtStartRoom.Text) + 2
ProgressBar.Visible = True

cmdGo.Enabled = False
cmdClose.Enabled = False
frmMain.Enabled = False

For x = Val(txtStartRoom.Text) + 1 To Val(txtEndRoom.Text)
    If nStatus = 0 Then
        nStatus = BTRCALL(BDELETE, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    End If
    
    RoomKeyStruct.RoomNum = x

    nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), RoomKeyStruct, KEY_BUF_LEN, 0)
    
    ProgressBar.Value = ProgressBar.Value + 1
    stsStatus.Panels(2).Text = RoomKeyStruct.RoomNum
    If Not bUseCPU = True Then DoEvents
Next

ProgressBar.Value = ProgressBar.Max

ProgressBar.Visible = False
cmdGo.Enabled = True
cmdClose.Enabled = True
frmMain.Enabled = True

MsgBox "Complete!", vbInformation

ProgressBar.Visible = False

Exit Sub
error:
Call HandleError
ProgressBar.Visible = False
cmdGo.Enabled = True
cmdClose.Enabled = True
frmMain.Enabled = True
End Sub

Private Sub cmdGo_Click()
On Error GoTo error:
Dim nStatus As Integer, x As Integer, y As Integer
Dim RoomStart As Long, RoomEnd As Long, CurrentRoom As Long
Dim fso As FileSystemObject, fil As String, ts As TextStream, frm As Form

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

x = MsgBox("Are you sure you want to continue?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Action")
If Not x = vbYes Then Exit Sub

If Val(txtStartRoom.Text) > Val(txtEndRoom.Text) Then
    MsgBox "Illegal Range Entered!", vbExclamation
    Exit Sub
End If

RoomStart = Val(txtStartRoom.Text)
RoomEnd = Val(txtEndRoom.Text)
RoomKeyStruct.MapNum = Val(txtStartMap.Text)
RoomKeyStruct.RoomNum = RoomStart
CurrentRoom = RoomStart

nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Couldn't get first room, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

Set fso = CreateObject("Scripting.FileSystemObject")
If Right(App.Path, 1) = "\" Then
    fil = App.Path & "NMR-Log_MassRoom.txt"
Else
    fil = App.Path & "\NMR-Log_MassRoom.txt"
End If
If fso.FileExists(fil) = True Then fso.DeleteFile fil, True
Set ts = fso.OpenTextFile(fil, ForWriting, True)

ts.WriteLine ("Range edit job started " & Date & " @ " & Time)
ts.WriteLine ("Editing Rooms " & RoomStart & " to " & RoomEnd & " of Map " & txtStartMap.Text)
ts.WriteBlankLines (1)

stsStatus.Panels(1).Text = "w" & strDatCallLetters & strDatSuffix_MP

For Each frm In Forms
    If LCase(Left(frm.Name, Len("room editor"))) = "room editor" Then
        Unload frm
    ElseIf Not frm Is Me And Not frm Is frmMain Then
        frm.WindowState = vbMinimized
        frm.Enabled = False
        If frmMain.tbTaskBar.Visible Then frm.Hide
    End If
    Set frm = Nothing
Next

ProgressBar.Value = 0
ProgressBar.Visible = True
ProgressBar.Max = RoomEnd - RoomStart + 2

bStopProcess = False
bErrors = False
cmdGo.Enabled = False
fraMain.Enabled = False
Call LockMenus

DoEvents
Do While CurrentRoom <= RoomEnd And bStopProcess = False
    stsStatus.Panels(2).Text = CurrentRoom
    RoomKeyStruct.RoomNum = CurrentRoom
    
    nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), RoomKeyStruct, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        bErrors = True
        ts.WriteLine ("Room: " & CurrentRoom & " - GET Error: " & nStatus)
        GoTo GotoNextRoom:
    End If
    
    RoomRowToStruct Roomdatabuf.buf
    
    If chkName.Value = 1 Then Roomrec.Name = RTrim(txtName.Text) & Chr(0)
    If chkAnsiMap.Value = 1 Then Roomrec.AnsiMap = RTrim(txtAnsiMap.Text) & Chr(0)
    If chkRoomType.Value = 1 Then Roomrec.Type = cmbType.ListIndex
    If chkShop.Value = 1 Then Roomrec.ShopNum = Val(txtShopNum.Text)
    If chkMinIndex.Value = 1 Then Roomrec.MinIndex = Val(txtMinIndex.Text)
    If chkMaxIndex.Value = 1 Then Roomrec.MaxIndex = Val(txtMaxIndex.Text)
    If ChkPermNPC.Value = 1 Then Roomrec.PermNPC = Val(txtPermNPC.Text)
    If chkLight.Value = 1 Then Roomrec.Light = Val(txtLight.Text)
    If chkMonsterType.Value = 1 Then Roomrec.MonsterType = cmbMonsterType.ListIndex
    If chkMaxRegen.Value = 1 Then Roomrec.MaxRegen = Val(txtMaxRegen.Text)
    If chkDeathRoom.Value = 1 Then Roomrec.DeathRoom = Val(txtDeathRoom.Text)
    If chkCmdText.Value = 1 Then Roomrec.CmdText = Val(txtCmdText.Text)
    If chkDelay.Value = 1 Then Roomrec.Delay = Val(txtDelay.Text)
    If chkMaxArea.Value = 1 Then Roomrec.MaxArea = Val(txtMaxArea.Text)
    If chkControlRoom.Value = 1 Then Roomrec.ControlRoom = Val(txtControlRoom.Text)
    If chkGangHouseNumber.Value = 1 Then Roomrec.GangHouseNumber = Val(txtGangHouseNumber.Text)
    If chkMonies.Value = 1 Then
        Roomrec.Runic = Val(txtRunic.Text)
        Roomrec.Platinum = Val(txtPlatinum.Text)
        Roomrec.Gold = Val(txtGold.Text)
        Roomrec.Silver = Val(txtSilver.Text)
        Roomrec.Copper = Val(txtCopper.Text)
    End If
    If chkInvisMonies.Value = 1 Then
        Roomrec.InvisRunic = Val(txtInvisRunic.Text)
        Roomrec.InvisPlatinum = Val(txtInvisPlatinum.Text)
        Roomrec.InvisGold = Val(txtInvisGold.Text)
        Roomrec.InvisSilver = Val(txtInvisSilver.Text)
        Roomrec.InvisCopper = Val(txtInvisCopper.Text)
    End If
    If chkSpell.Value = 1 Then Roomrec.Spell = Val(txtSpell.Text)
    If chkExitRoom.Value = 1 Then Roomrec.ExitRoom = Val(txtExitRoom.Text)
    If chkAttributes.Value = 1 Then Roomrec.Attributes = Val(txtAttributes.Text)
    
    If chkDescription.Value = 1 Then
        For x = 0 To 6
            Roomrec.Desc(x) = RTrim(txtDesc(x).Text) & Chr(0)
        Next x
    End If
    
    If chkItemsInRoom.Value = 1 Then
        For x = 0 To 16
            Roomrec.RoomItems(x) = Val(txtVisibleNumber(x).Text)
            Roomrec.RoomItemUses(x) = Val(txtVisibleUses(x).Text)
            Roomrec.RoomItemQty(x) = Val(txtVisibleQty(x).Text)
        Next x
    End If
    
    If chkHiddenItems.Value = 1 Then
        For x = 0 To 14
            Roomrec.InvisItems(x) = Val(txtHiddenNumber(x).Text)
            Roomrec.InvisItemUses(x) = Val(txtHiddenUses(x).Text)
            Roomrec.InvisItemQty(x) = Val(txtHiddenQty(x).Text)
        Next x
    End If
    
    If chkCurrentRoomMons.Value = 1 Then
        For x = 0 To 14
            Roomrec.CurrentRoomMon(x) = Val(txtCurrentRoomMon(x).Text)
        Next x
    End If
    
    For y = 0 To 9
        If chkExits(y).Value = 1 Then
            For x = 0 To 9
                Roomrec.RoomExit(x) = Val(txtRoomExit(x).Text)
                Roomrec.RoomType(x) = cmbRoomType(x).ListIndex
                Roomrec.Para1(x) = Val(txtRoomPara(x).Text)
                Roomrec.Para2(x) = Val(txtRoomWPara(x).Text)
                Roomrec.Para3(x) = Val(txtRoomLPara1(x).Text)
                Roomrec.Para4(x) = Val(txtRoomLPara2(x).Text)
            Next x
        End If
    Next y
    
    If chkPlacedItems.Value = 1 Then
        For x = 0 To 9
            Roomrec.PlacedItems(x) = Val(txtPlacedItems(x).Text)
        Next x
    End If
    
    nStatus = UpdateRoom
    If Not nStatus = 0 Then ts.WriteLine ("Room: " & CurrentRoom & " - Update Error: " & nStatus)
    
GotoNextRoom:
    CurrentRoom = CurrentRoom + 1
    ProgressBar.Value = ProgressBar.Value + 1
    If Not bUseCPU = True Then DoEvents
Loop

If bStopProcess Then
    ts.WriteBlankLines (1)
    ts.WriteLine "...canceled by user."
    GoTo ReEnable:
End If

ProgressBar.Value = ProgressBar.Max

ts.WriteBlankLines (1)
ts.WriteLine ("Complete - " & Date & " @ " & Time)
ts.Close

If Not bErrors Then
    MsgBox "Complete!", vbInformation
Else
    x = MsgBox("Complete but had errors, view log?", vbYesNo + vbQuestion + vbDefaultButton1)
    If x = vbYes Then Call cmdLog_Click
End If

ReEnable:
On Error Resume Next

For Each frm In Forms
    If Not frm Is Me And Not frm Is frmMain Then
        frm.Enabled = True
    End If
Next

ts.Close
cmdGo.Enabled = True
cmdClose.Enabled = True
fraMain.Enabled = True
ProgressBar.Visible = False
Call UnLockMenus
stsStatus.Panels(1).Text = ""
stsStatus.Panels(2).Text = ""
Set fso = Nothing
Set ts = Nothing
Set frm = Nothing

Exit Sub
error:
Call HandleError
Resume ReEnable:
End Sub

Private Sub cmdLog_Click()
Dim fso As FileSystemObject, fil As String

Set fso = CreateObject("Scripting.FileSystemObject")

If Right(App.Path, 1) = "\" Then
    fil = App.Path & "NMR-Log_MassRoom.txt"
Else
    fil = App.Path & "\NMR-Log_MassRoom.txt"
End If

If fso.FileExists(fil) = True Then
    Call ShellExecute(0&, "open", fil, vbNullString, vbNullString, vbNormalFocus)
Else
    MsgBox fil & " was not found.", vbInformation
End If

Set fso = Nothing
End Sub

Private Sub cmdSelectAll_Click()
Dim i As Integer

    chkName.Value = 1
    chkAnsiMap.Value = 1
    chkRoomType.Value = 1
    chkShop.Value = 1
    chkMinIndex.Value = 1
    chkMaxIndex.Value = 1
    ChkPermNPC.Value = 1
    chkLight.Value = 1
    chkMonsterType.Value = 1
    chkMaxRegen.Value = 1
    chkDeathRoom.Value = 1
    chkCmdText.Value = 1
    chkDelay.Value = 1
    chkMaxArea.Value = 1
    chkControlRoom.Value = 1
    chkMonies.Value = 1
    chkMonies.Value = 1
    chkMonies.Value = 1
    chkMonies.Value = 1
    chkMonies.Value = 1
    chkSpell.Value = 1
    chkExitRoom.Value = 1
    chkAttributes.Value = 1
    chkDescription.Value = 1
    chkItemsInRoom.Value = 1
    chkHiddenItems.Value = 1
    chkCurrentRoomMons.Value = 1
    chkGangHouseNumber.Value = 1
    For i = 0 To 9
        chkExits(i).Value = 1
    Next
    chkPlacedItems.Value = 1

End Sub

Private Sub cmdSelectNone_Click()
Dim i As Integer
    chkName.Value = 0
    chkAnsiMap.Value = 0
    chkRoomType.Value = 0
    chkShop.Value = 0
    chkMinIndex.Value = 0
    chkMaxIndex.Value = 0
    ChkPermNPC.Value = 0
    chkLight.Value = 0
    chkMonsterType.Value = 0
    chkMaxRegen.Value = 0
    chkDeathRoom.Value = 0
    chkCmdText.Value = 0
    chkDelay.Value = 0
    chkMaxArea.Value = 0
    chkControlRoom.Value = 0
    chkMonies.Value = 0
    chkMonies.Value = 0
    chkMonies.Value = 0
    chkMonies.Value = 0
    chkMonies.Value = 0
    chkSpell.Value = 0
    chkExitRoom.Value = 0
    chkAttributes.Value = 0
    chkDescription.Value = 0
    chkItemsInRoom.Value = 0
    chkHiddenItems.Value = 0
    chkCurrentRoomMons.Value = 0
    chkGangHouseNumber.Value = 0
    For i = 0 To 9
        chkExits(i).Value = 0
    Next
    chkPlacedItems.Value = 0

End Sub

Private Sub txtAnsiMap_GotFocus()
Call SelectAll(txtAnsiMap)

End Sub

Private Sub txtAttributes_GotFocus()
Call SelectAll(txtAttributes)

End Sub

Private Sub txtCmdText_GotFocus()
Call SelectAll(txtCmdText)

End Sub

Private Sub txtControlRoom_GotFocus()
Call SelectAll(txtControlRoom)

End Sub

Private Sub txtCopper_GotFocus()
Call SelectAll(txtCopper)

End Sub

Private Sub txtCurrentRoomMon_GotFocus(Index As Integer)
Call SelectAll(txtCurrentRoomMon(Index))

End Sub

Private Sub txtDeathRoom_GotFocus()
Call SelectAll(txtDeathRoom)

End Sub

Private Sub txtDelay_GotFocus()
Call SelectAll(txtDelay)

End Sub

Private Sub txtDesc_Change(Index As Integer)
If Index = 6 Then Exit Sub
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

Private Sub txtEndRoom_GotFocus()
Call SelectAll(txtEndRoom)

End Sub

Private Sub txtExitRoom_GotFocus()
Call SelectAll(txtExitRoom)

End Sub

Private Sub txtGangHouseNumber_GotFocus()
Call SelectAll(txtGangHouseNumber)

End Sub

Private Sub txtGold_GotFocus()
Call SelectAll(txtGold)

End Sub

Private Sub txtHiddenNumber_Change(Index As Integer)
On Error GoTo error:

txtHiddenName(Index).Text = GetItemName(Val(txtHiddenNumber(Index).Text))

out:
Exit Sub
error:
Call HandleError("txtHiddenNumber_Change")
Resume out:
End Sub

Private Sub txtHiddenNumber_GotFocus(Index As Integer)
Call SelectAll(txtHiddenNumber(Index))

End Sub

Private Sub txtHiddenQty_GotFocus(Index As Integer)
Call SelectAll(txtHiddenQty(Index))

End Sub

Private Sub txtHiddenUses_GotFocus(Index As Integer)
Call SelectAll(txtHiddenUses(Index))

End Sub

Private Sub txtLight_GotFocus()
Call SelectAll(txtLight)

End Sub

Private Sub txtMaxArea_GotFocus()
Call SelectAll(txtMaxArea)

End Sub

Private Sub txtMaxIndex_GotFocus()
Call SelectAll(txtMaxIndex)

End Sub

Private Sub txtMaxRegen_GotFocus()
Call SelectAll(txtMaxRegen)

End Sub

Private Sub txtMinIndex_GotFocus()
Call SelectAll(txtMinIndex)

End Sub

Private Sub txtName_GotFocus()
Call SelectAll(txtName)

End Sub

Private Sub txtPermNPC_Change()
On Error GoTo error:

txtPermNPCName.Text = GetMonsterName(Val(txtPermNPC.Text))

out:
Exit Sub
error:
Call HandleError("txtPermNPC_Change")
Resume out:
End Sub

Private Sub txtPermNPC_GotFocus()
Call SelectAll(txtPermNPC)

End Sub

Private Sub txtPlacedItems_Change(Index As Integer)
On Error GoTo error:

txtPlacedItemsName(Index).Text = GetItemName(Val(txtPlacedItems(Index).Text))

out:
Exit Sub
error:
Call HandleError("txtPlacedItems_Change")
Resume out:
End Sub

Private Sub txtPlacedItems_GotFocus(Index As Integer)
Call SelectAll(txtPlacedItems(Index))

End Sub

Private Sub txtPlatinum_GotFocus()
Call SelectAll(txtPlatinum)

End Sub

Private Sub txtRoomExit_GotFocus(Index As Integer)
Call SelectAll(txtRoomExit(Index))

End Sub

Private Sub txtRoomLPara1_GotFocus(Index As Integer)
Call SelectAll(txtRoomLPara1(Index))

End Sub

Private Sub txtRoomLPara2_GotFocus(Index As Integer)
Call SelectAll(txtRoomLPara2(Index))

End Sub

Private Sub txtRoomPara_GotFocus(Index As Integer)
Call SelectAll(txtRoomPara(Index))

End Sub

Private Sub txtRoomWPara_GotFocus(Index As Integer)
Call SelectAll(txtRoomWPara(Index))

End Sub

Private Sub txtRunic_GotFocus()
Call SelectAll(txtRunic)

End Sub

Private Sub txtShopNum_GotFocus()
Call SelectAll(txtShopNum)

End Sub

Private Sub txtSilver_GotFocus()
Call SelectAll(txtSilver)

End Sub

Private Sub txtInvisCopper_GotFocus()
Call SelectAll(txtInvisCopper)

End Sub

Private Sub txtInvisSilver_GotFocus()
Call SelectAll(txtInvisSilver)

End Sub

Private Sub txtInvisGold_GotFocus()
Call SelectAll(txtInvisGold)

End Sub

Private Sub txtInvisPlatinum_GotFocus()
Call SelectAll(txtInvisPlatinum)

End Sub

Private Sub txtInvisRunic_GotFocus()
Call SelectAll(txtInvisRunic)

End Sub

Private Sub txtSpell_GotFocus()
Call SelectAll(txtSpell)

End Sub

Private Sub txtStartMap_GotFocus()
Call SelectAll(txtStartMap)

End Sub

Private Sub txtStartRoom_GotFocus()
Call SelectAll(txtStartRoom)

End Sub

Private Sub txtVisibleNumber_Change(Index As Integer)
On Error GoTo error:

txtVisibleName(Index).Text = GetItemName(Val(txtVisibleNumber(Index).Text))

out:
Exit Sub
error:
Call HandleError("txtVisibleNumber_Change")
Resume out:
End Sub

Private Sub txtVisibleNumber_GotFocus(Index As Integer)
Call SelectAll(txtVisibleNumber(Index))

End Sub

Private Sub txtVisibleQty_GotFocus(Index As Integer)
Call SelectAll(txtVisibleQty(Index))

End Sub

Private Sub txtVisibleUses_GotFocus(Index As Integer)
Call SelectAll(txtVisibleUses(Index))

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim i As Integer

cmbType.ListIndex = 0
cmbMonsterType.ListIndex = 0
For i = 0 To 9
    cmbRoomType(i).ListIndex = 0
Next i
Me.Show
Me.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)
If cmdGo.Enabled = False Then
    Cancel = 1
    Exit Sub
End If
End Sub
