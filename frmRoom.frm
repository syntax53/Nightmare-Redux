VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmRoom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Room Editor"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8265
   Icon            =   "frmRoom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   8265
   Begin VB.CommandButton cmdGotoFirstLastMapRoom 
      Caption         =   "Goto Last Room on Map"
      Height          =   375
      Index           =   1
      Left            =   6240
      TabIndex        =   459
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdGotoFirstLastMapRoom 
      Caption         =   "Goto First Room on Map"
      Height          =   375
      Index           =   0
      Left            =   6240
      TabIndex        =   457
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton cmdMinusOneRoom 
      Caption         =   "&- 1 room"
      Height          =   435
      Left            =   6300
      TabIndex        =   4
      Top             =   1260
      Width           =   915
   End
   Begin VB.CommandButton cmdPlusOneRoom 
      Caption         =   "&+ 1 room"
      Height          =   435
      Left            =   7260
      TabIndex        =   3
      Top             =   1260
      Width           =   915
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   660
      MaxLength       =   52
      TabIndex        =   18
      Top             =   60
      Width           =   2535
   End
   Begin VB.TextBox txtRoom 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   60
      Width           =   735
   End
   Begin VB.TextBox txtMap 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   60
      Width           =   735
   End
   Begin VB.CommandButton cmdMap 
      Caption         =   "M&AP"
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
      Left            =   6240
      TabIndex        =   16
      Top             =   6420
      Width           =   1935
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "De&lete This Room"
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert New Room"
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton cmdMapEditor 
      Caption         =   "&Map Editor"
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
      Left            =   6240
      TabIndex        =   15
      Top             =   6060
      Width           =   1935
   End
   Begin VB.CommandButton cmdDiscard 
      Caption         =   "Discar&d"
      Height          =   435
      Left            =   6240
      TabIndex        =   14
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdSave 
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
      Height          =   495
      Left            =   6240
      TabIndex        =   13
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdCopy2Exist 
      Caption         =   "Copy to Existing Room"
      Height          =   375
      Left            =   6240
      TabIndex        =   12
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton cmdCopy2New 
      Caption         =   "Copy to New Room"
      Height          =   375
      Left            =   6240
      TabIndex        =   11
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "&Goto"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   7260
      TabIndex        =   2
      Top             =   240
      Width           =   915
   End
   Begin VB.TextBox txtGotoRoom 
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
      Height          =   360
      Left            =   6300
      TabIndex        =   1
      Text            =   "1"
      Top             =   810
      Width           =   855
   End
   Begin VB.TextBox txtGotoMap 
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
      Height          =   360
      Left            =   6300
      TabIndex        =   0
      Text            =   "1"
      Top             =   240
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   60
      TabIndex        =   23
      Top             =   420
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   11245
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmRoom.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Exits"
      TabPicture(1)   =   "frmRoom.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdGotoRoom(9)"
      Tab(1).Control(1)=   "cmdGotoRoom(8)"
      Tab(1).Control(2)=   "cmdGotoRoom(7)"
      Tab(1).Control(3)=   "cmdGotoRoom(6)"
      Tab(1).Control(4)=   "cmdGotoRoom(5)"
      Tab(1).Control(5)=   "cmdGotoRoom(4)"
      Tab(1).Control(6)=   "cmdGotoRoom(3)"
      Tab(1).Control(7)=   "cmdGotoRoom(2)"
      Tab(1).Control(8)=   "cmdGotoRoom(1)"
      Tab(1).Control(9)=   "cmdGotoRoom(0)"
      Tab(1).Control(10)=   "txtRoomLPara2(0)"
      Tab(1).Control(11)=   "txtRoomLPara2(1)"
      Tab(1).Control(12)=   "txtRoomLPara2(2)"
      Tab(1).Control(13)=   "txtRoomLPara2(3)"
      Tab(1).Control(14)=   "txtRoomLPara2(4)"
      Tab(1).Control(15)=   "txtRoomLPara2(5)"
      Tab(1).Control(16)=   "txtRoomLPara2(6)"
      Tab(1).Control(17)=   "txtRoomLPara2(7)"
      Tab(1).Control(18)=   "txtRoomLPara2(8)"
      Tab(1).Control(19)=   "txtRoomLPara2(9)"
      Tab(1).Control(20)=   "txtRoomLPara1(0)"
      Tab(1).Control(21)=   "txtRoomLPara1(1)"
      Tab(1).Control(22)=   "txtRoomLPara1(2)"
      Tab(1).Control(23)=   "txtRoomLPara1(3)"
      Tab(1).Control(24)=   "txtRoomLPara1(4)"
      Tab(1).Control(25)=   "txtRoomLPara1(5)"
      Tab(1).Control(26)=   "txtRoomLPara1(6)"
      Tab(1).Control(27)=   "txtRoomLPara1(7)"
      Tab(1).Control(28)=   "txtRoomLPara1(8)"
      Tab(1).Control(29)=   "txtRoomLPara1(9)"
      Tab(1).Control(30)=   "txtRoomWPara(0)"
      Tab(1).Control(31)=   "txtRoomWPara(1)"
      Tab(1).Control(32)=   "txtRoomWPara(2)"
      Tab(1).Control(33)=   "txtRoomWPara(3)"
      Tab(1).Control(34)=   "txtRoomWPara(4)"
      Tab(1).Control(35)=   "txtRoomWPara(5)"
      Tab(1).Control(36)=   "txtRoomWPara(6)"
      Tab(1).Control(37)=   "txtRoomWPara(7)"
      Tab(1).Control(38)=   "txtRoomWPara(8)"
      Tab(1).Control(39)=   "txtRoomWPara(9)"
      Tab(1).Control(40)=   "txtRoomPara(9)"
      Tab(1).Control(41)=   "txtRoomPara(8)"
      Tab(1).Control(42)=   "txtRoomPara(7)"
      Tab(1).Control(43)=   "txtRoomPara(6)"
      Tab(1).Control(44)=   "txtRoomPara(5)"
      Tab(1).Control(45)=   "txtRoomPara(4)"
      Tab(1).Control(46)=   "txtRoomPara(3)"
      Tab(1).Control(47)=   "txtRoomPara(2)"
      Tab(1).Control(48)=   "txtRoomPara(1)"
      Tab(1).Control(49)=   "txtRoomPara(0)"
      Tab(1).Control(50)=   "txtRoomExit(9)"
      Tab(1).Control(51)=   "txtRoomExit(8)"
      Tab(1).Control(52)=   "txtRoomExit(7)"
      Tab(1).Control(53)=   "txtRoomExit(6)"
      Tab(1).Control(54)=   "txtRoomExit(5)"
      Tab(1).Control(55)=   "txtRoomExit(4)"
      Tab(1).Control(56)=   "txtRoomExit(3)"
      Tab(1).Control(57)=   "txtRoomExit(2)"
      Tab(1).Control(58)=   "txtRoomExit(1)"
      Tab(1).Control(59)=   "txtRoomExit(0)"
      Tab(1).Control(60)=   "cmbRoomType(0)"
      Tab(1).Control(61)=   "cmbRoomType(1)"
      Tab(1).Control(62)=   "cmbRoomType(2)"
      Tab(1).Control(63)=   "cmbRoomType(3)"
      Tab(1).Control(64)=   "cmbRoomType(4)"
      Tab(1).Control(65)=   "cmbRoomType(5)"
      Tab(1).Control(66)=   "cmbRoomType(6)"
      Tab(1).Control(67)=   "cmbRoomType(7)"
      Tab(1).Control(68)=   "cmbRoomType(8)"
      Tab(1).Control(69)=   "cmbRoomType(9)"
      Tab(1).Control(70)=   "Label7"
      Tab(1).Control(71)=   "Line1"
      Tab(1).Control(72)=   "lblPara1(10)"
      Tab(1).Control(73)=   "lblPara2(10)"
      Tab(1).Control(74)=   "lblPara3(10)"
      Tab(1).Control(75)=   "lblPara4(10)"
      Tab(1).Control(76)=   "lblPara1(9)"
      Tab(1).Control(77)=   "lblPara2(9)"
      Tab(1).Control(78)=   "lblPara3(9)"
      Tab(1).Control(79)=   "lblPara4(9)"
      Tab(1).Control(80)=   "lblPara1(8)"
      Tab(1).Control(81)=   "lblPara2(8)"
      Tab(1).Control(82)=   "lblPara3(8)"
      Tab(1).Control(83)=   "lblPara4(8)"
      Tab(1).Control(84)=   "lblPara1(7)"
      Tab(1).Control(85)=   "lblPara2(7)"
      Tab(1).Control(86)=   "lblPara3(7)"
      Tab(1).Control(87)=   "lblPara4(7)"
      Tab(1).Control(88)=   "lblPara1(6)"
      Tab(1).Control(89)=   "lblPara2(6)"
      Tab(1).Control(90)=   "lblPara3(6)"
      Tab(1).Control(91)=   "lblPara4(6)"
      Tab(1).Control(92)=   "lblPara1(5)"
      Tab(1).Control(93)=   "lblPara2(5)"
      Tab(1).Control(94)=   "lblPara3(5)"
      Tab(1).Control(95)=   "lblPara4(5)"
      Tab(1).Control(96)=   "lblPara1(4)"
      Tab(1).Control(97)=   "lblPara2(4)"
      Tab(1).Control(98)=   "lblPara3(4)"
      Tab(1).Control(99)=   "lblPara4(4)"
      Tab(1).Control(100)=   "lblPara1(3)"
      Tab(1).Control(101)=   "lblPara2(3)"
      Tab(1).Control(102)=   "lblPara3(3)"
      Tab(1).Control(103)=   "lblPara4(3)"
      Tab(1).Control(104)=   "lblPara1(2)"
      Tab(1).Control(105)=   "lblPara2(2)"
      Tab(1).Control(106)=   "lblPara3(2)"
      Tab(1).Control(107)=   "lblPara4(2)"
      Tab(1).Control(108)=   "lblPara1(1)"
      Tab(1).Control(109)=   "lblPara2(1)"
      Tab(1).Control(110)=   "lblPara3(1)"
      Tab(1).Control(111)=   "lblPara4(1)"
      Tab(1).Control(112)=   "lblPara4(0)"
      Tab(1).Control(113)=   "lblPara3(0)"
      Tab(1).Control(114)=   "lblPara2(0)"
      Tab(1).Control(115)=   "lblPara1(0)"
      Tab(1).Control(116)=   "lblRoomNum(0)"
      Tab(1).Control(117)=   "label(13)"
      Tab(1).Control(118)=   "label(12)"
      Tab(1).Control(119)=   "label(10)"
      Tab(1).Control(120)=   "label(9)"
      Tab(1).Control(121)=   "label(8)"
      Tab(1).Control(122)=   "label(7)"
      Tab(1).Control(123)=   "label(6)"
      Tab(1).Control(124)=   "label(5)"
      Tab(1).Control(125)=   "label(4)"
      Tab(1).Control(126)=   "label(3)"
      Tab(1).Control(127)=   "lblExitType(0)"
      Tab(1).ControlCount=   128
      TabCaption(2)   =   "NPC / Items / Coins  "
      TabPicture(2)   =   "frmRoom.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(2)=   "cmdEditPlacedItem(9)"
      Tab(2).Control(3)=   "cmdEditPlacedItem(8)"
      Tab(2).Control(4)=   "cmdEditPlacedItem(7)"
      Tab(2).Control(5)=   "cmdEditPlacedItem(6)"
      Tab(2).Control(6)=   "cmdEditPlacedItem(5)"
      Tab(2).Control(7)=   "cmdEditPlacedItem(4)"
      Tab(2).Control(8)=   "cmdEditPlacedItem(3)"
      Tab(2).Control(9)=   "cmdEditPlacedItem(2)"
      Tab(2).Control(10)=   "cmdEditPlacedItem(1)"
      Tab(2).Control(11)=   "cmdEditPlacedItem(0)"
      Tab(2).Control(12)=   "cmdEditPermNPC"
      Tab(2).Control(13)=   "txtPermNPC"
      Tab(2).Control(14)=   "txtPermNPCName"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "txtPlacedItems(9)"
      Tab(2).Control(16)=   "txtPlacedItems(8)"
      Tab(2).Control(17)=   "txtPlacedItems(7)"
      Tab(2).Control(18)=   "txtPlacedItems(6)"
      Tab(2).Control(19)=   "txtPlacedItems(5)"
      Tab(2).Control(20)=   "txtPlacedItems(4)"
      Tab(2).Control(21)=   "txtPlacedItems(3)"
      Tab(2).Control(22)=   "txtPlacedItems(2)"
      Tab(2).Control(23)=   "txtPlacedItems(1)"
      Tab(2).Control(24)=   "txtPlacedItems(0)"
      Tab(2).Control(25)=   "txtPlacedItemsName(0)"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "txtPlacedItemsName(1)"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "txtPlacedItemsName(2)"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "txtPlacedItemsName(3)"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "txtPlacedItemsName(4)"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "txtPlacedItemsName(5)"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "txtPlacedItemsName(6)"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "txtPlacedItemsName(7)"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "txtPlacedItemsName(8)"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "txtPlacedItemsName(9)"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "label(20)"
      Tab(2).Control(36)=   "label(26)"
      Tab(2).ControlCount=   37
      TabCaption(3)   =   "Visible Items"
      TabPicture(3)   =   "frmRoom.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtVisibleQty(16)"
      Tab(3).Control(1)=   "txtVisibleUses(16)"
      Tab(3).Control(2)=   "txtVisibleName(16)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "txtVisibleNumber(16)"
      Tab(3).Control(4)=   "cmdVisibleGoto(16)"
      Tab(3).Control(5)=   "txtVisibleQty(15)"
      Tab(3).Control(6)=   "txtVisibleUses(15)"
      Tab(3).Control(7)=   "txtVisibleName(15)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "txtVisibleNumber(15)"
      Tab(3).Control(9)=   "cmdVisibleGoto(15)"
      Tab(3).Control(10)=   "txtVisibleQty(14)"
      Tab(3).Control(11)=   "txtVisibleUses(14)"
      Tab(3).Control(12)=   "txtVisibleName(14)"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "txtVisibleNumber(14)"
      Tab(3).Control(14)=   "cmdVisibleGoto(14)"
      Tab(3).Control(15)=   "txtVisibleQty(13)"
      Tab(3).Control(16)=   "txtVisibleUses(13)"
      Tab(3).Control(17)=   "txtVisibleName(13)"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "txtVisibleNumber(13)"
      Tab(3).Control(19)=   "cmdVisibleGoto(13)"
      Tab(3).Control(20)=   "txtVisibleQty(12)"
      Tab(3).Control(21)=   "txtVisibleUses(12)"
      Tab(3).Control(22)=   "txtVisibleName(12)"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "txtVisibleNumber(12)"
      Tab(3).Control(24)=   "cmdVisibleGoto(12)"
      Tab(3).Control(25)=   "txtVisibleQty(11)"
      Tab(3).Control(26)=   "txtVisibleUses(11)"
      Tab(3).Control(27)=   "txtVisibleName(11)"
      Tab(3).Control(27).Enabled=   0   'False
      Tab(3).Control(28)=   "txtVisibleNumber(11)"
      Tab(3).Control(29)=   "cmdVisibleGoto(11)"
      Tab(3).Control(30)=   "txtVisibleQty(10)"
      Tab(3).Control(31)=   "txtVisibleUses(10)"
      Tab(3).Control(32)=   "txtVisibleName(10)"
      Tab(3).Control(32).Enabled=   0   'False
      Tab(3).Control(33)=   "txtVisibleNumber(10)"
      Tab(3).Control(34)=   "cmdVisibleGoto(10)"
      Tab(3).Control(35)=   "txtVisibleQty(9)"
      Tab(3).Control(36)=   "txtVisibleUses(9)"
      Tab(3).Control(37)=   "txtVisibleName(9)"
      Tab(3).Control(37).Enabled=   0   'False
      Tab(3).Control(38)=   "txtVisibleNumber(9)"
      Tab(3).Control(39)=   "cmdVisibleGoto(9)"
      Tab(3).Control(40)=   "txtVisibleQty(8)"
      Tab(3).Control(41)=   "txtVisibleUses(8)"
      Tab(3).Control(42)=   "txtVisibleName(8)"
      Tab(3).Control(42).Enabled=   0   'False
      Tab(3).Control(43)=   "txtVisibleNumber(8)"
      Tab(3).Control(44)=   "cmdVisibleGoto(8)"
      Tab(3).Control(45)=   "txtVisibleQty(7)"
      Tab(3).Control(46)=   "txtVisibleUses(7)"
      Tab(3).Control(47)=   "txtVisibleName(7)"
      Tab(3).Control(47).Enabled=   0   'False
      Tab(3).Control(48)=   "txtVisibleNumber(7)"
      Tab(3).Control(49)=   "cmdVisibleGoto(7)"
      Tab(3).Control(50)=   "txtVisibleQty(6)"
      Tab(3).Control(51)=   "txtVisibleUses(6)"
      Tab(3).Control(52)=   "txtVisibleName(6)"
      Tab(3).Control(52).Enabled=   0   'False
      Tab(3).Control(53)=   "txtVisibleNumber(6)"
      Tab(3).Control(54)=   "cmdVisibleGoto(6)"
      Tab(3).Control(55)=   "txtVisibleQty(5)"
      Tab(3).Control(56)=   "txtVisibleUses(5)"
      Tab(3).Control(57)=   "txtVisibleName(5)"
      Tab(3).Control(57).Enabled=   0   'False
      Tab(3).Control(58)=   "txtVisibleNumber(5)"
      Tab(3).Control(59)=   "cmdVisibleGoto(5)"
      Tab(3).Control(60)=   "txtVisibleQty(4)"
      Tab(3).Control(61)=   "txtVisibleUses(4)"
      Tab(3).Control(62)=   "txtVisibleName(4)"
      Tab(3).Control(62).Enabled=   0   'False
      Tab(3).Control(63)=   "txtVisibleNumber(4)"
      Tab(3).Control(64)=   "cmdVisibleGoto(4)"
      Tab(3).Control(65)=   "txtVisibleQty(3)"
      Tab(3).Control(66)=   "txtVisibleUses(3)"
      Tab(3).Control(67)=   "txtVisibleName(3)"
      Tab(3).Control(67).Enabled=   0   'False
      Tab(3).Control(68)=   "txtVisibleNumber(3)"
      Tab(3).Control(69)=   "cmdVisibleGoto(3)"
      Tab(3).Control(70)=   "txtVisibleQty(2)"
      Tab(3).Control(71)=   "txtVisibleUses(2)"
      Tab(3).Control(72)=   "txtVisibleName(2)"
      Tab(3).Control(72).Enabled=   0   'False
      Tab(3).Control(73)=   "txtVisibleNumber(2)"
      Tab(3).Control(74)=   "cmdVisibleGoto(2)"
      Tab(3).Control(75)=   "txtVisibleQty(1)"
      Tab(3).Control(76)=   "txtVisibleUses(1)"
      Tab(3).Control(77)=   "txtVisibleName(1)"
      Tab(3).Control(77).Enabled=   0   'False
      Tab(3).Control(78)=   "txtVisibleNumber(1)"
      Tab(3).Control(79)=   "cmdVisibleGoto(1)"
      Tab(3).Control(80)=   "txtVisibleQty(0)"
      Tab(3).Control(81)=   "txtVisibleUses(0)"
      Tab(3).Control(82)=   "txtVisibleName(0)"
      Tab(3).Control(82).Enabled=   0   'False
      Tab(3).Control(83)=   "txtVisibleNumber(0)"
      Tab(3).Control(84)=   "cmdVisibleGoto(0)"
      Tab(3).Control(85)=   "Label11"
      Tab(3).Control(86)=   "label(18)"
      Tab(3).Control(87)=   "label(17)"
      Tab(3).Control(88)=   "label(16)"
      Tab(3).Control(89)=   "label(15)"
      Tab(3).ControlCount=   90
      TabCaption(4)   =   "Hidden Items"
      TabPicture(4)   =   "frmRoom.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtHiddenQty(14)"
      Tab(4).Control(1)=   "txtHiddenUses(14)"
      Tab(4).Control(2)=   "txtHiddenName(14)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "txtHiddenNumber(14)"
      Tab(4).Control(4)=   "cmdHiddenGoto(14)"
      Tab(4).Control(5)=   "txtHiddenQty(13)"
      Tab(4).Control(6)=   "txtHiddenUses(13)"
      Tab(4).Control(7)=   "txtHiddenName(13)"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "txtHiddenNumber(13)"
      Tab(4).Control(9)=   "cmdHiddenGoto(13)"
      Tab(4).Control(10)=   "txtHiddenQty(12)"
      Tab(4).Control(11)=   "txtHiddenUses(12)"
      Tab(4).Control(12)=   "txtHiddenName(12)"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "txtHiddenNumber(12)"
      Tab(4).Control(14)=   "cmdHiddenGoto(12)"
      Tab(4).Control(15)=   "txtHiddenQty(11)"
      Tab(4).Control(16)=   "txtHiddenUses(11)"
      Tab(4).Control(17)=   "txtHiddenName(11)"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "txtHiddenNumber(11)"
      Tab(4).Control(19)=   "cmdHiddenGoto(11)"
      Tab(4).Control(20)=   "txtHiddenQty(10)"
      Tab(4).Control(21)=   "txtHiddenUses(10)"
      Tab(4).Control(22)=   "txtHiddenName(10)"
      Tab(4).Control(22).Enabled=   0   'False
      Tab(4).Control(23)=   "txtHiddenNumber(10)"
      Tab(4).Control(24)=   "cmdHiddenGoto(10)"
      Tab(4).Control(25)=   "txtHiddenQty(9)"
      Tab(4).Control(26)=   "txtHiddenUses(9)"
      Tab(4).Control(27)=   "txtHiddenName(9)"
      Tab(4).Control(27).Enabled=   0   'False
      Tab(4).Control(28)=   "txtHiddenNumber(9)"
      Tab(4).Control(29)=   "cmdHiddenGoto(9)"
      Tab(4).Control(30)=   "txtHiddenQty(8)"
      Tab(4).Control(31)=   "txtHiddenUses(8)"
      Tab(4).Control(32)=   "txtHiddenName(8)"
      Tab(4).Control(32).Enabled=   0   'False
      Tab(4).Control(33)=   "txtHiddenNumber(8)"
      Tab(4).Control(34)=   "cmdHiddenGoto(8)"
      Tab(4).Control(35)=   "txtHiddenQty(7)"
      Tab(4).Control(36)=   "txtHiddenUses(7)"
      Tab(4).Control(37)=   "txtHiddenName(7)"
      Tab(4).Control(37).Enabled=   0   'False
      Tab(4).Control(38)=   "txtHiddenNumber(7)"
      Tab(4).Control(39)=   "cmdHiddenGoto(7)"
      Tab(4).Control(40)=   "txtHiddenQty(6)"
      Tab(4).Control(41)=   "txtHiddenUses(6)"
      Tab(4).Control(42)=   "txtHiddenName(6)"
      Tab(4).Control(42).Enabled=   0   'False
      Tab(4).Control(43)=   "txtHiddenNumber(6)"
      Tab(4).Control(44)=   "cmdHiddenGoto(6)"
      Tab(4).Control(45)=   "txtHiddenQty(5)"
      Tab(4).Control(46)=   "txtHiddenUses(5)"
      Tab(4).Control(47)=   "txtHiddenName(5)"
      Tab(4).Control(47).Enabled=   0   'False
      Tab(4).Control(48)=   "txtHiddenNumber(5)"
      Tab(4).Control(49)=   "cmdHiddenGoto(5)"
      Tab(4).Control(50)=   "txtHiddenQty(4)"
      Tab(4).Control(51)=   "txtHiddenUses(4)"
      Tab(4).Control(52)=   "txtHiddenName(4)"
      Tab(4).Control(52).Enabled=   0   'False
      Tab(4).Control(53)=   "txtHiddenNumber(4)"
      Tab(4).Control(54)=   "cmdHiddenGoto(4)"
      Tab(4).Control(55)=   "txtHiddenQty(3)"
      Tab(4).Control(56)=   "txtHiddenUses(3)"
      Tab(4).Control(57)=   "txtHiddenName(3)"
      Tab(4).Control(57).Enabled=   0   'False
      Tab(4).Control(58)=   "txtHiddenNumber(3)"
      Tab(4).Control(59)=   "cmdHiddenGoto(3)"
      Tab(4).Control(60)=   "txtHiddenQty(2)"
      Tab(4).Control(61)=   "txtHiddenUses(2)"
      Tab(4).Control(62)=   "txtHiddenName(2)"
      Tab(4).Control(62).Enabled=   0   'False
      Tab(4).Control(63)=   "txtHiddenNumber(2)"
      Tab(4).Control(64)=   "cmdHiddenGoto(2)"
      Tab(4).Control(65)=   "txtHiddenQty(1)"
      Tab(4).Control(66)=   "txtHiddenUses(1)"
      Tab(4).Control(67)=   "txtHiddenName(1)"
      Tab(4).Control(67).Enabled=   0   'False
      Tab(4).Control(68)=   "txtHiddenNumber(1)"
      Tab(4).Control(69)=   "cmdHiddenGoto(1)"
      Tab(4).Control(70)=   "txtHiddenQty(0)"
      Tab(4).Control(71)=   "txtHiddenUses(0)"
      Tab(4).Control(72)=   "txtHiddenName(0)"
      Tab(4).Control(72).Enabled=   0   'False
      Tab(4).Control(73)=   "txtHiddenNumber(0)"
      Tab(4).Control(74)=   "cmdHiddenGoto(0)"
      Tab(4).Control(75)=   "Label3"
      Tab(4).Control(76)=   "label(36)"
      Tab(4).Control(77)=   "label(35)"
      Tab(4).Control(78)=   "label(27)"
      Tab(4).Control(79)=   "label(19)"
      Tab(4).ControlCount=   80
      TabCaption(5)   =   "Other"
      TabPicture(5)   =   "frmRoom.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txtNote"
      Tab(5).Control(1)=   "txtCurrentRoomMon(1)"
      Tab(5).Control(2)=   "txtCurrentRoomMon(2)"
      Tab(5).Control(3)=   "txtCurrentRoomMon(3)"
      Tab(5).Control(4)=   "txtCurrentRoomMon(4)"
      Tab(5).Control(5)=   "txtCurrentRoomMon(5)"
      Tab(5).Control(6)=   "txtCurrentRoomMon(6)"
      Tab(5).Control(7)=   "txtCurrentRoomMon(7)"
      Tab(5).Control(8)=   "txtCurrentRoomMon(8)"
      Tab(5).Control(9)=   "txtCurrentRoomMon(9)"
      Tab(5).Control(10)=   "txtCurrentRoomMon(10)"
      Tab(5).Control(11)=   "txtCurrentRoomMon(11)"
      Tab(5).Control(12)=   "txtCurrentRoomMon(12)"
      Tab(5).Control(13)=   "txtCurrentRoomMon(13)"
      Tab(5).Control(14)=   "txtCurrentRoomMon(0)"
      Tab(5).Control(15)=   "txtCurrentRoomMon(14)"
      Tab(5).Control(16)=   "Label8"
      Tab(5).ControlCount=   17
      Begin VB.Frame Frame4 
         Caption         =   "Visible Coins"
         Height          =   2295
         Left            =   -71280
         TabIndex        =   446
         Top             =   600
         Width           =   2175
         Begin VB.TextBox txtRunic 
            Height          =   285
            Left            =   1200
            TabIndex        =   451
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtPlatinum 
            Height          =   285
            Left            =   1200
            TabIndex        =   450
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtGold 
            Height          =   285
            Left            =   1200
            TabIndex        =   449
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtSilver 
            Height          =   285
            Left            =   1200
            TabIndex        =   448
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox txtCopper 
            Height          =   285
            Left            =   1200
            TabIndex        =   447
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Runic"
            Height          =   195
            Index           =   44
            Left            =   240
            TabIndex        =   456
            Top             =   360
            Width           =   420
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Platinum"
            Height          =   195
            Index           =   37
            Left            =   240
            TabIndex        =   455
            Top             =   720
            Width           =   600
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Gold"
            Height          =   195
            Index           =   38
            Left            =   240
            TabIndex        =   454
            Top             =   1080
            Width           =   330
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Silver"
            Height          =   195
            Index           =   39
            Left            =   240
            TabIndex        =   453
            Top             =   1440
            Width           =   390
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Copper"
            Height          =   195
            Index           =   40
            Left            =   240
            TabIndex        =   452
            Top             =   1800
            Width           =   510
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Hidden Coins"
         Height          =   2295
         Left            =   -71280
         TabIndex        =   435
         Top             =   3000
         Width           =   2175
         Begin VB.TextBox txtInvisGold 
            Height          =   285
            Left            =   1200
            TabIndex        =   445
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtInvisCopper 
            Height          =   285
            Left            =   1200
            TabIndex        =   444
            Top             =   1800
            Width           =   735
         End
         Begin VB.TextBox txtInvisSilver 
            Height          =   285
            Left            =   1200
            TabIndex        =   443
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox txtInvisPlatinum 
            Height          =   285
            Left            =   1200
            TabIndex        =   442
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtInvisRunic 
            Height          =   285
            Left            =   1200
            TabIndex        =   441
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Copper"
            Height          =   255
            Left            =   240
            TabIndex        =   440
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Silver"
            Height          =   255
            Left            =   240
            TabIndex        =   439
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Gold"
            Height          =   255
            Left            =   240
            TabIndex        =   438
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Platinum"
            Height          =   255
            Left            =   240
            TabIndex        =   437
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Runic"
            Height          =   255
            Left            =   240
            TabIndex        =   436
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.TextBox txtHiddenQty 
         Height          =   285
         Index           =   14
         Left            =   -70395
         TabIndex        =   431
         Top             =   4860
         Width           =   615
      End
      Begin VB.TextBox txtHiddenUses 
         Height          =   285
         Index           =   14
         Left            =   -71010
         TabIndex        =   430
         Top             =   4860
         Width           =   615
      End
      Begin VB.TextBox txtHiddenName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   14
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   429
         TabStop         =   0   'False
         Top             =   4860
         Width           =   2295
      End
      Begin VB.TextBox txtHiddenNumber 
         Height          =   285
         Index           =   14
         Left            =   -73920
         TabIndex        =   428
         Top             =   4860
         Width           =   615
      End
      Begin VB.CommandButton cmdHiddenGoto 
         Height          =   195
         Index           =   14
         Left            =   -74160
         TabIndex        =   427
         Top             =   4905
         Width           =   195
      End
      Begin VB.TextBox txtHiddenQty 
         Height          =   285
         Index           =   13
         Left            =   -70395
         TabIndex        =   426
         Top             =   4560
         Width           =   615
      End
      Begin VB.TextBox txtHiddenUses 
         Height          =   285
         Index           =   13
         Left            =   -71010
         TabIndex        =   425
         Top             =   4560
         Width           =   615
      End
      Begin VB.TextBox txtHiddenName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   13
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   424
         TabStop         =   0   'False
         Top             =   4560
         Width           =   2295
      End
      Begin VB.TextBox txtHiddenNumber 
         Height          =   285
         Index           =   13
         Left            =   -73920
         TabIndex        =   423
         Top             =   4560
         Width           =   615
      End
      Begin VB.CommandButton cmdHiddenGoto 
         Height          =   195
         Index           =   13
         Left            =   -74160
         TabIndex        =   422
         Top             =   4605
         Width           =   195
      End
      Begin VB.TextBox txtHiddenQty 
         Height          =   285
         Index           =   12
         Left            =   -70395
         TabIndex        =   421
         Top             =   4260
         Width           =   615
      End
      Begin VB.TextBox txtHiddenUses 
         Height          =   285
         Index           =   12
         Left            =   -71010
         TabIndex        =   420
         Top             =   4260
         Width           =   615
      End
      Begin VB.TextBox txtHiddenName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   12
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   419
         TabStop         =   0   'False
         Top             =   4260
         Width           =   2295
      End
      Begin VB.TextBox txtHiddenNumber 
         Height          =   285
         Index           =   12
         Left            =   -73920
         TabIndex        =   418
         Top             =   4260
         Width           =   615
      End
      Begin VB.CommandButton cmdHiddenGoto 
         Height          =   195
         Index           =   12
         Left            =   -74160
         TabIndex        =   417
         Top             =   4305
         Width           =   195
      End
      Begin VB.TextBox txtHiddenQty 
         Height          =   285
         Index           =   11
         Left            =   -70395
         TabIndex        =   416
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox txtHiddenUses 
         Height          =   285
         Index           =   11
         Left            =   -71010
         TabIndex        =   415
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox txtHiddenName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   11
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   414
         TabStop         =   0   'False
         Top             =   3960
         Width           =   2295
      End
      Begin VB.TextBox txtHiddenNumber 
         Height          =   285
         Index           =   11
         Left            =   -73920
         TabIndex        =   413
         Top             =   3960
         Width           =   615
      End
      Begin VB.CommandButton cmdHiddenGoto 
         Height          =   195
         Index           =   11
         Left            =   -74160
         TabIndex        =   412
         Top             =   4005
         Width           =   195
      End
      Begin VB.TextBox txtHiddenQty 
         Height          =   285
         Index           =   10
         Left            =   -70395
         TabIndex        =   411
         Top             =   3660
         Width           =   615
      End
      Begin VB.TextBox txtHiddenUses 
         Height          =   285
         Index           =   10
         Left            =   -71010
         TabIndex        =   410
         Top             =   3660
         Width           =   615
      End
      Begin VB.TextBox txtHiddenName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   10
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   409
         TabStop         =   0   'False
         Top             =   3660
         Width           =   2295
      End
      Begin VB.TextBox txtHiddenNumber 
         Height          =   285
         Index           =   10
         Left            =   -73920
         TabIndex        =   408
         Top             =   3660
         Width           =   615
      End
      Begin VB.CommandButton cmdHiddenGoto 
         Height          =   195
         Index           =   10
         Left            =   -74160
         TabIndex        =   407
         Top             =   3705
         Width           =   195
      End
      Begin VB.TextBox txtHiddenQty 
         Height          =   285
         Index           =   9
         Left            =   -70395
         TabIndex        =   406
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox txtHiddenUses 
         Height          =   285
         Index           =   9
         Left            =   -71010
         TabIndex        =   405
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox txtHiddenName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   9
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   404
         TabStop         =   0   'False
         Top             =   3360
         Width           =   2295
      End
      Begin VB.TextBox txtHiddenNumber 
         Height          =   285
         Index           =   9
         Left            =   -73920
         TabIndex        =   403
         Top             =   3360
         Width           =   615
      End
      Begin VB.CommandButton cmdHiddenGoto 
         Height          =   195
         Index           =   9
         Left            =   -74160
         TabIndex        =   402
         Top             =   3405
         Width           =   195
      End
      Begin VB.TextBox txtHiddenQty 
         Height          =   285
         Index           =   8
         Left            =   -70395
         TabIndex        =   401
         Top             =   3060
         Width           =   615
      End
      Begin VB.TextBox txtHiddenUses 
         Height          =   285
         Index           =   8
         Left            =   -71010
         TabIndex        =   400
         Top             =   3060
         Width           =   615
      End
      Begin VB.TextBox txtHiddenName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   8
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   399
         TabStop         =   0   'False
         Top             =   3060
         Width           =   2295
      End
      Begin VB.TextBox txtHiddenNumber 
         Height          =   285
         Index           =   8
         Left            =   -73920
         TabIndex        =   398
         Top             =   3060
         Width           =   615
      End
      Begin VB.CommandButton cmdHiddenGoto 
         Height          =   195
         Index           =   8
         Left            =   -74160
         TabIndex        =   397
         Top             =   3105
         Width           =   195
      End
      Begin VB.TextBox txtHiddenQty 
         Height          =   285
         Index           =   7
         Left            =   -70395
         TabIndex        =   396
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox txtHiddenUses 
         Height          =   285
         Index           =   7
         Left            =   -71010
         TabIndex        =   395
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox txtHiddenName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   7
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   394
         TabStop         =   0   'False
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox txtHiddenNumber 
         Height          =   285
         Index           =   7
         Left            =   -73920
         TabIndex        =   393
         Top             =   2760
         Width           =   615
      End
      Begin VB.CommandButton cmdHiddenGoto 
         Height          =   195
         Index           =   7
         Left            =   -74160
         TabIndex        =   392
         Top             =   2805
         Width           =   195
      End
      Begin VB.TextBox txtHiddenQty 
         Height          =   285
         Index           =   6
         Left            =   -70395
         TabIndex        =   391
         Top             =   2460
         Width           =   615
      End
      Begin VB.TextBox txtHiddenUses 
         Height          =   285
         Index           =   6
         Left            =   -71010
         TabIndex        =   390
         Top             =   2460
         Width           =   615
      End
      Begin VB.TextBox txtHiddenName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   6
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   389
         TabStop         =   0   'False
         Top             =   2460
         Width           =   2295
      End
      Begin VB.TextBox txtHiddenNumber 
         Height          =   285
         Index           =   6
         Left            =   -73920
         TabIndex        =   388
         Top             =   2460
         Width           =   615
      End
      Begin VB.CommandButton cmdHiddenGoto 
         Height          =   195
         Index           =   6
         Left            =   -74160
         TabIndex        =   387
         Top             =   2505
         Width           =   195
      End
      Begin VB.TextBox txtHiddenQty 
         Height          =   285
         Index           =   5
         Left            =   -70395
         TabIndex        =   386
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtHiddenUses 
         Height          =   285
         Index           =   5
         Left            =   -71010
         TabIndex        =   385
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtHiddenName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   5
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   384
         TabStop         =   0   'False
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txtHiddenNumber 
         Height          =   285
         Index           =   5
         Left            =   -73920
         TabIndex        =   383
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton cmdHiddenGoto 
         Height          =   195
         Index           =   5
         Left            =   -74160
         TabIndex        =   382
         Top             =   2205
         Width           =   195
      End
      Begin VB.TextBox txtHiddenQty 
         Height          =   285
         Index           =   4
         Left            =   -70395
         TabIndex        =   381
         Top             =   1860
         Width           =   615
      End
      Begin VB.TextBox txtHiddenUses 
         Height          =   285
         Index           =   4
         Left            =   -71010
         TabIndex        =   380
         Top             =   1860
         Width           =   615
      End
      Begin VB.TextBox txtHiddenName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   4
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   379
         TabStop         =   0   'False
         Top             =   1860
         Width           =   2295
      End
      Begin VB.TextBox txtHiddenNumber 
         Height          =   285
         Index           =   4
         Left            =   -73920
         TabIndex        =   378
         Top             =   1860
         Width           =   615
      End
      Begin VB.CommandButton cmdHiddenGoto 
         Height          =   195
         Index           =   4
         Left            =   -74160
         TabIndex        =   377
         Top             =   1905
         Width           =   195
      End
      Begin VB.TextBox txtHiddenQty 
         Height          =   285
         Index           =   3
         Left            =   -70395
         TabIndex        =   376
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtHiddenUses 
         Height          =   285
         Index           =   3
         Left            =   -71010
         TabIndex        =   375
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtHiddenName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   374
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtHiddenNumber 
         Height          =   285
         Index           =   3
         Left            =   -73920
         TabIndex        =   373
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton cmdHiddenGoto 
         Height          =   195
         Index           =   3
         Left            =   -74160
         TabIndex        =   372
         Top             =   1605
         Width           =   195
      End
      Begin VB.TextBox txtHiddenQty 
         Height          =   285
         Index           =   2
         Left            =   -70395
         TabIndex        =   371
         Top             =   1260
         Width           =   615
      End
      Begin VB.TextBox txtHiddenUses 
         Height          =   285
         Index           =   2
         Left            =   -71010
         TabIndex        =   370
         Top             =   1260
         Width           =   615
      End
      Begin VB.TextBox txtHiddenName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   369
         TabStop         =   0   'False
         Top             =   1260
         Width           =   2295
      End
      Begin VB.TextBox txtHiddenNumber 
         Height          =   285
         Index           =   2
         Left            =   -73920
         TabIndex        =   368
         Top             =   1260
         Width           =   615
      End
      Begin VB.CommandButton cmdHiddenGoto 
         Height          =   195
         Index           =   2
         Left            =   -74160
         TabIndex        =   367
         Top             =   1305
         Width           =   195
      End
      Begin VB.TextBox txtHiddenQty 
         Height          =   285
         Index           =   1
         Left            =   -70395
         TabIndex        =   366
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtHiddenUses 
         Height          =   285
         Index           =   1
         Left            =   -71010
         TabIndex        =   365
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtHiddenName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   364
         TabStop         =   0   'False
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtHiddenNumber 
         Height          =   285
         Index           =   1
         Left            =   -73920
         TabIndex        =   363
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdHiddenGoto 
         Height          =   195
         Index           =   1
         Left            =   -74160
         TabIndex        =   362
         Top             =   1005
         Width           =   195
      End
      Begin VB.TextBox txtHiddenQty 
         Height          =   285
         Index           =   0
         Left            =   -70395
         TabIndex        =   361
         Top             =   660
         Width           =   615
      End
      Begin VB.TextBox txtHiddenUses 
         Height          =   285
         Index           =   0
         Left            =   -71010
         TabIndex        =   360
         Top             =   660
         Width           =   615
      End
      Begin VB.TextBox txtHiddenName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   359
         TabStop         =   0   'False
         Top             =   660
         Width           =   2295
      End
      Begin VB.TextBox txtHiddenNumber 
         Height          =   285
         Index           =   0
         Left            =   -73920
         TabIndex        =   358
         Top             =   660
         Width           =   615
      End
      Begin VB.CommandButton cmdHiddenGoto 
         Height          =   195
         Index           =   0
         Left            =   -74160
         TabIndex        =   357
         Top             =   705
         Width           =   195
      End
      Begin VB.TextBox txtVisibleQty 
         Height          =   285
         Index           =   16
         Left            =   -70395
         TabIndex        =   352
         Top             =   5460
         Width           =   615
      End
      Begin VB.TextBox txtVisibleUses 
         Height          =   285
         Index           =   16
         Left            =   -71010
         TabIndex        =   351
         Top             =   5460
         Width           =   615
      End
      Begin VB.TextBox txtVisibleName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   16
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   350
         TabStop         =   0   'False
         Top             =   5460
         Width           =   2295
      End
      Begin VB.TextBox txtVisibleNumber 
         Height          =   285
         Index           =   16
         Left            =   -73920
         TabIndex        =   349
         Top             =   5460
         Width           =   615
      End
      Begin VB.CommandButton cmdVisibleGoto 
         Height          =   195
         Index           =   16
         Left            =   -74160
         TabIndex        =   348
         Top             =   5505
         Width           =   195
      End
      Begin VB.TextBox txtVisibleQty 
         Height          =   285
         Index           =   15
         Left            =   -70395
         TabIndex        =   347
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox txtVisibleUses 
         Height          =   285
         Index           =   15
         Left            =   -71010
         TabIndex        =   346
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox txtVisibleName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   15
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   345
         TabStop         =   0   'False
         Top             =   5160
         Width           =   2295
      End
      Begin VB.TextBox txtVisibleNumber 
         Height          =   285
         Index           =   15
         Left            =   -73920
         TabIndex        =   344
         Top             =   5160
         Width           =   615
      End
      Begin VB.CommandButton cmdVisibleGoto 
         Height          =   195
         Index           =   15
         Left            =   -74160
         TabIndex        =   343
         Top             =   5205
         Width           =   195
      End
      Begin VB.TextBox txtVisibleQty 
         Height          =   285
         Index           =   14
         Left            =   -70395
         TabIndex        =   342
         Top             =   4860
         Width           =   615
      End
      Begin VB.TextBox txtVisibleUses 
         Height          =   285
         Index           =   14
         Left            =   -71010
         TabIndex        =   341
         Top             =   4860
         Width           =   615
      End
      Begin VB.TextBox txtVisibleName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   14
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   340
         TabStop         =   0   'False
         Top             =   4860
         Width           =   2295
      End
      Begin VB.TextBox txtVisibleNumber 
         Height          =   285
         Index           =   14
         Left            =   -73920
         TabIndex        =   339
         Top             =   4860
         Width           =   615
      End
      Begin VB.CommandButton cmdVisibleGoto 
         Height          =   195
         Index           =   14
         Left            =   -74160
         TabIndex        =   338
         Top             =   4905
         Width           =   195
      End
      Begin VB.TextBox txtVisibleQty 
         Height          =   285
         Index           =   13
         Left            =   -70395
         TabIndex        =   337
         Top             =   4560
         Width           =   615
      End
      Begin VB.TextBox txtVisibleUses 
         Height          =   285
         Index           =   13
         Left            =   -71010
         TabIndex        =   336
         Top             =   4560
         Width           =   615
      End
      Begin VB.TextBox txtVisibleName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   13
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   335
         TabStop         =   0   'False
         Top             =   4560
         Width           =   2295
      End
      Begin VB.TextBox txtVisibleNumber 
         Height          =   285
         Index           =   13
         Left            =   -73920
         TabIndex        =   334
         Top             =   4560
         Width           =   615
      End
      Begin VB.CommandButton cmdVisibleGoto 
         Height          =   195
         Index           =   13
         Left            =   -74160
         TabIndex        =   333
         Top             =   4605
         Width           =   195
      End
      Begin VB.TextBox txtVisibleQty 
         Height          =   285
         Index           =   12
         Left            =   -70395
         TabIndex        =   332
         Top             =   4260
         Width           =   615
      End
      Begin VB.TextBox txtVisibleUses 
         Height          =   285
         Index           =   12
         Left            =   -71010
         TabIndex        =   331
         Top             =   4260
         Width           =   615
      End
      Begin VB.TextBox txtVisibleName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   12
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   330
         TabStop         =   0   'False
         Top             =   4260
         Width           =   2295
      End
      Begin VB.TextBox txtVisibleNumber 
         Height          =   285
         Index           =   12
         Left            =   -73920
         TabIndex        =   329
         Top             =   4260
         Width           =   615
      End
      Begin VB.CommandButton cmdVisibleGoto 
         Height          =   195
         Index           =   12
         Left            =   -74160
         TabIndex        =   328
         Top             =   4305
         Width           =   195
      End
      Begin VB.TextBox txtVisibleQty 
         Height          =   285
         Index           =   11
         Left            =   -70395
         TabIndex        =   327
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox txtVisibleUses 
         Height          =   285
         Index           =   11
         Left            =   -71010
         TabIndex        =   326
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox txtVisibleName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   11
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   325
         TabStop         =   0   'False
         Top             =   3960
         Width           =   2295
      End
      Begin VB.TextBox txtVisibleNumber 
         Height          =   285
         Index           =   11
         Left            =   -73920
         TabIndex        =   324
         Top             =   3960
         Width           =   615
      End
      Begin VB.CommandButton cmdVisibleGoto 
         Height          =   195
         Index           =   11
         Left            =   -74160
         TabIndex        =   323
         Top             =   4005
         Width           =   195
      End
      Begin VB.TextBox txtVisibleQty 
         Height          =   285
         Index           =   10
         Left            =   -70395
         TabIndex        =   322
         Top             =   3660
         Width           =   615
      End
      Begin VB.TextBox txtVisibleUses 
         Height          =   285
         Index           =   10
         Left            =   -71010
         TabIndex        =   321
         Top             =   3660
         Width           =   615
      End
      Begin VB.TextBox txtVisibleName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   10
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   320
         TabStop         =   0   'False
         Top             =   3660
         Width           =   2295
      End
      Begin VB.TextBox txtVisibleNumber 
         Height          =   285
         Index           =   10
         Left            =   -73920
         TabIndex        =   319
         Top             =   3660
         Width           =   615
      End
      Begin VB.CommandButton cmdVisibleGoto 
         Height          =   195
         Index           =   10
         Left            =   -74160
         TabIndex        =   318
         Top             =   3705
         Width           =   195
      End
      Begin VB.TextBox txtVisibleQty 
         Height          =   285
         Index           =   9
         Left            =   -70395
         TabIndex        =   317
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox txtVisibleUses 
         Height          =   285
         Index           =   9
         Left            =   -71010
         TabIndex        =   316
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox txtVisibleName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   9
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   315
         TabStop         =   0   'False
         Top             =   3360
         Width           =   2295
      End
      Begin VB.TextBox txtVisibleNumber 
         Height          =   285
         Index           =   9
         Left            =   -73920
         TabIndex        =   314
         Top             =   3360
         Width           =   615
      End
      Begin VB.CommandButton cmdVisibleGoto 
         Height          =   195
         Index           =   9
         Left            =   -74160
         TabIndex        =   313
         Top             =   3405
         Width           =   195
      End
      Begin VB.TextBox txtVisibleQty 
         Height          =   285
         Index           =   8
         Left            =   -70395
         TabIndex        =   312
         Top             =   3060
         Width           =   615
      End
      Begin VB.TextBox txtVisibleUses 
         Height          =   285
         Index           =   8
         Left            =   -71010
         TabIndex        =   311
         Top             =   3060
         Width           =   615
      End
      Begin VB.TextBox txtVisibleName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   8
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   310
         TabStop         =   0   'False
         Top             =   3060
         Width           =   2295
      End
      Begin VB.TextBox txtVisibleNumber 
         Height          =   285
         Index           =   8
         Left            =   -73920
         TabIndex        =   309
         Top             =   3060
         Width           =   615
      End
      Begin VB.CommandButton cmdVisibleGoto 
         Height          =   195
         Index           =   8
         Left            =   -74160
         TabIndex        =   308
         Top             =   3105
         Width           =   195
      End
      Begin VB.TextBox txtVisibleQty 
         Height          =   285
         Index           =   7
         Left            =   -70395
         TabIndex        =   307
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox txtVisibleUses 
         Height          =   285
         Index           =   7
         Left            =   -71010
         TabIndex        =   306
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox txtVisibleName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   7
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   305
         TabStop         =   0   'False
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox txtVisibleNumber 
         Height          =   285
         Index           =   7
         Left            =   -73920
         TabIndex        =   304
         Top             =   2760
         Width           =   615
      End
      Begin VB.CommandButton cmdVisibleGoto 
         Height          =   195
         Index           =   7
         Left            =   -74160
         TabIndex        =   303
         Top             =   2805
         Width           =   195
      End
      Begin VB.TextBox txtVisibleQty 
         Height          =   285
         Index           =   6
         Left            =   -70395
         TabIndex        =   302
         Top             =   2460
         Width           =   615
      End
      Begin VB.TextBox txtVisibleUses 
         Height          =   285
         Index           =   6
         Left            =   -71010
         TabIndex        =   301
         Top             =   2460
         Width           =   615
      End
      Begin VB.TextBox txtVisibleName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   6
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   300
         TabStop         =   0   'False
         Top             =   2460
         Width           =   2295
      End
      Begin VB.TextBox txtVisibleNumber 
         Height          =   285
         Index           =   6
         Left            =   -73920
         TabIndex        =   299
         Top             =   2460
         Width           =   615
      End
      Begin VB.CommandButton cmdVisibleGoto 
         Height          =   195
         Index           =   6
         Left            =   -74160
         TabIndex        =   298
         Top             =   2505
         Width           =   195
      End
      Begin VB.TextBox txtVisibleQty 
         Height          =   285
         Index           =   5
         Left            =   -70395
         TabIndex        =   297
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtVisibleUses 
         Height          =   285
         Index           =   5
         Left            =   -71010
         TabIndex        =   296
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtVisibleName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   5
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   295
         TabStop         =   0   'False
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txtVisibleNumber 
         Height          =   285
         Index           =   5
         Left            =   -73920
         TabIndex        =   294
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton cmdVisibleGoto 
         Height          =   195
         Index           =   5
         Left            =   -74160
         TabIndex        =   293
         Top             =   2205
         Width           =   195
      End
      Begin VB.TextBox txtVisibleQty 
         Height          =   285
         Index           =   4
         Left            =   -70395
         TabIndex        =   292
         Top             =   1860
         Width           =   615
      End
      Begin VB.TextBox txtVisibleUses 
         Height          =   285
         Index           =   4
         Left            =   -71010
         TabIndex        =   291
         Top             =   1860
         Width           =   615
      End
      Begin VB.TextBox txtVisibleName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   4
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   290
         TabStop         =   0   'False
         Top             =   1860
         Width           =   2295
      End
      Begin VB.TextBox txtVisibleNumber 
         Height          =   285
         Index           =   4
         Left            =   -73920
         TabIndex        =   289
         Top             =   1860
         Width           =   615
      End
      Begin VB.CommandButton cmdVisibleGoto 
         Height          =   195
         Index           =   4
         Left            =   -74160
         TabIndex        =   288
         Top             =   1905
         Width           =   195
      End
      Begin VB.TextBox txtVisibleQty 
         Height          =   285
         Index           =   3
         Left            =   -70395
         TabIndex        =   287
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtVisibleUses 
         Height          =   285
         Index           =   3
         Left            =   -71010
         TabIndex        =   286
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtVisibleName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   285
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtVisibleNumber 
         Height          =   285
         Index           =   3
         Left            =   -73920
         TabIndex        =   284
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton cmdVisibleGoto 
         Height          =   195
         Index           =   3
         Left            =   -74160
         TabIndex        =   283
         Top             =   1605
         Width           =   195
      End
      Begin VB.TextBox txtVisibleQty 
         Height          =   285
         Index           =   2
         Left            =   -70395
         TabIndex        =   282
         Top             =   1260
         Width           =   615
      End
      Begin VB.TextBox txtVisibleUses 
         Height          =   285
         Index           =   2
         Left            =   -71010
         TabIndex        =   281
         Top             =   1260
         Width           =   615
      End
      Begin VB.TextBox txtVisibleName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   280
         TabStop         =   0   'False
         Top             =   1260
         Width           =   2295
      End
      Begin VB.TextBox txtVisibleNumber 
         Height          =   285
         Index           =   2
         Left            =   -73920
         TabIndex        =   279
         Top             =   1260
         Width           =   615
      End
      Begin VB.CommandButton cmdVisibleGoto 
         Height          =   195
         Index           =   2
         Left            =   -74160
         TabIndex        =   278
         Top             =   1305
         Width           =   195
      End
      Begin VB.TextBox txtVisibleQty 
         Height          =   285
         Index           =   1
         Left            =   -70395
         TabIndex        =   277
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtVisibleUses 
         Height          =   285
         Index           =   1
         Left            =   -71010
         TabIndex        =   276
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtVisibleName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   275
         TabStop         =   0   'False
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtVisibleNumber 
         Height          =   285
         Index           =   1
         Left            =   -73920
         TabIndex        =   274
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdVisibleGoto 
         Height          =   195
         Index           =   1
         Left            =   -74160
         TabIndex        =   273
         Top             =   1005
         Width           =   195
      End
      Begin VB.TextBox txtNote 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   -73980
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   262
         Text            =   "frmRoom.frx":0972
         Top             =   1080
         Width           =   4935
      End
      Begin VB.TextBox txtCurrentRoomMon 
         Height          =   285
         Index           =   1
         Left            =   -74820
         TabIndex        =   248
         Top             =   1140
         Width           =   615
      End
      Begin VB.TextBox txtCurrentRoomMon 
         Height          =   285
         Index           =   2
         Left            =   -74820
         TabIndex        =   249
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtCurrentRoomMon 
         Height          =   285
         Index           =   3
         Left            =   -74820
         TabIndex        =   250
         Top             =   1740
         Width           =   615
      End
      Begin VB.TextBox txtCurrentRoomMon 
         Height          =   285
         Index           =   4
         Left            =   -74820
         TabIndex        =   251
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtCurrentRoomMon 
         Height          =   285
         Index           =   5
         Left            =   -74820
         TabIndex        =   252
         Top             =   2340
         Width           =   615
      End
      Begin VB.TextBox txtCurrentRoomMon 
         Height          =   285
         Index           =   6
         Left            =   -74820
         TabIndex        =   253
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox txtCurrentRoomMon 
         Height          =   285
         Index           =   7
         Left            =   -74820
         TabIndex        =   254
         Top             =   2940
         Width           =   615
      End
      Begin VB.TextBox txtCurrentRoomMon 
         Height          =   285
         Index           =   8
         Left            =   -74820
         TabIndex        =   255
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox txtCurrentRoomMon 
         Height          =   285
         Index           =   9
         Left            =   -74820
         TabIndex        =   256
         Top             =   3540
         Width           =   615
      End
      Begin VB.TextBox txtCurrentRoomMon 
         Height          =   285
         Index           =   10
         Left            =   -74820
         TabIndex        =   257
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox txtCurrentRoomMon 
         Height          =   285
         Index           =   11
         Left            =   -74820
         TabIndex        =   258
         Top             =   4140
         Width           =   615
      End
      Begin VB.TextBox txtCurrentRoomMon 
         Height          =   285
         Index           =   12
         Left            =   -74820
         TabIndex        =   259
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox txtCurrentRoomMon 
         Height          =   285
         Index           =   13
         Left            =   -74820
         TabIndex        =   260
         Top             =   4740
         Width           =   615
      End
      Begin VB.TextBox txtCurrentRoomMon 
         Height          =   285
         Index           =   0
         Left            =   -74820
         TabIndex        =   247
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtCurrentRoomMon 
         Height          =   285
         Index           =   14
         Left            =   -74820
         TabIndex        =   261
         Top             =   5040
         Width           =   615
      End
      Begin VB.TextBox txtVisibleQty 
         Height          =   285
         Index           =   0
         Left            =   -70395
         TabIndex        =   272
         Top             =   660
         Width           =   615
      End
      Begin VB.TextBox txtVisibleUses 
         Height          =   285
         Index           =   0
         Left            =   -71010
         TabIndex        =   271
         Top             =   660
         Width           =   615
      End
      Begin VB.TextBox txtVisibleName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   -73305
         Locked          =   -1  'True
         MaxLength       =   28
         TabIndex        =   270
         TabStop         =   0   'False
         Top             =   660
         Width           =   2295
      End
      Begin VB.TextBox txtVisibleNumber 
         Height          =   285
         Index           =   0
         Left            =   -73920
         TabIndex        =   269
         Top             =   660
         Width           =   615
      End
      Begin VB.CommandButton cmdVisibleGoto 
         Height          =   195
         Index           =   0
         Left            =   -74160
         TabIndex        =   268
         Top             =   705
         Width           =   195
      End
      Begin VB.CommandButton cmdGotoRoom 
         Height          =   195
         Index           =   9
         Left            =   -74880
         TabIndex        =   129
         Top             =   5760
         Width           =   195
      End
      Begin VB.CommandButton cmdGotoRoom 
         Height          =   195
         Index           =   8
         Left            =   -74880
         TabIndex        =   122
         Top             =   5220
         Width           =   195
      End
      Begin VB.CommandButton cmdGotoRoom 
         Height          =   195
         Index           =   7
         Left            =   -74880
         TabIndex        =   115
         Top             =   4740
         Width           =   195
      End
      Begin VB.CommandButton cmdGotoRoom 
         Height          =   195
         Index           =   6
         Left            =   -74880
         TabIndex        =   108
         Top             =   4140
         Width           =   195
      End
      Begin VB.CommandButton cmdGotoRoom 
         Height          =   195
         Index           =   5
         Left            =   -74880
         TabIndex        =   101
         Top             =   3600
         Width           =   195
      End
      Begin VB.CommandButton cmdGotoRoom 
         Height          =   195
         Index           =   4
         Left            =   -74880
         TabIndex        =   94
         Top             =   3060
         Width           =   195
      End
      Begin VB.CommandButton cmdGotoRoom 
         Height          =   195
         Index           =   3
         Left            =   -74880
         TabIndex        =   87
         Top             =   2520
         Width           =   195
      End
      Begin VB.CommandButton cmdGotoRoom 
         Height          =   195
         Index           =   2
         Left            =   -74880
         TabIndex        =   80
         Top             =   1980
         Width           =   195
      End
      Begin VB.CommandButton cmdGotoRoom 
         Height          =   195
         Index           =   1
         Left            =   -74880
         TabIndex        =   73
         Top             =   1440
         Width           =   195
      End
      Begin VB.CommandButton cmdGotoRoom 
         Height          =   195
         Index           =   0
         Left            =   -74880
         TabIndex        =   66
         Top             =   900
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPlacedItem 
         Height          =   195
         Index           =   9
         Left            =   -74700
         TabIndex        =   166
         Top             =   4500
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPlacedItem 
         Height          =   195
         Index           =   8
         Left            =   -74700
         TabIndex        =   163
         Top             =   4200
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPlacedItem 
         Height          =   195
         Index           =   7
         Left            =   -74700
         TabIndex        =   160
         Top             =   3900
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPlacedItem 
         Height          =   195
         Index           =   6
         Left            =   -74700
         TabIndex        =   157
         Top             =   3600
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPlacedItem 
         Height          =   195
         Index           =   5
         Left            =   -74700
         TabIndex        =   154
         Top             =   3300
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPlacedItem 
         Height          =   195
         Index           =   4
         Left            =   -74700
         TabIndex        =   151
         Top             =   3000
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPlacedItem 
         Height          =   195
         Index           =   3
         Left            =   -74700
         TabIndex        =   148
         Top             =   2700
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPlacedItem 
         Height          =   195
         Index           =   2
         Left            =   -74700
         TabIndex        =   145
         Top             =   2400
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPlacedItem 
         Height          =   195
         Index           =   1
         Left            =   -74700
         TabIndex        =   142
         Top             =   2100
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPlacedItem 
         Height          =   195
         Index           =   0
         Left            =   -74700
         TabIndex        =   139
         Top             =   1800
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPermNPC 
         Height          =   195
         Left            =   -74700
         TabIndex        =   136
         Top             =   1020
         Width           =   195
      End
      Begin VB.TextBox txtPermNPC 
         Height          =   285
         Left            =   -74400
         TabIndex        =   137
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtPermNPCName 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   -73785
         Locked          =   -1  'True
         TabIndex        =   138
         TabStop         =   0   'False
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtRoomLPara2 
         Height          =   285
         Index           =   0
         Left            =   -69780
         TabIndex        =   72
         Top             =   900
         Width           =   615
      End
      Begin VB.TextBox txtRoomLPara2 
         Height          =   285
         Index           =   1
         Left            =   -69825
         TabIndex        =   79
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtRoomLPara2 
         Height          =   285
         Index           =   2
         Left            =   -69825
         TabIndex        =   86
         Top             =   1980
         Width           =   615
      End
      Begin VB.TextBox txtRoomLPara2 
         Height          =   285
         Index           =   3
         Left            =   -69825
         TabIndex        =   93
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox txtRoomLPara2 
         Height          =   285
         Index           =   4
         Left            =   -69825
         TabIndex        =   100
         Top             =   3060
         Width           =   615
      End
      Begin VB.TextBox txtRoomLPara2 
         Height          =   285
         Index           =   5
         Left            =   -69825
         TabIndex        =   107
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtRoomLPara2 
         Height          =   285
         Index           =   6
         Left            =   -69825
         TabIndex        =   114
         Top             =   4140
         Width           =   615
      End
      Begin VB.TextBox txtRoomLPara2 
         Height          =   285
         Index           =   7
         Left            =   -69825
         TabIndex        =   121
         Top             =   4680
         Width           =   615
      End
      Begin VB.TextBox txtRoomLPara2 
         Height          =   285
         Index           =   8
         Left            =   -69825
         TabIndex        =   128
         Top             =   5220
         Width           =   615
      End
      Begin VB.TextBox txtRoomLPara2 
         Height          =   315
         Index           =   9
         Left            =   -69825
         TabIndex        =   135
         Top             =   5760
         Width           =   615
      End
      Begin VB.TextBox txtRoomLPara1 
         Height          =   285
         Index           =   0
         Left            =   -70545
         TabIndex        =   71
         Top             =   900
         Width           =   615
      End
      Begin VB.TextBox txtRoomLPara1 
         Height          =   285
         Index           =   1
         Left            =   -70500
         TabIndex        =   78
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtRoomLPara1 
         Height          =   285
         Index           =   2
         Left            =   -70545
         TabIndex        =   85
         Top             =   1980
         Width           =   615
      End
      Begin VB.TextBox txtRoomLPara1 
         Height          =   285
         Index           =   3
         Left            =   -70545
         TabIndex        =   92
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox txtRoomLPara1 
         Height          =   285
         Index           =   4
         Left            =   -70545
         TabIndex        =   99
         Top             =   3060
         Width           =   615
      End
      Begin VB.TextBox txtRoomLPara1 
         Height          =   285
         Index           =   5
         Left            =   -70545
         TabIndex        =   106
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtRoomLPara1 
         Height          =   285
         Index           =   6
         Left            =   -70545
         TabIndex        =   113
         Top             =   4140
         Width           =   615
      End
      Begin VB.TextBox txtRoomLPara1 
         Height          =   285
         Index           =   7
         Left            =   -70545
         TabIndex        =   120
         Top             =   4680
         Width           =   615
      End
      Begin VB.TextBox txtRoomLPara1 
         Height          =   285
         Index           =   8
         Left            =   -70545
         TabIndex        =   127
         Top             =   5220
         Width           =   615
      End
      Begin VB.TextBox txtRoomLPara1 
         Height          =   315
         Index           =   9
         Left            =   -70545
         TabIndex        =   134
         Top             =   5760
         Width           =   615
      End
      Begin VB.TextBox txtRoomWPara 
         Height          =   285
         Index           =   0
         Left            =   -71265
         TabIndex        =   70
         Top             =   900
         Width           =   615
      End
      Begin VB.TextBox txtRoomWPara 
         Height          =   285
         Index           =   1
         Left            =   -71265
         TabIndex        =   77
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtRoomWPara 
         Height          =   285
         Index           =   2
         Left            =   -71265
         TabIndex        =   84
         Top             =   1980
         Width           =   615
      End
      Begin VB.TextBox txtRoomWPara 
         Height          =   285
         Index           =   3
         Left            =   -71265
         TabIndex        =   91
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox txtRoomWPara 
         Height          =   285
         Index           =   4
         Left            =   -71265
         TabIndex        =   98
         Top             =   3060
         Width           =   615
      End
      Begin VB.TextBox txtRoomWPara 
         Height          =   285
         Index           =   5
         Left            =   -71265
         TabIndex        =   105
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtRoomWPara 
         Height          =   285
         Index           =   6
         Left            =   -71265
         TabIndex        =   112
         Top             =   4140
         Width           =   615
      End
      Begin VB.TextBox txtRoomWPara 
         Height          =   285
         Index           =   7
         Left            =   -71265
         TabIndex        =   119
         Top             =   4680
         Width           =   615
      End
      Begin VB.TextBox txtRoomWPara 
         Height          =   285
         Index           =   8
         Left            =   -71265
         TabIndex        =   126
         Top             =   5220
         Width           =   615
      End
      Begin VB.TextBox txtRoomWPara 
         Height          =   315
         Index           =   9
         Left            =   -71265
         TabIndex        =   133
         Top             =   5760
         Width           =   615
      End
      Begin VB.TextBox txtRoomPara 
         Height          =   315
         Index           =   9
         Left            =   -72000
         TabIndex        =   132
         Top             =   5760
         Width           =   615
      End
      Begin VB.TextBox txtRoomPara 
         Height          =   285
         Index           =   8
         Left            =   -72000
         TabIndex        =   125
         Top             =   5220
         Width           =   615
      End
      Begin VB.TextBox txtRoomPara 
         Height          =   285
         Index           =   7
         Left            =   -72000
         TabIndex        =   118
         Top             =   4680
         Width           =   615
      End
      Begin VB.TextBox txtRoomPara 
         Height          =   285
         Index           =   6
         Left            =   -72000
         TabIndex        =   111
         Top             =   4140
         Width           =   615
      End
      Begin VB.TextBox txtRoomPara 
         Height          =   285
         Index           =   5
         Left            =   -72000
         TabIndex        =   104
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtRoomPara 
         Height          =   285
         Index           =   4
         Left            =   -72000
         TabIndex        =   97
         Top             =   3060
         Width           =   615
      End
      Begin VB.TextBox txtRoomPara 
         Height          =   285
         Index           =   3
         Left            =   -72000
         TabIndex        =   90
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox txtRoomPara 
         Height          =   285
         Index           =   2
         Left            =   -72000
         TabIndex        =   83
         Top             =   1980
         Width           =   615
      End
      Begin VB.TextBox txtRoomPara 
         Height          =   285
         Index           =   1
         Left            =   -72000
         TabIndex        =   76
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtRoomPara 
         Height          =   285
         Index           =   0
         Left            =   -72000
         TabIndex        =   69
         Top             =   900
         Width           =   615
      End
      Begin VB.TextBox txtRoomExit 
         Height          =   315
         Index           =   9
         Left            =   -74265
         TabIndex        =   130
         Top             =   5760
         Width           =   615
      End
      Begin VB.TextBox txtRoomExit 
         Height          =   285
         Index           =   8
         Left            =   -74265
         TabIndex        =   123
         Top             =   5220
         Width           =   615
      End
      Begin VB.TextBox txtRoomExit 
         Height          =   285
         Index           =   7
         Left            =   -74265
         TabIndex        =   116
         Top             =   4680
         Width           =   615
      End
      Begin VB.TextBox txtRoomExit 
         Height          =   285
         Index           =   6
         Left            =   -74265
         TabIndex        =   109
         Top             =   4140
         Width           =   615
      End
      Begin VB.TextBox txtRoomExit 
         Height          =   285
         Index           =   5
         Left            =   -74265
         TabIndex        =   102
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtRoomExit 
         Height          =   285
         Index           =   4
         Left            =   -74265
         TabIndex        =   95
         Top             =   3060
         Width           =   615
      End
      Begin VB.TextBox txtRoomExit 
         Height          =   285
         Index           =   3
         Left            =   -74265
         TabIndex        =   88
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox txtRoomExit 
         Height          =   285
         Index           =   2
         Left            =   -74265
         TabIndex        =   81
         Top             =   1980
         Width           =   615
      End
      Begin VB.TextBox txtRoomExit 
         Height          =   285
         Index           =   1
         Left            =   -74265
         TabIndex        =   74
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtRoomExit 
         Height          =   285
         Index           =   0
         Left            =   -74265
         TabIndex        =   67
         Top             =   900
         Width           =   615
      End
      Begin VB.ComboBox cmbRoomType 
         Height          =   315
         Index           =   0
         ItemData        =   "frmRoom.frx":0AF1
         Left            =   -73575
         List            =   "frmRoom.frx":0B40
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   900
         Width           =   1455
      End
      Begin VB.ComboBox cmbRoomType 
         Height          =   315
         Index           =   1
         ItemData        =   "frmRoom.frx":0C10
         Left            =   -73575
         List            =   "frmRoom.frx":0C5F
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   1440
         Width           =   1455
      End
      Begin VB.ComboBox cmbRoomType 
         Height          =   315
         Index           =   2
         ItemData        =   "frmRoom.frx":0D2F
         Left            =   -73575
         List            =   "frmRoom.frx":0D7E
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   1980
         Width           =   1455
      End
      Begin VB.ComboBox cmbRoomType 
         Height          =   315
         Index           =   3
         ItemData        =   "frmRoom.frx":0E4E
         Left            =   -73575
         List            =   "frmRoom.frx":0E9D
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   2520
         Width           =   1455
      End
      Begin VB.ComboBox cmbRoomType 
         Height          =   315
         Index           =   4
         ItemData        =   "frmRoom.frx":0F6D
         Left            =   -73575
         List            =   "frmRoom.frx":0FBC
         Style           =   2  'Dropdown List
         TabIndex        =   96
         Top             =   3060
         Width           =   1455
      End
      Begin VB.ComboBox cmbRoomType 
         Height          =   315
         Index           =   5
         ItemData        =   "frmRoom.frx":108C
         Left            =   -73575
         List            =   "frmRoom.frx":10DB
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   3600
         Width           =   1455
      End
      Begin VB.ComboBox cmbRoomType 
         Height          =   315
         Index           =   6
         ItemData        =   "frmRoom.frx":11AB
         Left            =   -73575
         List            =   "frmRoom.frx":11FA
         Style           =   2  'Dropdown List
         TabIndex        =   110
         Top             =   4140
         Width           =   1455
      End
      Begin VB.ComboBox cmbRoomType 
         Height          =   315
         Index           =   7
         ItemData        =   "frmRoom.frx":12CA
         Left            =   -73575
         List            =   "frmRoom.frx":1319
         Style           =   2  'Dropdown List
         TabIndex        =   117
         Top             =   4680
         Width           =   1455
      End
      Begin VB.ComboBox cmbRoomType 
         Height          =   315
         Index           =   8
         ItemData        =   "frmRoom.frx":13E9
         Left            =   -73575
         List            =   "frmRoom.frx":1438
         Style           =   2  'Dropdown List
         TabIndex        =   124
         Top             =   5220
         Width           =   1455
      End
      Begin VB.ComboBox cmbRoomType 
         Height          =   315
         Index           =   9
         ItemData        =   "frmRoom.frx":1508
         Left            =   -73575
         List            =   "frmRoom.frx":1557
         Style           =   2  'Dropdown List
         TabIndex        =   131
         Top             =   5760
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   "Description"
         Height          =   2835
         Left            =   180
         TabIndex        =   24
         Top             =   360
         Width           =   5715
         Begin VB.CommandButton cmdCopyDesc 
            Caption         =   "Pas&te"
            Height          =   315
            Index           =   1
            Left            =   1020
            TabIndex        =   33
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "Find Ne&xt"
            Height          =   315
            Index           =   1
            Left            =   3180
            TabIndex        =   35
            Top             =   2400
            Width           =   1095
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "Find"
            Height          =   315
            Index           =   0
            Left            =   2220
            TabIndex        =   34
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            Height          =   315
            Left            =   4500
            TabIndex        =   36
            Top             =   2400
            Width           =   1095
         End
         Begin VB.CommandButton cmdCopyDesc 
            Caption         =   "Copy"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   32
            Top             =   2400
            Width           =   915
         End
         Begin VB.TextBox txtDesc 
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   6
            Left            =   120
            MaxLength       =   70
            TabIndex        =   31
            Top             =   2040
            Width           =   5475
         End
         Begin VB.TextBox txtDesc 
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   5
            Left            =   120
            MaxLength       =   70
            TabIndex        =   30
            Top             =   1740
            Width           =   5475
         End
         Begin VB.TextBox txtDesc 
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   4
            Left            =   120
            MaxLength       =   70
            TabIndex        =   29
            Top             =   1440
            Width           =   5475
         End
         Begin VB.TextBox txtDesc 
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   3
            Left            =   120
            MaxLength       =   70
            TabIndex        =   28
            Top             =   1140
            Width           =   5475
         End
         Begin VB.TextBox txtDesc 
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   2
            Left            =   120
            MaxLength       =   70
            TabIndex        =   27
            Top             =   840
            Width           =   5475
         End
         Begin VB.TextBox txtDesc 
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   0
            Left            =   120
            MaxLength       =   70
            TabIndex        =   25
            Top             =   240
            Width           =   5475
         End
         Begin VB.TextBox txtDesc 
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   1
            Left            =   120
            MaxLength       =   70
            TabIndex        =   26
            Top             =   540
            Width           =   5475
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Advanced"
         Height          =   3075
         Left            =   180
         TabIndex        =   37
         Top             =   3180
         Width           =   5715
         Begin VB.CommandButton cmdClearAdvanced 
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
            Height          =   315
            Left            =   4680
            TabIndex        =   65
            Top             =   1680
            Width           =   675
         End
         Begin VB.CommandButton cmdCopyPaste 
            Caption         =   "Past&e"
            Height          =   315
            Index           =   1
            Left            =   4500
            TabIndex        =   64
            Top             =   1080
            Width           =   1035
         End
         Begin VB.CommandButton cmdViewIndex 
            Caption         =   "Index List"
            Height          =   315
            Left            =   4500
            TabIndex        =   62
            Top             =   240
            Width           =   1035
         End
         Begin VB.CommandButton cmdCopyPaste 
            Caption         =   "&Copy"
            Height          =   315
            Index           =   0
            Left            =   4500
            TabIndex        =   63
            Top             =   780
            Width           =   1035
         End
         Begin VB.CommandButton cmdAttributesQ 
            Caption         =   "?"
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
            Left            =   1980
            TabIndex        =   41
            Top             =   885
            Width           =   195
         End
         Begin VB.TextBox txtGangHouseNumber 
            Height          =   285
            Left            =   1320
            TabIndex        =   39
            Top             =   540
            Width           =   615
         End
         Begin VB.CommandButton cmdEditExitRoom 
            Height          =   195
            Left            =   1980
            TabIndex        =   49
            Top             =   2085
            Width           =   195
         End
         Begin VB.CommandButton cmdEditDeathRoom 
            Height          =   195
            Left            =   1980
            TabIndex        =   47
            Top             =   1785
            Width           =   195
         End
         Begin VB.CommandButton cmdEditShop 
            Height          =   195
            Left            =   1980
            TabIndex        =   43
            Top             =   1200
            Width           =   195
         End
         Begin VB.CommandButton cmdEditControlRoom 
            Height          =   195
            Left            =   4320
            TabIndex        =   57
            Top             =   1500
            Width           =   195
         End
         Begin VB.CommandButton cmdEditCmdText 
            Height          =   195
            Left            =   1980
            TabIndex        =   51
            Top             =   2400
            Width           =   195
         End
         Begin VB.CommandButton cmdEditSpell 
            Height          =   195
            Left            =   1980
            TabIndex        =   45
            Top             =   1500
            Width           =   195
         End
         Begin VB.TextBox txtCmdTextDisplay 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   205
            Top             =   2640
            Width           =   1755
         End
         Begin VB.ComboBox cmbType 
            Height          =   315
            ItemData        =   "frmRoom.frx":1627
            Left            =   3660
            List            =   "frmRoom.frx":1643
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   2355
            Width           =   1875
         End
         Begin VB.ComboBox cmbMonsterType 
            Height          =   315
            ItemData        =   "frmRoom.frx":1683
            Left            =   3660
            List            =   "frmRoom.frx":16FF
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   2040
            Width           =   1875
         End
         Begin VB.TextBox txtAttributes 
            Height          =   285
            Left            =   1320
            TabIndex        =   40
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtExitRoom 
            Height          =   285
            Left            =   1320
            TabIndex        =   48
            Top             =   2040
            Width           =   615
         End
         Begin VB.TextBox txtSpell 
            Height          =   285
            Left            =   1320
            TabIndex        =   44
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtDelay 
            Height          =   285
            Left            =   3660
            TabIndex        =   55
            Top             =   1140
            Width           =   615
         End
         Begin VB.TextBox txtMaxArea 
            Height          =   285
            Left            =   3660
            TabIndex        =   58
            Top             =   1740
            Width           =   615
         End
         Begin VB.TextBox txtControlRoom 
            Height          =   285
            Left            =   3660
            TabIndex        =   56
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtCmdText 
            Height          =   285
            Left            =   1320
            TabIndex        =   50
            Top             =   2355
            Width           =   615
         End
         Begin VB.TextBox txtDeathRoom 
            Height          =   285
            Left            =   1320
            TabIndex        =   46
            Top             =   1740
            Width           =   615
         End
         Begin VB.TextBox txtShopNum 
            Height          =   285
            Left            =   1320
            TabIndex        =   42
            Top             =   1140
            Width           =   615
         End
         Begin VB.TextBox txtAnsiMap 
            Height          =   285
            Left            =   3660
            TabIndex        =   61
            Top             =   2700
            Width           =   1875
         End
         Begin VB.TextBox txtMaxIndex 
            Height          =   285
            Left            =   3660
            TabIndex        =   53
            Top             =   540
            Width           =   615
         End
         Begin VB.TextBox txtMinIndex 
            Height          =   285
            Left            =   3660
            TabIndex        =   52
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtMaxRegen 
            Height          =   285
            Left            =   3660
            TabIndex        =   54
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtLight 
            Height          =   285
            Left            =   1320
            TabIndex        =   38
            Top             =   240
            Width           =   615
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "GangHouse #"
            Height          =   195
            Index           =   14
            Left            =   180
            TabIndex        =   246
            Top             =   540
            Width           =   1005
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Attributes"
            Height          =   195
            Index           =   43
            Left            =   180
            TabIndex        =   185
            Top             =   840
            Width           =   660
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Exit Room"
            Height          =   195
            Index           =   42
            Left            =   180
            TabIndex        =   184
            Top             =   2040
            Width           =   720
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Room Spell"
            Height          =   195
            Index           =   41
            Left            =   180
            TabIndex        =   183
            Top             =   1440
            Width           =   810
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Delay"
            Height          =   195
            Index           =   34
            Left            =   2460
            TabIndex        =   182
            Top             =   1140
            Width           =   405
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Max Area"
            Height          =   195
            Index           =   33
            Left            =   2460
            TabIndex        =   181
            Top             =   1755
            Width           =   675
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "ControlRoom"
            Height          =   195
            Index           =   32
            Left            =   2460
            TabIndex        =   180
            Top             =   1470
            Width           =   915
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Cmd Text"
            Height          =   195
            Index           =   31
            Left            =   180
            TabIndex        =   179
            Top             =   2355
            Width           =   675
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Death Room"
            Height          =   195
            Index           =   30
            Left            =   180
            TabIndex        =   178
            Top             =   1740
            Width           =   900
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Monster Type"
            Height          =   195
            Index           =   29
            Left            =   2460
            TabIndex        =   177
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Max Regen"
            Height          =   195
            Index           =   28
            Left            =   2460
            TabIndex        =   176
            Top             =   840
            Width           =   825
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Light"
            Height          =   195
            Index           =   25
            Left            =   180
            TabIndex        =   175
            Top             =   240
            Width           =   345
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Max Index"
            Height          =   195
            Index           =   24
            Left            =   2460
            TabIndex        =   174
            Top             =   540
            Width           =   735
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Min Index"
            Height          =   195
            Index           =   23
            Left            =   2460
            TabIndex        =   173
            Top             =   240
            Width           =   690
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Shop #"
            Height          =   195
            Index           =   22
            Left            =   180
            TabIndex        =   172
            Top             =   1140
            Width           =   525
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Room Type"
            Height          =   195
            Index           =   21
            Left            =   2460
            TabIndex        =   171
            Top             =   2370
            Width           =   825
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Ansi Map"
            Height          =   195
            Index           =   2
            Left            =   2460
            TabIndex        =   170
            Top             =   2700
            Width           =   660
         End
      End
      Begin VB.TextBox txtPlacedItems 
         Height          =   285
         Index           =   9
         Left            =   -74400
         TabIndex        =   167
         Top             =   4500
         Width           =   615
      End
      Begin VB.TextBox txtPlacedItems 
         Height          =   285
         Index           =   8
         Left            =   -74400
         TabIndex        =   164
         Top             =   4200
         Width           =   615
      End
      Begin VB.TextBox txtPlacedItems 
         Height          =   285
         Index           =   7
         Left            =   -74400
         TabIndex        =   161
         Top             =   3900
         Width           =   615
      End
      Begin VB.TextBox txtPlacedItems 
         Height          =   285
         Index           =   6
         Left            =   -74400
         TabIndex        =   158
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtPlacedItems 
         Height          =   285
         Index           =   5
         Left            =   -74400
         TabIndex        =   155
         Top             =   3300
         Width           =   615
      End
      Begin VB.TextBox txtPlacedItems 
         Height          =   285
         Index           =   4
         Left            =   -74400
         TabIndex        =   152
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox txtPlacedItems 
         Height          =   285
         Index           =   3
         Left            =   -74400
         TabIndex        =   149
         Top             =   2700
         Width           =   615
      End
      Begin VB.TextBox txtPlacedItems 
         Height          =   285
         Index           =   2
         Left            =   -74400
         TabIndex        =   146
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox txtPlacedItems 
         Height          =   285
         Index           =   1
         Left            =   -74400
         TabIndex        =   143
         Top             =   2100
         Width           =   615
      End
      Begin VB.TextBox txtPlacedItems 
         Height          =   285
         Index           =   0
         Left            =   -74400
         TabIndex        =   140
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtPlacedItemsName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   -73800
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   141
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txtPlacedItemsName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   -73800
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   144
         TabStop         =   0   'False
         Top             =   2100
         Width           =   2295
      End
      Begin VB.TextBox txtPlacedItemsName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   -73800
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   147
         TabStop         =   0   'False
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox txtPlacedItemsName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   -73800
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   150
         TabStop         =   0   'False
         Top             =   2700
         Width           =   2295
      End
      Begin VB.TextBox txtPlacedItemsName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   4
         Left            =   -73800
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   153
         TabStop         =   0   'False
         Top             =   3000
         Width           =   2295
      End
      Begin VB.TextBox txtPlacedItemsName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   5
         Left            =   -73800
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   156
         TabStop         =   0   'False
         Top             =   3300
         Width           =   2295
      End
      Begin VB.TextBox txtPlacedItemsName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   6
         Left            =   -73800
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   159
         TabStop         =   0   'False
         Top             =   3600
         Width           =   2295
      End
      Begin VB.TextBox txtPlacedItemsName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   7
         Left            =   -73800
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   162
         TabStop         =   0   'False
         Top             =   3900
         Width           =   2295
      End
      Begin VB.TextBox txtPlacedItemsName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   8
         Left            =   -73800
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   165
         TabStop         =   0   'False
         Top             =   4200
         Width           =   2295
      End
      Begin VB.TextBox txtPlacedItemsName 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   9
         Left            =   -73800
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   168
         TabStop         =   0   'False
         Top             =   4500
         Width           =   2295
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Hover over the para description for a full description.  If any of these are wrong tell me!"
         Enabled         =   0   'False
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
         Left            =   -74580
         TabIndex        =   434
         Top             =   6120
         Width           =   5595
      End
      Begin VB.Label Label3 
         Caption         =   "The quantity totals here are the total minus one.  (EX: enter 0 for 1 ... 1 for 2 ... 3 for 4, etc.)"
         Height          =   435
         Left            =   -73920
         TabIndex        =   433
         Top             =   5280
         Width           =   3435
      End
      Begin VB.Label Label11 
         Caption         =   "The quantity totals here are the total minus one.  (EX: enter 0 for 1 ... 1 for 2 ... 3 for 4, etc.)"
         Height          =   435
         Left            =   -73920
         TabIndex        =   432
         Top             =   5820
         Width           =   3435
      End
      Begin VB.Label label 
         Alignment       =   2  'Center
         Caption         =   "#"
         Height          =   255
         Index           =   36
         Left            =   -73920
         TabIndex        =   353
         Top             =   420
         Width           =   615
      End
      Begin VB.Label label 
         Alignment       =   2  'Center
         Caption         =   "Name"
         Height          =   255
         Index           =   35
         Left            =   -73320
         TabIndex        =   354
         Top             =   420
         Width           =   2295
      End
      Begin VB.Label label 
         Alignment       =   2  'Center
         Caption         =   "Uses"
         Height          =   255
         Index           =   27
         Left            =   -70980
         TabIndex        =   355
         Top             =   420
         Width           =   615
      End
      Begin VB.Label label 
         Alignment       =   2  'Center
         Caption         =   "Qty -1"
         Height          =   255
         Index           =   19
         Left            =   -70380
         TabIndex        =   356
         Top             =   420
         Width           =   615
      End
      Begin VB.Label label 
         Alignment       =   2  'Center
         Caption         =   "#"
         Height          =   255
         Index           =   18
         Left            =   -73920
         TabIndex        =   264
         Top             =   420
         Width           =   615
      End
      Begin VB.Label label 
         Alignment       =   2  'Center
         Caption         =   "Name"
         Height          =   255
         Index           =   17
         Left            =   -73320
         TabIndex        =   265
         Top             =   420
         Width           =   2295
      End
      Begin VB.Label label 
         Alignment       =   2  'Center
         Caption         =   "Uses"
         Height          =   255
         Index           =   16
         Left            =   -70980
         TabIndex        =   266
         Top             =   420
         Width           =   615
      End
      Begin VB.Label label 
         Alignment       =   2  'Center
         Caption         =   "Qty -1"
         Height          =   255
         Index           =   15
         Left            =   -70380
         TabIndex        =   267
         Top             =   420
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Current Monsters In Room"
         Height          =   255
         Left            =   -74820
         TabIndex        =   263
         Top             =   480
         Width           =   2055
      End
      Begin VB.Line Line1 
         X1              =   -74400
         X2              =   -69120
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Label lblPara1 
         Alignment       =   2  'Center
         Caption         =   "Para1"
         Height          =   195
         Index           =   10
         Left            =   -72000
         TabIndex        =   245
         Top             =   420
         Width           =   570
      End
      Begin VB.Label lblPara2 
         Alignment       =   2  'Center
         Caption         =   "Para2"
         Height          =   195
         Index           =   10
         Left            =   -71250
         TabIndex        =   244
         Top             =   420
         Width           =   570
      End
      Begin VB.Label lblPara3 
         Alignment       =   2  'Center
         Caption         =   "Para3"
         Height          =   195
         Index           =   10
         Left            =   -70530
         TabIndex        =   243
         Top             =   420
         Width           =   570
      End
      Begin VB.Label lblPara4 
         Alignment       =   2  'Center
         Caption         =   "Para4"
         Height          =   195
         Index           =   10
         Left            =   -69810
         TabIndex        =   242
         Top             =   420
         Width           =   570
      End
      Begin VB.Label lblPara1 
         Alignment       =   2  'Center
         Caption         =   "Para1"
         Height          =   195
         Index           =   9
         Left            =   -72060
         TabIndex        =   241
         Top             =   5580
         Width           =   690
      End
      Begin VB.Label lblPara2 
         Alignment       =   2  'Center
         Caption         =   "Para2"
         Height          =   195
         Index           =   9
         Left            =   -71310
         TabIndex        =   240
         Top             =   5580
         Width           =   690
      End
      Begin VB.Label lblPara3 
         Alignment       =   2  'Center
         Caption         =   "Para3"
         Height          =   195
         Index           =   9
         Left            =   -70590
         TabIndex        =   239
         Top             =   5580
         Width           =   690
      End
      Begin VB.Label lblPara4 
         Alignment       =   2  'Center
         Caption         =   "Para4"
         Height          =   195
         Index           =   9
         Left            =   -69870
         TabIndex        =   238
         Top             =   5580
         Width           =   690
      End
      Begin VB.Label lblPara1 
         Alignment       =   2  'Center
         Caption         =   "Para1"
         Height          =   195
         Index           =   8
         Left            =   -72060
         TabIndex        =   237
         Top             =   5040
         Width           =   690
      End
      Begin VB.Label lblPara2 
         Alignment       =   2  'Center
         Caption         =   "Para2"
         Height          =   195
         Index           =   8
         Left            =   -71310
         TabIndex        =   236
         Top             =   5040
         Width           =   690
      End
      Begin VB.Label lblPara3 
         Alignment       =   2  'Center
         Caption         =   "Para3"
         Height          =   195
         Index           =   8
         Left            =   -70590
         TabIndex        =   235
         Top             =   5040
         Width           =   690
      End
      Begin VB.Label lblPara4 
         Alignment       =   2  'Center
         Caption         =   "Para4"
         Height          =   195
         Index           =   8
         Left            =   -69870
         TabIndex        =   234
         Top             =   5040
         Width           =   690
      End
      Begin VB.Label lblPara1 
         Alignment       =   2  'Center
         Caption         =   "Para1"
         Height          =   195
         Index           =   7
         Left            =   -72060
         TabIndex        =   233
         Top             =   4500
         Width           =   690
      End
      Begin VB.Label lblPara2 
         Alignment       =   2  'Center
         Caption         =   "Para2"
         Height          =   195
         Index           =   7
         Left            =   -71310
         TabIndex        =   232
         Top             =   4500
         Width           =   690
      End
      Begin VB.Label lblPara3 
         Alignment       =   2  'Center
         Caption         =   "Para3"
         Height          =   195
         Index           =   7
         Left            =   -70590
         TabIndex        =   231
         Top             =   4500
         Width           =   690
      End
      Begin VB.Label lblPara4 
         Alignment       =   2  'Center
         Caption         =   "Para4"
         Height          =   195
         Index           =   7
         Left            =   -69870
         TabIndex        =   230
         Top             =   4500
         Width           =   690
      End
      Begin VB.Label lblPara1 
         Alignment       =   2  'Center
         Caption         =   "Para1"
         Height          =   195
         Index           =   6
         Left            =   -72060
         TabIndex        =   229
         Top             =   3960
         Width           =   690
      End
      Begin VB.Label lblPara2 
         Alignment       =   2  'Center
         Caption         =   "Para2"
         Height          =   195
         Index           =   6
         Left            =   -71310
         TabIndex        =   228
         Top             =   3960
         Width           =   690
      End
      Begin VB.Label lblPara3 
         Alignment       =   2  'Center
         Caption         =   "Para3"
         Height          =   195
         Index           =   6
         Left            =   -70590
         TabIndex        =   227
         Top             =   3960
         Width           =   690
      End
      Begin VB.Label lblPara4 
         Alignment       =   2  'Center
         Caption         =   "Para4"
         Height          =   195
         Index           =   6
         Left            =   -69870
         TabIndex        =   226
         Top             =   3960
         Width           =   690
      End
      Begin VB.Label lblPara1 
         Alignment       =   2  'Center
         Caption         =   "Para1"
         Height          =   195
         Index           =   5
         Left            =   -72060
         TabIndex        =   225
         Top             =   3420
         Width           =   690
      End
      Begin VB.Label lblPara2 
         Alignment       =   2  'Center
         Caption         =   "Para2"
         Height          =   195
         Index           =   5
         Left            =   -71310
         TabIndex        =   224
         Top             =   3420
         Width           =   690
      End
      Begin VB.Label lblPara3 
         Alignment       =   2  'Center
         Caption         =   "Para3"
         Height          =   195
         Index           =   5
         Left            =   -70590
         TabIndex        =   223
         Top             =   3420
         Width           =   690
      End
      Begin VB.Label lblPara4 
         Alignment       =   2  'Center
         Caption         =   "Para4"
         Height          =   195
         Index           =   5
         Left            =   -69870
         TabIndex        =   222
         Top             =   3420
         Width           =   690
      End
      Begin VB.Label lblPara1 
         Alignment       =   2  'Center
         Caption         =   "Para1"
         Height          =   195
         Index           =   4
         Left            =   -72045
         TabIndex        =   221
         Top             =   2880
         Width           =   690
      End
      Begin VB.Label lblPara2 
         Alignment       =   2  'Center
         Caption         =   "Para2"
         Height          =   195
         Index           =   4
         Left            =   -71295
         TabIndex        =   220
         Top             =   2880
         Width           =   690
      End
      Begin VB.Label lblPara3 
         Alignment       =   2  'Center
         Caption         =   "Para3"
         Height          =   195
         Index           =   4
         Left            =   -70575
         TabIndex        =   219
         Top             =   2880
         Width           =   690
      End
      Begin VB.Label lblPara4 
         Alignment       =   2  'Center
         Caption         =   "Para4"
         Height          =   195
         Index           =   4
         Left            =   -69855
         TabIndex        =   218
         Top             =   2880
         Width           =   690
      End
      Begin VB.Label lblPara1 
         Alignment       =   2  'Center
         Caption         =   "Para1"
         Height          =   195
         Index           =   3
         Left            =   -72045
         TabIndex        =   217
         Top             =   2340
         Width           =   690
      End
      Begin VB.Label lblPara2 
         Alignment       =   2  'Center
         Caption         =   "Para2"
         Height          =   195
         Index           =   3
         Left            =   -71295
         TabIndex        =   216
         Top             =   2340
         Width           =   690
      End
      Begin VB.Label lblPara3 
         Alignment       =   2  'Center
         Caption         =   "Para3"
         Height          =   195
         Index           =   3
         Left            =   -70575
         TabIndex        =   215
         Top             =   2340
         Width           =   690
      End
      Begin VB.Label lblPara4 
         Alignment       =   2  'Center
         Caption         =   "Para4"
         Height          =   195
         Index           =   3
         Left            =   -69855
         TabIndex        =   214
         Top             =   2340
         Width           =   690
      End
      Begin VB.Label lblPara1 
         Alignment       =   2  'Center
         Caption         =   "Para1"
         Height          =   195
         Index           =   2
         Left            =   -72045
         TabIndex        =   213
         Top             =   1800
         Width           =   690
      End
      Begin VB.Label lblPara2 
         Alignment       =   2  'Center
         Caption         =   "Para2"
         Height          =   195
         Index           =   2
         Left            =   -71295
         TabIndex        =   212
         Top             =   1800
         Width           =   690
      End
      Begin VB.Label lblPara3 
         Alignment       =   2  'Center
         Caption         =   "Para3"
         Height          =   195
         Index           =   2
         Left            =   -70575
         TabIndex        =   211
         Top             =   1800
         Width           =   690
      End
      Begin VB.Label lblPara4 
         Alignment       =   2  'Center
         Caption         =   "Para4"
         Height          =   195
         Index           =   2
         Left            =   -69855
         TabIndex        =   210
         Top             =   1800
         Width           =   690
      End
      Begin VB.Label lblPara1 
         Alignment       =   2  'Center
         Caption         =   "Para1"
         Height          =   195
         Index           =   1
         Left            =   -72045
         TabIndex        =   209
         Top             =   1260
         Width           =   690
      End
      Begin VB.Label lblPara2 
         Alignment       =   2  'Center
         Caption         =   "Para2"
         Height          =   195
         Index           =   1
         Left            =   -71295
         TabIndex        =   208
         Top             =   1260
         Width           =   690
      End
      Begin VB.Label lblPara3 
         Alignment       =   2  'Center
         Caption         =   "Para3"
         Height          =   195
         Index           =   1
         Left            =   -70575
         TabIndex        =   207
         Top             =   1260
         Width           =   690
      End
      Begin VB.Label lblPara4 
         Alignment       =   2  'Center
         Caption         =   "Para4"
         Height          =   195
         Index           =   1
         Left            =   -69855
         TabIndex        =   206
         Top             =   1260
         Width           =   690
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Permanent NPC"
         Height          =   195
         Index           =   20
         Left            =   -74400
         TabIndex        =   204
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label lblPara4 
         Alignment       =   2  'Center
         Caption         =   "Para4"
         Height          =   195
         Index           =   0
         Left            =   -69855
         TabIndex        =   203
         Top             =   720
         Width           =   690
      End
      Begin VB.Label lblPara3 
         Alignment       =   2  'Center
         Caption         =   "Para3"
         Height          =   195
         Index           =   0
         Left            =   -70575
         TabIndex        =   202
         Top             =   720
         Width           =   690
      End
      Begin VB.Label lblPara2 
         Alignment       =   2  'Center
         Caption         =   "Para2"
         Height          =   195
         Index           =   0
         Left            =   -71295
         TabIndex        =   201
         Top             =   720
         Width           =   690
      End
      Begin VB.Label lblPara1 
         Alignment       =   2  'Center
         Caption         =   "Para1"
         Height          =   195
         Index           =   0
         Left            =   -72045
         TabIndex        =   200
         Top             =   720
         Width           =   690
      End
      Begin VB.Label lblRoomNum 
         Alignment       =   2  'Center
         Caption         =   "Room #"
         Height          =   195
         Index           =   0
         Left            =   -74220
         TabIndex        =   199
         Top             =   420
         Width           =   570
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "D"
         Height          =   195
         Index           =   13
         Left            =   -74505
         TabIndex        =   198
         Top             =   5760
         Width           =   120
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "U"
         Height          =   195
         Index           =   12
         Left            =   -74505
         TabIndex        =   197
         Top             =   5220
         Width           =   120
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "SW"
         Height          =   195
         Index           =   10
         Left            =   -74580
         TabIndex        =   196
         Top             =   4740
         Width           =   270
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "SE"
         Height          =   195
         Index           =   9
         Left            =   -74550
         TabIndex        =   195
         Top             =   4140
         Width           =   210
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "NW"
         Height          =   195
         Index           =   8
         Left            =   -74580
         TabIndex        =   194
         Top             =   3600
         Width           =   285
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "NE"
         Height          =   195
         Index           =   7
         Left            =   -74550
         TabIndex        =   193
         Top             =   3060
         Width           =   225
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "W"
         Height          =   195
         Index           =   6
         Left            =   -74520
         TabIndex        =   192
         Top             =   2520
         Width           =   165
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "E"
         Height          =   195
         Index           =   5
         Left            =   -74490
         TabIndex        =   191
         Top             =   1980
         Width           =   105
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "S"
         Height          =   195
         Index           =   4
         Left            =   -74490
         TabIndex        =   190
         Top             =   1440
         Width           =   105
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "N"
         Height          =   195
         Index           =   3
         Left            =   -74505
         TabIndex        =   189
         Top             =   900
         Width           =   120
      End
      Begin VB.Label lblExitType 
         Alignment       =   2  'Center
         Caption         =   "Exit Type"
         Height          =   195
         Index           =   0
         Left            =   -73260
         TabIndex        =   188
         Top             =   420
         Width           =   810
      End
      Begin VB.Label label 
         Caption         =   "Placed Items"
         Height          =   195
         Index           =   26
         Left            =   -74400
         TabIndex        =   169
         Top             =   1560
         Width           =   2835
      End
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "Last"
      Height          =   375
      Left            =   7560
      TabIndex        =   8
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   375
      Left            =   7200
      TabIndex        =   5
      ToolTipText     =   "Previous Record"
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      Height          =   375
      Left            =   6840
      TabIndex        =   6
      ToolTipText     =   "Next Record"
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "First"
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Direct Record Navigation:"
      Height          =   195
      Left            =   6240
      TabIndex        =   458
      Top             =   2640
      Width           =   1935
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
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   495
   End
   Begin VB.Label label 
      Caption         =   "Room"
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
      Index           =   11
      Left            =   4860
      TabIndex        =   21
      Top             =   120
      Width           =   495
   End
   Begin VB.Label label 
      Caption         =   "Map"
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
      Left            =   3480
      TabIndex        =   19
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Map #"
      Height          =   255
      Left            =   6300
      TabIndex        =   187
      Top             =   30
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Room #"
      Height          =   255
      Left            =   6300
      TabIndex        =   186
      Top             =   630
      Width           =   855
   End
End
Attribute VB_Name = "frmRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim sSearchText As String
Dim bSearchStop As Boolean
'Dim bWarnedAboutCopy As Boolean
Public RoomCopy As Boolean
Dim objTooltip As clsToolTip
Public bLoaded As Boolean

Private Sub cmdClearAdvanced_Click()
On Error GoTo error:

txtLight.Text = "0"
txtGangHouseNumber.Text = "0"
txtAttributes.Text = "0"
txtShopNum.Text = "0"
txtSpell.Text = "0"
txtDeathRoom.Text = "0"
txtExitRoom.Text = "0"
txtCmdText.Text = "0"
txtMinIndex.Text = "0"
txtMaxIndex.Text = "0"
txtMaxRegen.Text = "0"
txtDelay.Text = "0"
txtControlRoom.Text = "0"
txtMaxArea.Text = "0"
cmbMonsterType.ListIndex = 0
cmbType.ListIndex = 0
txtAnsiMap.Text = "WCCMAP01.ANS"

Exit Sub
error:
Call HandleError("cmdClearAdvanced_Click")
End Sub

Private Sub cmdGotoFirstLastMapRoom_Click(Index As Integer)
Dim nStatus As Integer, nBtrieveAction As Integer, nMapToFind As Long, nMaxRooms As Long
Dim sText As String
On Error GoTo error:

If bLoaded = True Then saverecord

nMapToFind = Val(txtMap.Text)
If nMapToFind < 1 Then Exit Sub

nStatus = BTRCALL(BSTAT, RoomPosBlock, DBStatDatabuf, Len(Roomdatabuf), 0, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    nMaxRooms = 30000
Else
    DBStatRowToStruct DBStatDatabuf.buf
    nMaxRooms = DBStat.nRecords
End If

If Index = 0 Then
    sText = "First"
    nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then MsgBox "cmdGotoFirstLastMapRoom_Click(), BGETFIRST, Room, Error: " & BtrieveErrorCode(nStatus)
Else
    sText = "Last"
    nStatus = BTRCALL(BGETLAST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then MsgBox "cmdGotoFirstLastMapRoom_Click(), BGETLAST, Room, Error: " & BtrieveErrorCode(nStatus)
End If
If Not nStatus = 0 Then Exit Sub

If Index = 0 Then
    nBtrieveAction = BGETNEXT
Else
    nBtrieveAction = BGETPREVIOUS
End If

Me.Enabled = False
If FormIsLoaded("frmMap") Then frmMap.Enabled = False
If FormIsLoaded("frmMapEditor") Then frmMapEditor.Enabled = False

frmProgressBar.sCaption = "Finding " & sText & " Room on Map"
frmProgressBar.lblCaption = "Finding " & sText & " Room on Map..."
frmProgressBar.cmdCancel.Enabled = False

Call frmProgressBar.SetRange(nMaxRooms)

frmProgressBar.lblNote.Visible = False
frmProgressBar.lblPanel(0).Caption = ""
frmProgressBar.lblPanel(1).Caption = ""
frmProgressBar.Show
DoEvents

frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_MP
frmProgressBar.lblPanel(1).Caption = "Scanning Rooms..."
DoEvents

Do While nStatus = 0
    nStatus = BTRCALL(nBtrieveAction, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    If nStatus = 0 Then
        Call RoomRowToStruct(Roomdatabuf.buf)
        If Roomrec.MapNumber = nMapToFind Then GoTo found
    End If
    
    Call frmProgressBar.IncreaseProgress
    If Not bUseCPU Then DoEvents
Loop

found:
Me.Enabled = True
DoEvents

If nStatus = 0 Then
    frmProgressBar.ProgressBar.Value = frmProgressBar.ProgressBar.Max
    DoEvents

    DispRoomInfo Roomdatabuf.buf
    If cmdNext.Enabled = False Then Call frmMapEditor.StartMapping
Else
    MsgBox "Failed to find room, Error: " & BtrieveErrorCode(nStatus), vbExclamation
End If

out:
On Error Resume Next
Unload frmProgressBar
Me.Enabled = True
If FormIsLoaded("frmMap") Then frmMap.Enabled = True
If FormIsLoaded("frmMapEditor") Then frmMapEditor.Enabled = True
Exit Sub
error:
Call HandleError("cmdGotoFirstLastMapRoom_Click")
Resume out:
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim nStatus As Integer
bLoaded = False

Set objTooltip = New clsToolTip
With objTooltip
    .Style = ttStyleStandard
    .DelayTime = 200
    .VisibleTime = 25000
    .BkColor = &HC0FFFF
    .txtColor = &H0
    .TipWidth = 200
End With

Me.Top = ReadINI("Windows", "RoomTop")
Me.Left = ReadINI("Windows", "RoomLeft")

nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "LoadRoom, BGETFIRST, Room, Error: " & BtrieveErrorCode(nStatus)
Else
    bLoaded = True

    RoomKeyStruct.MapNum = Val(ReadINI("Options", "LastMap"))
    RoomKeyStruct.RoomNum = Val(ReadINI("Options", "LastRoom"))
    
    nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    
    DispRoomInfo Roomdatabuf.buf
End If



Me.Show
Me.SetFocus
txtGotoMap.SetFocus

End Sub

Private Sub cmbRoomType_Click(Index As Integer)

lblPara1(Index).Caption = ""
lblPara2(Index).Caption = ""
lblPara3(Index).Caption = ""
lblPara4(Index).Caption = ""
objTooltip.DelToolTip txtRoomPara(Index).hwnd
objTooltip.DelToolTip txtRoomWPara(Index).hwnd
objTooltip.DelToolTip txtRoomLPara1(Index).hwnd
objTooltip.DelToolTip txtRoomLPara2(Index).hwnd
        
Select Case cmbRoomType(Index).ListIndex
    Case 0: 'Normal

    Case 1: 'Spell
        lblPara1(Index).Caption = "opn spell"
        lblPara2(Index).Caption = "type"
        lblPara3(Index).Caption = "msg"
        lblPara4(Index).Caption = "cls spell"
        objTooltip.SetToolTipObj txtRoomPara(Index).hwnd, "spell to open", False
        objTooltip.SetToolTipObj txtRoomWPara(Index).hwnd, "1) Open" & vbCrLf & "2) Closed" & vbCrLf & "3) Passthrough" & vbCrLf & "4) Perm Open" & vbCrLf & "5) One Time Closed", False
        objTooltip.SetToolTipObj txtRoomLPara1(Index).hwnd, "message block", False
        objTooltip.SetToolTipObj txtRoomLPara2(Index).hwnd, "spell to close", False
    
    Case 2: 'Key
        lblPara1(Index).Caption = "item #"
        lblPara2(Index).Caption = "status"
        lblPara3(Index).Caption = "diff."
        lblPara4(Index).Caption = "time"
        objTooltip.SetToolTipObj txtRoomPara(Index).hwnd, "item #", False
        objTooltip.SetToolTipObj txtRoomWPara(Index).hwnd, "Status - always locked'?", False
        objTooltip.SetToolTipObj txtRoomLPara1(Index).hwnd, "Picklock difficulty", False
        objTooltip.SetToolTipObj txtRoomLPara2(Index).hwnd, "Number of 5 minute blocks to stay open", False
    
    Case 3: 'Item
        lblPara1(Index).Caption = "item"
        lblPara2(Index).Caption = "msg-no"
        lblPara3(Index).Caption = "msg-ok"
        objTooltip.SetToolTipObj txtRoomPara(Index).hwnd, "item #", False
        objTooltip.SetToolTipObj txtRoomWPara(Index).hwnd, "Message on failed passage", False
        objTooltip.SetToolTipObj txtRoomLPara1(Index).hwnd, "Message on passage", False
        
    Case 4: 'Toll
        lblPara1(Index).Caption = "cost"
        objTooltip.SetToolTipObj txtRoomPara(Index).hwnd, "Amount in gold which will be charged", False

    Case 5: 'Action
        lblPara1(Index).Caption = "msg"
        objTooltip.SetToolTipObj txtRoomPara(Index).hwnd, "message displayed", False
    
    Case 6: 'Hidden
        lblPara1(Index).Caption = "type"
        lblPara2(Index).Caption = "#actions"
        lblPara3(Index).Caption = "msg"
        lblPara4(Index).Caption = "msg"
        objTooltip.SetToolTipObj txtRoomPara(Index).hwnd, _
            "1) Hidden but Passable" & vbCrLf _
            & "2) Needs Search" & vbCrLf _
            & "4) Has Been Searched" & vbCrLf _
            & "8) All Actions Performed (see below)" & vbCrLf _
            & vbCrLf _
            & "Values >= 16 for multiple actions:" & vbCrLf _
            & "Enter 16*((2^n)-1) where n == total number of actions" & vbCrLf _
            & vbCrLf _
            & "(1 exit == 16, 2 exits == 48, 3 == 112, ... 6 == 1008, max 9)" & vbCrLf _
            & vbCrLf _
            & "NOTE: these values change dynamicly as you perform the actions.", False
        objTooltip.SetToolTipObj txtRoomWPara(Index).hwnd, "# actions needed to activate, negative number means order doesn't matter", False
        objTooltip.SetToolTipObj txtRoomLPara1(Index).hwnd, "What the room sees when the passage opens", False
        objTooltip.SetToolTipObj txtRoomLPara2(Index).hwnd, "What you see as the exit when you look", False
        
    Case 7: 'Door
        lblPara1(Index).Caption = "state"
        lblPara2(Index).Caption = "chance"
        lblPara3(Index).Caption = "time"
        lblPara4(Index).Caption = "key"
        objTooltip.SetToolTipObj txtRoomPara(Index).hwnd, "State, 1) always locked", False
        objTooltip.SetToolTipObj txtRoomWPara(Index).hwnd, "Chance to picklock/bash, lower the number harder the chance", False
        objTooltip.SetToolTipObj txtRoomLPara1(Index).hwnd, "Number of 5 minute blocks to stay open", False
        objTooltip.SetToolTipObj txtRoomLPara2(Index).hwnd, "Key required", False
        
    Case 8: 'map Change
        lblPara1(Index).Caption = "map"
        objTooltip.SetToolTipObj txtRoomPara(Index).hwnd, "map number", False
        
    Case 9: 'Trap
        lblPara1(Index).Caption = "damage"
        lblPara2(Index).Caption = "type"
        lblPara3(Index).Caption = "pass msg"
        lblPara4(Index).Caption = "fail msg"
        objTooltip.SetToolTipObj txtRoomPara(Index).hwnd, "Max Damage", False
        objTooltip.SetToolTipObj txtRoomWPara(Index).hwnd, _
            "1. Active" & vbCrLf _
            & "2. Inactive" & vbCrLf _
            & "3. Hidden" & vbCrLf _
            & "4. Active mov" & vbCrLf _
            & "5. Inactive m", False
        objTooltip.SetToolTipObj txtRoomLPara1(Index).hwnd, "Message when trap goes off", False
        objTooltip.SetToolTipObj txtRoomLPara2(Index).hwnd, "Message on disarm fail", False
        
    Case 10: 'Text
        lblPara1(Index).Caption = "msg"
        lblPara2(Index).Caption = "msg"
        lblPara3(Index).Caption = "msg"
        objTooltip.SetToolTipObj txtRoomPara(Index).hwnd, "msg with commands", False
        objTooltip.SetToolTipObj txtRoomWPara(Index).hwnd, _
            "-(line 1) What the user sees" & vbCrLf _
            & "-(line 2) What the room left from sees" & vbCrLf _
            & "-(line 3) What the room entered into sees", False
        objTooltip.SetToolTipObj txtRoomLPara1(Index).hwnd, "message returned", False
        
    Case 11: 'Gate
        lblPara1(Index).Caption = "state"
        lblPara2(Index).Caption = "chance"
        lblPara3(Index).Caption = "time"
        lblPara4(Index).Caption = "key"
        objTooltip.SetToolTipObj txtRoomPara(Index).hwnd, "State, 1) always locked", False
        objTooltip.SetToolTipObj txtRoomWPara(Index).hwnd, "Chance to picklock/bash", False
        objTooltip.SetToolTipObj txtRoomLPara1(Index).hwnd, "Number of 5 minute blocks to stay open", False
        objTooltip.SetToolTipObj txtRoomLPara2(Index).hwnd, "Key required", False
        
    Case 12: 'Remote Action
        lblPara1(Index).Caption = "msg"
        lblPara2(Index).Caption = "#"
        lblPara3(Index).Caption = "msg"
        lblPara4(Index).Caption = "item"
        objTooltip.SetToolTipObj txtRoomPara(Index).hwnd, _
            "-(line 1) Remote Action word one" & vbCrLf _
            & "-(line 2) Remote Action word two" & vbCrLf _
            & "-(line 3) Remote Action word three", False
        objTooltip.SetToolTipObj txtRoomWPara(Index).hwnd, "# = the exit number (0 to 9 for North though Down) that this action effects." _
            & vbCrLf & vbCrLf & "-if there are 2 or more actions needed then add 10 starting with the first action to each remote action's para2." _
            & vbCrLf & vbCrLf & "-so a room that needed 3 remote actions on the South (1) exit would have in para2: 11 for the first action, 21 for the second, and 31 for the third." _
            & vbCrLf & vbCrLf & "-a room that needed only one action on the NE (5) exit would just need a 5 in the para2 field.", False
        
'            & vbCrLf _
'            & "1. north" & vbCrLf _
'            & "2. south" & vbCrLf _
'            & "3. east" & vbCrLf _
'            & "4. west" & vbCrLf _
'            & "5. northeast" & vbCrLf _
'            & "6. northwest" & vbCrLf _
'            & "7. southeast" & vbCrLf _
'            & "8. southwest" & vbCrLf _
'            & "9. up" & vbCrLf _
'            & "10. down", False
        objTooltip.SetToolTipObj txtRoomLPara1(Index).hwnd, _
            "-(line 1) What the user sees when they do the action" & vbCrLf _
            & "-(line 2) What the room sees (UN)", False
        objTooltip.SetToolTipObj txtRoomLPara2(Index).hwnd, "item required", False
        
    Case 13: 'Class
        lblPara1(Index).Caption = "class-ok"
        lblPara2(Index).Caption = "class-no"
        lblPara3(Index).Caption = "fail msg"
        objTooltip.SetToolTipObj txtRoomPara(Index).hwnd, "Class Allowed", False
        objTooltip.SetToolTipObj txtRoomWPara(Index).hwnd, "Class not allowed", False
        objTooltip.SetToolTipObj txtRoomLPara1(Index).hwnd, _
            "Message on failure" & vbCrLf _
            & vbCrLf _
            & "-(line 1) What the user sees when stopped" & vbCrLf _
            & "-(line 2) What the room sees", False
        
    Case 14: 'Race
        lblPara1(Index).Caption = "race-ok"
        lblPara2(Index).Caption = "race-no"
        lblPara3(Index).Caption = "fail msg"
        objTooltip.SetToolTipObj txtRoomPara(Index).hwnd, "Race Allowed", False
        objTooltip.SetToolTipObj txtRoomWPara(Index).hwnd, "Race Not Allowed", False
        objTooltip.SetToolTipObj txtRoomLPara1(Index).hwnd, _
            "Message on failure" & vbCrLf _
            & vbCrLf _
            & "-(line 1) What the user sees when stopped" & vbCrLf _
            & "-(line 2) What the room sees", False
        
    Case 15: 'Level
        lblPara1(Index).Caption = "> LVL"
        lblPara2(Index).Caption = "< LVL"
        lblPara3(Index).Caption = "fail msg"
        objTooltip.SetToolTipObj txtRoomPara(Index).hwnd, "Min Level", False
        objTooltip.SetToolTipObj txtRoomWPara(Index).hwnd, "Max Level", False
        objTooltip.SetToolTipObj txtRoomLPara1(Index).hwnd, _
            "Message on failure" & vbCrLf _
            & vbCrLf _
            & "-(line 1) What the user sees when stopped" & vbCrLf _
            & "-(line 2) What the room sees", False
        
    Case 16: 'Timed
        lblPara1(Index).Caption = "msg"
        lblPara2(Index).Caption = "time"
        lblPara3(Index).Caption = "status"
        objTooltip.SetToolTipObj txtRoomPara(Index).hwnd, _
            "-(line 1) What the room sees on opening (%%s)=(direction)" & vbCrLf _
            & "-(line 2) What the room sees on closing (%%s)=(direction)" & vbCrLf _
            & "-(line 3) Msg to user when moving into closed exit (%%s)=(direction)", False
        objTooltip.SetToolTipObj txtRoomWPara(Index).hwnd, "Number of 5 minute cycles", False
        objTooltip.SetToolTipObj txtRoomLPara1(Index).hwnd, "Status (open=0/closed=1)", False
        
    Case 17: 'Ticket
        lblPara1(Index).Caption = "Item"
        lblPara2(Index).Caption = "fail msg"
        lblPara3(Index).Caption = "ok msg"
        objTooltip.SetToolTipObj txtRoomPara(Index).hwnd, "Item Required", False
        objTooltip.SetToolTipObj txtRoomWPara(Index).hwnd, "Message on failed passage", False
        objTooltip.SetToolTipObj txtRoomLPara1(Index).hwnd, "Message on passage", False
        
    Case 18: 'User Count
        lblPara1(Index).Caption = "max"
        lblPara2(Index).Caption = "current"
        objTooltip.SetToolTipObj txtRoomPara(Index).hwnd, "Max Users allowed through", False
        objTooltip.SetToolTipObj txtRoomWPara(Index).hwnd, "Current count of users", False
        
    Case 19: 'Block Guard
        
    Case 20: 'Alignment
        lblPara1(Index).Caption = "> align"
        lblPara2(Index).Caption = "< align"
        objTooltip.SetToolTipObj txtRoomPara(Index).hwnd, "Minimum Alignment Required", False
        objTooltip.SetToolTipObj txtRoomWPara(Index).hwnd, "Maximum Alignment Allowed", False
        
    Case 21: 'Delay
        
    Case 22: 'Cast
        lblPara1(Index).Caption = "pre spell"
        lblPara2(Index).Caption = "post spell"
        lblPara3(Index).Caption = "pass msg"
        lblPara4(Index).Caption = "look msg"
        objTooltip.SetToolTipObj txtRoomPara(Index).hwnd, "Spell before move", False
        objTooltip.SetToolTipObj txtRoomWPara(Index).hwnd, "Spell after move", False
        objTooltip.SetToolTipObj txtRoomLPara1(Index).hwnd, "Message on passage", False
        objTooltip.SetToolTipObj txtRoomLPara2(Index).hwnd, "Msg when you look through exit (0=normal?)", False
        
    Case 23: 'Ability
        lblPara1(Index).Caption = "abil"
        lblPara2(Index).Caption = "> Value"
        lblPara3(Index).Caption = "< Value"
        lblPara4(Index).Caption = "fail msg"
        objTooltip.SetToolTipObj txtRoomPara(Index).hwnd, "Ability Required", False
        objTooltip.SetToolTipObj txtRoomWPara(Index).hwnd, "Minimum value", False
        objTooltip.SetToolTipObj txtRoomLPara1(Index).hwnd, "Maximum Value", False
        objTooltip.SetToolTipObj txtRoomLPara2(Index).hwnd, "Message on Failure", False
        
    Case 24: 'Spell Trap
        lblPara1(Index).Caption = "spell"
        lblPara2(Index).Caption = "type"
        lblPara3(Index).Caption = "fail msg"
        lblPara4(Index).Caption = "pass msg"
        objTooltip.SetToolTipObj txtRoomPara(Index).hwnd, "Spell Number", False
        objTooltip.SetToolTipObj txtRoomWPara(Index).hwnd, _
            "1. Active" & vbCrLf _
            & "2. Inactive" & vbCrLf _
            & "3. Hidden" & vbCrLf _
            & "4. Active move" & vbCrLf _
            & "5. Inactive move", False
        objTooltip.SetToolTipObj txtRoomLPara1(Index).hwnd, "Message on disarm fail", False
        objTooltip.SetToolTipObj txtRoomLPara2(Index).hwnd, "Message when trap goes off on passage", False
        
End Select

End Sub

Private Sub cmdAttributesQ_Click()
Dim str As String

str = "The attributes field is a bitmask, you can turn certain switches on and off and convert the binary bitmask to a decimal number." & vbCrLf
str = str & vbCrLf
str = str & "RM_PROTECTED  = 00000001;     { PVP-combat is not allowed in this room }" & vbCrLf
str = str & "RM_PATROL     = 00000010;     { Patrollable (exact effects unknown) }" & vbCrLf
str = str & "RM_OWNABLE    = 00000100;     { 'BUY ROOM' executes in this room }" & vbCrLf
str = str & "RM_BIT_4      = 00001000;     { Unknown }" & vbCrLf
str = str & "RM_CLEAR      = 00010000;     { Clears room at cleanup? }" & vbCrLf
str = str & "RM_BIT_6      = 00100000;     { Unknown }" & vbCrLf
str = str & "RM_BIT_7      = 01000000;     { Ganghouse room }" & vbCrLf
str = str & "RM_BIT_8      = 10000000;     { Unknown }" & vbCrLf
str = str & vbCrLf
str = str & "example:" & vbCrLf
str = str & "00000001 = 1 = RM_PROTECTED" & vbCrLf
str = str & "00000010 = 2 = RM_PATROL" & vbCrLf
str = str & "00000011 = 3 = RM_PROTECTED + RM_PATROL" & vbCrLf
str = str & "00001011 = 11 = RM_PROTECTED + RM_PATROL + RM_BIT_4" & vbCrLf
str = str & "00001100 = 12 = RM_BIT_4 + RM_OWNABLE" & vbCrLf
str = str & "00001101 = 13 = RM_PROTECTED + RM_BIT_4 + RM_OWNABLE" & vbCrLf
str = str & "00001111 = 15 = RM_PROTECTED + RM_PATROL + RM_BIT_4 + RM_OWNABLE" & vbCrLf
str = str & "00010000 = 16 = RM_CLEAR" & vbCrLf

MsgBox str

End Sub

Private Sub cmdClear_Click()
Dim x As Integer

For x = 0 To 6
    txtDesc(x).Text = ""
Next x

End Sub

Private Sub cmdCopy2Exist_Click()
Dim nStatus As Integer
Dim nNewMapNumber As Variant
Dim nNewRoomNumber As Variant

On Error GoTo error:

nNewMapNumber = InputBox("Map Number to copy to:", "Copy current room to existing record", txtMap.Text)
If nNewMapNumber = "" Then Exit Sub
nNewRoomNumber = InputBox("Room Number to copy to:", "Copy current room to existing record", txtRoom.Text + 1)
If nNewRoomNumber = "" Then Exit Sub

Call saverecord

RoomKeyStruct.MapNum = Val(nNewMapNumber)
RoomKeyStruct.RoomNum = Val(nNewRoomNumber)

nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "BGETEQUAL, Room, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

    Call RoomRowToStruct(Roomdatabuf.buf)

    FormValuesToRecord
    
    Roomrec.MapNumber = Val(nNewMapNumber)
    Roomrec.RoomNumber = Val(nNewRoomNumber)

UpdateRoomRecord

If cmdNext.Enabled = False Then Call frmMapEditor.StartMapping

Exit Sub
error:
Call HandleError("cmdCopy2Exist_Click")

End Sub

Private Sub cmdCopy2New_Click()
Dim nStatus As Integer
Dim nNewMapNumber As Variant
Dim nNewRoomNumber As Variant

On Error GoTo error:

nNewMapNumber = InputBox("New Map Number:", "Copy current room to new record", txtMap.Text)
If nNewMapNumber = "" Then Exit Sub
nNewRoomNumber = InputBox("New Room Number:", "Copy current room to new record", txtRoom.Text + 1)
If nNewRoomNumber = "" Then Exit Sub
    
    saverecord
    
    FormValuesToRecord
    
    Roomrec.MapNumber = Val(nNewMapNumber)
    Roomrec.RoomNumber = Val(nNewRoomNumber)
    
    RoomStructToRow Roomdatabuf.buf
    
    nStatus = BTRCALL(BINSERT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "cmdInsert, BINSERT, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    Else
        DispRoomInfo Roomdatabuf.buf
        If cmdNext.Enabled = False Then Call frmMapEditor.StartMapping
    End If


Exit Sub
error:
Call HandleError("cmdCopy2New_Click")

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

Private Sub cmdDelete_Click()
Dim nStatus As Integer
Dim nDelete As Integer

On Error GoTo error:

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
saverecord

nDelete = MsgBox("Delete this record from database?", vbYesNo, "Delete Record?")

If nDelete = 6 Then
    nStatus = BTRCALL(BDELETE, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "cmdDelete, BDELETE, Error: " & BtrieveErrorCode(nStatus)
    Else
        Form_Load
        If cmdNext.Enabled = False Then Call frmMapEditor.StartMapping
    End If
End If

Exit Sub
error:
Call HandleError("cmdDelete_Click")

End Sub

Private Sub cmdDiscard_Click()
Dim nStatus As Integer

On Error GoTo error:

RoomKeyStruct.MapNum = Val(txtMap.Text)
RoomKeyStruct.RoomNum = Val(txtRoom.Text)

nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdGoto_Click(), BGETEQUAL, Room, Error: " & BtrieveErrorCode(nStatus)
Else
    DispRoomInfo Roomdatabuf.buf
    If cmdNext.Enabled = False Then Call frmMapEditor.StartMapping
End If

Exit Sub
error:
Call HandleError("cmdDiscard_Click")

End Sub

Private Sub cmdEditCmdText_Click()
    Call frmTextblock.GotoTB(Val(txtCmdText.Text))

End Sub

Private Sub cmdEditControlRoom_Click()
    Call GotoRoom(Val(txtMap.Text), Val(txtControlRoom.Text))
End Sub

Private Sub cmdEditDeathRoom_Click()
    Call GotoRoom(Val(txtMap.Text), Val(txtDeathRoom.Text))
End Sub

Private Sub cmdEditExitRoom_Click()
    Call GotoRoom(Val(txtMap.Text), Val(txtExitRoom.Text))
End Sub

Private Sub cmdEditPermNPC_Click()
Call frmMonster.GotoMonster(Val(txtPermNPC.Text))
frmMonster.Show
frmMonster.SetFocus
End Sub

Private Sub cmdEditPlacedItem_Click(Index As Integer)
Call frmItem.GotoItem(Val(txtPlacedItems(Index).Text))
frmItem.Show
frmItem.SetFocus
End Sub

Private Sub cmdEditShop_Click()
Call frmShop.GotoShop(Val(txtShopNum.Text))
frmShop.Show
frmShop.SetFocus
End Sub

Private Sub cmdEditSpell_Click()
Call frmSpell.GotoSpell(Val(txtSpell.Text))
frmSpell.Show
frmSpell.SetFocus
End Sub

Private Sub cmdFind_Click(Index As Integer)
On Error GoTo error:
Dim nStatus As Integer, x As Integer, sTemp As String

bSearchStop = False

Select Case Index
    Case 0: 'find
        sTemp = InputBox("This will search the room names and room descriptions for the text you enter.", "Search Rooms for String", sSearchText)
        If sTemp = "" Then Exit Sub
        
        sSearchText = sTemp

        nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            MsgBox "BGETFIRST, Room, Error: " & BtrieveErrorCode(nStatus)
            Exit Sub
        End If
        
    Case 1: 'find next
        If sSearchText = "" Then
            sTemp = InputBox("This will search the room name and room description for the text you enter.", "Search Rooms for String", sSearchText)
            If sTemp = "" Then Exit Sub
            
            sSearchText = sTemp
        End If
        
        RoomKeyStruct.MapNum = Val(txtMap.Text)
        RoomKeyStruct.RoomNum = Val(txtRoom.Text)
        
        nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            MsgBox "BGETEQUAL, Map " & txtMap.Text & ", Room " & txtRoom.Text & ", Error: " & BtrieveErrorCode(nStatus)
            Exit Sub
        End If
        
        nStatus = BTRCALL(BGETNEXT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
        If Not nStatus = 0 Then
            If nStatus = 9 Then
                MsgBox "Not Found."
                Exit Sub
            Else
                MsgBox "BGETNEXT, Room, Error: " & BtrieveErrorCode(nStatus)
            End If
        End If
End Select
        
frmProgressBar.sCaption = "Search Rooms for String"
frmProgressBar.lblCaption.Caption = "Searching Rooms ..."
frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_MP
frmProgressBar.cmdCancel.Enabled = True
frmProgressBar.ProgressBar.Value = 0
Set frmProgressBar.FormOwner = Me

nStatus = BTRCALL(BSTAT, RoomPosBlock, DBStatDatabuf, Len(Roomdatabuf), 0, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    Call frmProgressBar.SetRange(20000)
Else
    DBStatRowToStruct DBStatDatabuf.buf
    Call frmProgressBar.SetRange(DBStat.nRecords)
End If

frmProgressBar.Show
frmMain.Enabled = False
DoEvents

Do While nStatus = 0 And bSearchStop = False
    RoomRowToStruct Roomdatabuf.buf
    frmProgressBar.lblPanel(1).Caption = Roomrec.RoomNumber
    
    If Not InStr(1, UCase(Roomrec.Name), UCase(sSearchText)) = 0 Then Exit Do
    
    For x = 0 To 6
        If Not InStr(1, UCase(Roomrec.Desc(x)), UCase(sSearchText)) = 0 Then Exit Do
    Next x
    
    nStatus = BTRCALL(BGETNEXT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    frmProgressBar.IncreaseProgress
    If Not bUseCPU Then DoEvents
Loop

frmMain.Enabled = True
Unload frmProgressBar

If bSearchStop = True Then Exit Sub

Select Case nStatus
    Case 0:
        Call DispRoomInfo(Roomdatabuf.buf)
        If cmdNext.Enabled = False Then Call frmMapEditor.StartMapping

    Case 9:
        MsgBox "Not found."

End Select

Exit Sub

error:
Call HandleError
frmMain.Enabled = True
Unload frmProgressBar
End Sub

Private Sub cmdFirst_Click()
Dim nStatus As Integer
On Error GoTo error:

If bLoaded = True Then saverecord

nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdFirst_Click, BGETFIRST, Room, Error: " & BtrieveErrorCode(nStatus)
Else
    DispRoomInfo Roomdatabuf.buf
End If

Exit Sub
error:
Call HandleError("cmdFirst_Click")

End Sub

Public Sub cmdGoto_Click()
Dim nStatus As Integer
On Error GoTo error:

If bLoaded = True Then saverecord

RoomKeyStruct.MapNum = Val(txtGotoMap.Text)
RoomKeyStruct.RoomNum = Val(txtGotoRoom.Text)

nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdGoto_Click(), BGETEQUAL, Room, Error: " & BtrieveErrorCode(nStatus)
Else
    DispRoomInfo Roomdatabuf.buf
    If cmdNext.Enabled = False Then Call frmMapEditor.StartMapping
End If

Exit Sub
error:
Call HandleError("cmdGoto_Click")

End Sub

Private Sub cmdGotoRoom_Click(Index As Integer)

If Val(txtRoomExit(Index).Text) = 0 Then Exit Sub


If cmbRoomType(Index).ListIndex = 8 Then
    Call GotoRoom(Val(txtRoomPara(Index).Text), Val(txtRoomExit(Index).Text))
Else
    Call GotoRoom(Val(txtMap.Text), Val(txtRoomExit(Index).Text))
End If

End Sub

Private Sub cmdHiddenGoto_Click(Index As Integer)
Call frmItem.GotoItem(Val(txtHiddenNumber(Index).Text))
frmItem.Show
frmItem.SetFocus
End Sub

Private Sub cmdInsert_Click()
Dim nStatus As Integer
Dim nNewMapNumber As Variant, nNewRoomNumber As Variant

On Error GoTo error:

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
If bLoaded = True Then saverecord

nNewMapNumber = InputBox("New Map Number:", "Insert", Val(txtMap.Text))
If nNewMapNumber = "" Then Exit Sub
nNewRoomNumber = InputBox("New Room Number:", "Insert", Val(txtRoom.Text) + 1)
If nNewRoomNumber = "" Then Exit Sub

Dim a As Integer
For a = 1 To Len(Roomdatabuf)
    Roomdatabuf.buf(a) = &H0
Next

RoomRowToStruct Roomdatabuf.buf

Roomrec.MapNumber = Val(nNewMapNumber)
Roomrec.RoomNumber = Val(nNewRoomNumber)
Roomrec.Name = "New Room" & Chr(0)
Roomrec.AnsiMap = "WCCMAP01.ANS" & Chr(0)

RoomStructToRow Roomdatabuf.buf

nStatus = BTRCALL(BINSERT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdInsert, BINSERT, Error: " & BtrieveErrorCode(nStatus)
Else
    DispRoomInfo Roomdatabuf.buf
    If cmdNext.Enabled = False Then Call frmMapEditor.StartMapping
End If

Exit Sub
error:
Call HandleError("cmdInsert_Click")

End Sub

Private Sub cmdLast_Click()
Dim nStatus As Integer
On Error GoTo error:

If bLoaded = True Then saverecord

nStatus = BTRCALL(BGETLAST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdLast_Click, BGETLAST, Room, Error: " & BtrieveErrorCode(nStatus)
Else
    DispRoomInfo Roomdatabuf.buf
End If

Exit Sub
error:
Call HandleError("cmdLast_Click")

End Sub

Private Sub cmdMap_Click()
Dim oForm As Form
On Error GoTo error:

On Error Resume Next
If bLoaded = True Then saverecord

For Each oForm In Forms
    If oForm.Name = frmMap.Name Then
        If frmMap.WindowState = vbMinimized Then
            Unload frmMap
            Load frmMap
        Else
            Call frmMap.cmdReload_Click
        End If
        Set oForm = Nothing
        Exit Sub
    End If
    Set oForm = Nothing
Next
Load frmMap

Exit Sub
error:
Call HandleError("cmdMap_Click")
End Sub

Private Sub cmdMapEditor_Click()

On Error GoTo error:

If bLoaded = True Then saverecord

frmMapEditor.StartMapping
frmMapEditor.SetFocus

Exit Sub
error:
Call HandleError("cmdMapEditor_Click")
End Sub

Private Sub cmdMinusOneRoom_Click()
Dim nStatus As Integer
On Error GoTo error:

If bLoaded = True Then saverecord

RoomKeyStruct.MapNum = Val(txtMap.Text)
RoomKeyStruct.RoomNum = Val(txtRoom.Text) - 1

nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdGoto_Click(), BGETEQUAL, Room, Error: " & BtrieveErrorCode(nStatus)
Else
    DispRoomInfo Roomdatabuf.buf
    If cmdNext.Enabled = False Then Call frmMapEditor.StartMapping
End If



Exit Sub
error:
Call HandleError("cmdMinusOneRoom_Click")

End Sub

Private Sub cmdNext_Click()
Dim nStatus As Integer
On Error GoTo error:

If bLoaded = True Then saverecord

nStatus = BTRCALL(BGETNEXT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdNext_Click, BGETNEXT, Room, Error: " & BtrieveErrorCode(nStatus)
Else
    DispRoomInfo Roomdatabuf.buf
End If

Exit Sub
error:
Call HandleError("cmdNext_Click")

End Sub

Private Sub cmdPlusOneRoom_Click()
Dim nStatus As Integer
On Error GoTo error:

If bLoaded = True Then saverecord

RoomKeyStruct.MapNum = Val(txtMap.Text)
RoomKeyStruct.RoomNum = Val(txtRoom.Text) + 1

nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdGoto_Click(), BGETEQUAL, Room, Error: " & BtrieveErrorCode(nStatus)
Else
    DispRoomInfo Roomdatabuf.buf
    If cmdNext.Enabled = False Then Call frmMapEditor.StartMapping
End If



Exit Sub
error:
Call HandleError("cmdPlusOneRoom_Click")

End Sub

Private Sub cmdPrevious_Click()
Dim nStatus As Integer
On Error GoTo error:

If bLoaded = True Then saverecord

nStatus = BTRCALL(BGETPREVIOUS, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdPrevious_Click, BGETPREVIOUS, Room, Error: " & BtrieveErrorCode(nStatus)
Else
    DispRoomInfo Roomdatabuf.buf
End If

Exit Sub
error:
Call HandleError("cmdPrevious_Click")
        
End Sub

Private Sub cmdSave_Click()
Dim nStatus As Integer

On Error GoTo error:

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
Call saverecord

RoomKeyStruct.MapNum = Val(txtMap.Text)
RoomKeyStruct.RoomNum = Val(txtRoom.Text)

nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdGoto_Click(), BGETEQUAL, Room, Error: " & BtrieveErrorCode(nStatus)
Else
    DispRoomInfo Roomdatabuf.buf
End If

If cmdNext.Enabled = False Then Call frmMapEditor.StartMapping

Exit Sub
error:
Call HandleError("cmdSave_Click")

End Sub

Private Sub cmdViewIndex_Click()
If FormIsLoaded("frmMonsterIndex") Then
    frmMonsterIndex.Show
    frmMonsterIndex.SetFocus
Else
    frmMonsterIndex.Show
End If
End Sub

Private Sub cmdVisibleGoto_Click(Index As Integer)
Call frmItem.GotoItem(Val(txtVisibleNumber(Index).Text))
frmItem.Show
frmItem.SetFocus
End Sub

Public Sub DispRoomInfo(row() As Byte)
On Error GoTo error:
bLoaded = True
Dim x As Integer, sPrefix As String

Call RoomRowToStruct(row())

If RoomCopy Then
    sPrefix = "COPY "
Else
    sPrefix = ""
End If
Me.Caption = "Room Editor " & sPrefix & "-- " & Roomrec.MapNumber & "/" & Roomrec.RoomNumber & " (" & ClipNull(Roomrec.Name) & ")"

If Roomrec.ControlRoom > 0 Then Call AddControlRoom(Roomrec.MapNumber, Roomrec.ControlRoom, Roomrec.RoomNumber)

'Text2.Text = Roomrec.unknown70
txtMap.Text = Roomrec.MapNumber
txtRoom.Text = Roomrec.RoomNumber
txtName.Text = Roomrec.Name
txtAnsiMap.Text = Roomrec.AnsiMap
cmbType.ListIndex = Roomrec.Type
txtShopNum.Text = Roomrec.ShopNum
txtMinIndex.Text = Roomrec.MinIndex
txtMaxIndex.Text = Roomrec.MaxIndex
txtPermNPC.Text = Roomrec.PermNPC
txtLight.Text = Roomrec.Light
cmbMonsterType.ListIndex = Roomrec.MonsterType
txtMaxRegen.Text = Roomrec.MaxRegen
txtDeathRoom.Text = Roomrec.DeathRoom
txtCmdText.Text = Roomrec.CmdText
txtDelay.Text = Roomrec.Delay
txtMaxArea.Text = Roomrec.MaxArea
txtControlRoom.Text = Roomrec.ControlRoom
txtRunic.Text = Roomrec.Runic
txtPlatinum.Text = Roomrec.Platinum
txtGold.Text = Roomrec.Gold
txtSilver.Text = Roomrec.Silver
txtCopper.Text = Roomrec.Copper
txtInvisRunic.Text = Roomrec.InvisRunic
txtInvisPlatinum.Text = Roomrec.InvisPlatinum
txtInvisGold.Text = Roomrec.InvisGold
txtInvisSilver.Text = Roomrec.InvisSilver
txtInvisCopper.Text = Roomrec.InvisCopper
txtSpell.Text = Roomrec.Spell
txtExitRoom.Text = Roomrec.ExitRoom
txtAttributes.Text = Roomrec.Attributes
'txtPermNPCName.Text = GetMonsterName(Roomrec.PermNPC)
'txtCmdTextDisplay.Text = GetTextblock(Roomrec.CmdText)
txtGangHouseNumber.Text = Roomrec.GangHouseNumber
    
For x = 0 To 6
    txtDesc(x).Text = Roomrec.Desc(x)
Next

For x = 0 To 16
    txtVisibleNumber(x).Text = Roomrec.RoomItems(x)
    'txtVisibleName(X).Text = GetItemName(Roomrec.RoomItems(X))
    txtVisibleUses(x).Text = Roomrec.RoomItemUses(x)
    txtVisibleQty(x).Text = Roomrec.RoomItemQty(x)
Next

For x = 0 To 14
    txtHiddenNumber(x).Text = Roomrec.InvisItems(x)
    'txtHiddenName(X).Text = GetItemName(Roomrec.InvisItems(X))
    txtHiddenUses(x).Text = Roomrec.InvisItemUses(x)
    txtHiddenQty(x).Text = Roomrec.InvisItemQty(x)
    txtCurrentRoomMon(x).Text = Roomrec.CurrentRoomMon(x)
Next

For x = 0 To 9
    txtRoomExit(x).Text = Roomrec.RoomExit(x)
    cmbRoomType(x).ListIndex = Roomrec.RoomType(x)
    txtRoomPara(x).Text = Roomrec.Para1(x)
    txtRoomWPara(x).Text = Roomrec.Para2(x)
    txtRoomLPara1(x).Text = Roomrec.Para3(x)
    txtRoomLPara2(x).Text = Roomrec.Para4(x)
    txtPlacedItems(x).Text = Roomrec.PlacedItems(x)
    'txtPlacedItemsName(X).Text = GetItemName(Roomrec.PlacedItems(X))
Next

Exit Sub
error:
Call HandleError
MsgBox "Warning, record was not completely displayed." & vbCrLf _
    & "Previous records stats may still be in memory.  Select 'Disable DB Writing'" & vbCrLf _
    & "from the file menu and then reload the editor.", vbExclamation
End Sub

Public Sub FormValuesToRecord()
On Error GoTo error:
Dim nStatus As Integer, x As Integer

'DoEvents
Roomrec.Name = RTrim(txtName.Text) & Chr(0)

Roomrec.AnsiMap = RTrim(txtAnsiMap.Text) & Chr(0)
Roomrec.Type = cmbType.ListIndex
Roomrec.ShopNum = Val(txtShopNum.Text)
Roomrec.MinIndex = Val(txtMinIndex.Text)
Roomrec.MaxIndex = Val(txtMaxIndex.Text)
Roomrec.PermNPC = Val(txtPermNPC.Text)
Roomrec.Light = Val(txtLight.Text)
Roomrec.MonsterType = cmbMonsterType.ListIndex
Roomrec.MaxRegen = Val(txtMaxRegen.Text)
Roomrec.DeathRoom = Val(txtDeathRoom.Text)
Roomrec.CmdText = Val(txtCmdText.Text)
Roomrec.Delay = Val(txtDelay.Text)
Roomrec.MaxArea = Val(txtMaxArea.Text)
Roomrec.ControlRoom = Val(txtControlRoom.Text)
Roomrec.Runic = Val(txtRunic.Text)
Roomrec.Platinum = Val(txtPlatinum.Text)
Roomrec.Gold = Val(txtGold.Text)
Roomrec.Silver = Val(txtSilver.Text)
Roomrec.Copper = Val(txtCopper.Text)
Roomrec.InvisRunic = Val(txtInvisRunic.Text)
Roomrec.InvisPlatinum = Val(txtInvisPlatinum.Text)
Roomrec.InvisGold = Val(txtInvisGold.Text)
Roomrec.InvisSilver = Val(txtInvisSilver.Text)
Roomrec.InvisCopper = Val(txtInvisCopper.Text)
Roomrec.Spell = Val(txtSpell.Text)
Roomrec.ExitRoom = Val(txtExitRoom.Text)
Roomrec.Attributes = Val(txtAttributes.Text)
Roomrec.GangHouseNumber = Val(txtGangHouseNumber.Text)

For x = 0 To 6
    Roomrec.Desc(x) = Trim(txtDesc(x).Text) & Chr(0)
Next

For x = 0 To 16
    Roomrec.RoomItems(x) = Val(txtVisibleNumber(x).Text)
    Roomrec.RoomItemUses(x) = Val(txtVisibleUses(x).Text)
    Roomrec.RoomItemQty(x) = Val(txtVisibleQty(x).Text)
Next

For x = 0 To 14
    Roomrec.InvisItems(x) = Val(txtHiddenNumber(x).Text)
    Roomrec.InvisItemUses(x) = Val(txtHiddenUses(x).Text)
    Roomrec.InvisItemQty(x) = Val(txtHiddenQty(x).Text)
    
    Roomrec.CurrentRoomMon(x) = Val(txtCurrentRoomMon(x).Text)
Next

For x = 0 To 9
    Roomrec.RoomExit(x) = Val(txtRoomExit(x).Text)
    Roomrec.RoomType(x) = cmbRoomType(x).ListIndex
    Roomrec.Para1(x) = Val(txtRoomPara(x).Text)
    Roomrec.Para2(x) = Val(txtRoomWPara(x).Text)
    Roomrec.Para3(x) = Val(txtRoomLPara1(x).Text)
    Roomrec.Para4(x) = Val(txtRoomLPara2(x).Text)
    Roomrec.PlacedItems(x) = Val(txtPlacedItems(x).Text)
Next

Exit Sub
error:
Call HandleError
End Sub

Public Sub GotoRoom(ByVal nMap As Long, ByVal nRoom As Long, Optional ByVal NoFocus As Boolean)
On Error GoTo error:
Dim nStatus As Integer, nGoto As Long
If bLoaded = True Then
    saverecord
Else
    Load Me
End If

RoomKeyStruct.MapNum = nMap
RoomKeyStruct.RoomNum = nRoom

nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "GotoRoom, BGETEQUAL, Room, Error: " & BtrieveErrorCode(nStatus)
Else
    DispRoomInfo Roomdatabuf.buf
    If cmdNext.Enabled = False Then Call frmMapEditor.StartMapping
End If

If Not NoFocus Then
    Me.Show
    Me.SetFocus
End If

Exit Sub
error:
Call HandleError
End Sub

Public Sub ReloadMap()
Unload frmMap
Call cmdMap_Click
End Sub

Public Sub saverecord()
On Error GoTo error:
Dim nStatus As Integer, x As Integer

RoomKeyStruct.MapNum = Val(txtMap.Text)
RoomKeyStruct.RoomNum = Val(txtRoom.Text)

nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Save Error, BGETEQUAL, Map " & txtMap.Text & ", Room " & txtRoom.Text & ", Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

RoomRowToStruct Roomdatabuf.buf

Call FormValuesToRecord

UpdateRoomRecord

Exit Sub
error:
Call HandleError
End Sub

Public Sub ToggleSearchStop()
bSearchStop = True
End Sub

Private Sub txtAnsiMap_GotFocus()
Call SelectAll(txtAnsiMap)

End Sub

Private Sub txtAttributes_GotFocus()
Call SelectAll(txtAttributes)

End Sub

Private Sub txtCmdText_Change()
On Error GoTo error:

txtCmdTextDisplay.Text = GetTextblock(Val(txtCmdText.Text))

out:
Exit Sub
error:
Call HandleError("txtCmdText_Change")
Resume out:
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

Private Sub txtExitRoom_GotFocus()
Call SelectAll(txtExitRoom)

End Sub

Private Sub txtGangHouseNumber_GotFocus()
Call SelectAll(txtGangHouseNumber)

End Sub

Private Sub txtGold_GotFocus()
Call SelectAll(txtGold)

End Sub

Private Sub txtGotoMap_GotFocus()
Call SelectAll(txtGotoMap)
End Sub

Private Sub txtGotoRoom_GotFocus()
Call SelectAll(txtGotoRoom)
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

Private Sub txtLight_GotFocus()
Call SelectAll(txtLight)

End Sub

Private Sub txtMap_GotFocus()
Call SelectAll(txtMap)

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

Private Sub txtRoom_GotFocus()
Call SelectAll(txtRoom)

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

Private Sub txtSpell_GotFocus()
Call SelectAll(txtSpell)

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

Private Sub UpdateRoomRecord()
Dim nStatus As Integer

nStatus = UpdateRoom
If Not nStatus = 0 Then
    MsgBox "Error Updating room record: " & BtrieveErrorCode(nStatus)
Else
    DispRoomInfo Roomdatabuf.buf
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
        
        If FormIsLoaded("frmMapEditor") Then Unload frmMapEditor
        
        Set objTooltip = Nothing
        
        If bLoaded = True Then saverecord
        Call WriteINI("Options", "LastMap", txtMap.Text)
        Call WriteINI("Options", "LastRoom", txtRoom.Text)
        If Me.WindowState = vbMinimized Then Exit Sub
        Call WriteINI("Windows", "RoomTop", Me.Top)
        Call WriteINI("Windows", "RoomLeft", Me.Left)
        
End Sub



