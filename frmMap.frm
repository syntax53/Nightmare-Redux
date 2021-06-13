VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMap 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map"
   ClientHeight    =   10080
   ClientLeft      =   600
   ClientTop       =   975
   ClientWidth     =   9840
   Icon            =   "frmMap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   672
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   656
   Begin VB.PictureBox picMap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   9855
      Left            =   0
      ScaleHeight     =   657
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   657
      TabIndex        =   5
      Top             =   240
      Width           =   9855
      Begin VB.Frame framOptions 
         Caption         =   "Options"
         Height          =   4755
         Left            =   60
         TabIndex        =   10
         Top             =   60
         Visible         =   0   'False
         Width           =   2355
         Begin VB.CommandButton cmdBuildControlRoomList 
            Caption         =   "Rebuild"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1500
            TabIndex        =   1711
            ToolTipText     =   "Rebuild Control Room List"
            Top             =   2100
            Width           =   795
         End
         Begin VB.OptionButton optMarkAux 
            Caption         =   "Control Rms"
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
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   1710
            Top             =   2100
            Width           =   1695
         End
         Begin VB.OptionButton optMarkAux 
            Caption         =   "Room Spells"
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
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   16
            Top             =   1380
            Width           =   1635
         End
         Begin VB.OptionButton optMarkAux 
            Caption         =   "Shops"
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
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   17
            Top             =   1620
            Width           =   1635
         End
         Begin VB.OptionButton optMarkAux 
            Caption         =   "Exit/Death Rooms"
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
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   18
            Top             =   1860
            Width           =   1695
         End
         Begin VB.CheckBox chkMarkAux 
            Caption         =   "Mark --"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   120
            TabIndex        =   15
            Top             =   1200
            Width           =   975
         End
         Begin VB.CheckBox chkNoColors 
            Caption         =   "No Color"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   1380
            TabIndex        =   26
            Top             =   3900
            Width           =   915
         End
         Begin VB.CheckBox chkFollowMapChanges 
            Caption         =   "Follow Map Changes"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1635
         End
         Begin VB.CheckBox chkDontFollowHidden 
            Caption         =   "Don't Follow Hidden Exits"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   120
            TabIndex        =   20
            Top             =   2640
            Width           =   1995
         End
         Begin VB.CheckBox chkNoLineColors 
            Caption         =   "No Line Color"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   120
            TabIndex        =   25
            Top             =   3900
            Width           =   1395
         End
         Begin VB.CheckBox chkMarkLair 
            Caption         =   "Mark Lair Rooms"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   1755
         End
         Begin VB.CheckBox chkDontMarkStart 
            Caption         =   "Don't Mark Starting Point"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   120
            TabIndex        =   19
            Top             =   2400
            Width           =   1995
         End
         Begin VB.CheckBox chkMarkCMD 
            Caption         =   "Mark Rooms w/Commands"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   2115
         End
         Begin VB.ComboBox cmbMapSize 
            Height          =   315
            ItemData        =   "frmMap.frx":08CA
            Left            =   120
            List            =   "frmMap.frx":08DA
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   4320
            Width           =   2115
         End
         Begin VB.CheckBox chkMarkNPC 
            Caption         =   "Mark Rooms w/Perm NPCs"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   2115
         End
         Begin VB.CheckBox chkUseLastExport 
            Caption         =   "Auto Use Last Export File"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   120
            TabIndex        =   23
            Top             =   3420
            Width           =   2115
         End
         Begin VB.CheckBox chkMonsterRegen 
            Caption         =   "Look up Monster Regen"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   120
            TabIndex        =   22
            Top             =   3120
            Width           =   2055
         End
         Begin VB.CheckBox chkUseWhiteBG 
            Caption         =   "Use White Background"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   120
            TabIndex        =   24
            Top             =   3660
            Width           =   2115
         End
         Begin VB.CheckBox chkNoTooltips 
            Caption         =   "Don't Show Tooltips"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   120
            TabIndex        =   21
            Top             =   2880
            Width           =   1995
         End
         Begin VB.Label Label1 
            Caption         =   "Map Size"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   4140
            Width           =   1215
         End
      End
      Begin VB.Frame framExporting 
         Height          =   1935
         Left            =   600
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   3855
         Begin MSComctlLib.ProgressBar ProgressBar 
            Height          =   375
            Left            =   180
            TabIndex        =   7
            Top             =   1380
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Exporting"
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
            Left            =   180
            TabIndex        =   9
            Top             =   300
            Width           =   3495
         End
         Begin VB.Label lblExportCount 
            Alignment       =   2  'Center
            Caption         =   "1 / 1680"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   300
            TabIndex        =   8
            Top             =   900
            Width           =   3315
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   8940
         Top             =   60
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   0
         X2              =   656
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   4800
         Shape           =   1  'Square
         Tag             =   "30x30"
         Top             =   4800
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   6960
         Shape           =   1  'Square
         Tag             =   "30x30"
         Top             =   6960
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   9600
         Shape           =   1  'Square
         Tag             =   "41 x 41"
         Top             =   9600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   9600
         Shape           =   1  'Square
         Tag             =   "41 x 19"
         Top             =   4320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   40
         Left            =   9420
         TabIndex        =   1709
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   39
         Left            =   9180
         TabIndex        =   1708
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   38
         Left            =   8940
         TabIndex        =   1707
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   37
         Left            =   8700
         TabIndex        =   1706
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   36
         Left            =   8460
         TabIndex        =   1705
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   35
         Left            =   8220
         TabIndex        =   1704
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   34
         Left            =   7980
         TabIndex        =   1703
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   33
         Left            =   7740
         TabIndex        =   1702
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   32
         Left            =   7500
         TabIndex        =   1701
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   31
         Left            =   7260
         TabIndex        =   1700
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   30
         Left            =   7020
         TabIndex        =   1699
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   29
         Left            =   6780
         TabIndex        =   1698
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   28
         Left            =   6540
         TabIndex        =   1697
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   27
         Left            =   6300
         TabIndex        =   1696
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   26
         Left            =   6060
         TabIndex        =   1695
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   25
         Left            =   5820
         TabIndex        =   1694
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   24
         Left            =   5580
         TabIndex        =   1693
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   23
         Left            =   5340
         TabIndex        =   1692
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   22
         Left            =   5100
         TabIndex        =   1691
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   21
         Left            =   4860
         TabIndex        =   1690
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   20
         Left            =   4620
         TabIndex        =   1689
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   19
         Left            =   4380
         TabIndex        =   1688
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   18
         Left            =   4140
         TabIndex        =   1687
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   17
         Left            =   3900
         TabIndex        =   1686
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   16
         Left            =   3660
         TabIndex        =   1685
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   15
         Left            =   3420
         TabIndex        =   1684
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   14
         Left            =   3180
         TabIndex        =   1683
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   13
         Left            =   2940
         TabIndex        =   1682
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   12
         Left            =   2700
         TabIndex        =   1681
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   11
         Left            =   2460
         TabIndex        =   1680
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   10
         Left            =   2220
         TabIndex        =   1679
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   9
         Left            =   1980
         TabIndex        =   1678
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   8
         Left            =   1740
         TabIndex        =   1677
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   7
         Left            =   1500
         TabIndex        =   1676
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   6
         Left            =   1260
         TabIndex        =   1675
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   5
         Left            =   1020
         TabIndex        =   1674
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   4
         Left            =   780
         TabIndex        =   1673
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   3
         Left            =   540
         TabIndex        =   1672
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   2
         Left            =   300
         TabIndex        =   1671
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1
         Left            =   60
         TabIndex        =   1670
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   41
         Left            =   9660
         TabIndex        =   1669
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   42
         Left            =   60
         TabIndex        =   1668
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   43
         Left            =   300
         TabIndex        =   1667
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   44
         Left            =   540
         TabIndex        =   1666
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   45
         Left            =   780
         TabIndex        =   1665
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   46
         Left            =   1020
         TabIndex        =   1664
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   47
         Left            =   1260
         TabIndex        =   1663
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   48
         Left            =   1500
         TabIndex        =   1662
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   49
         Left            =   1740
         TabIndex        =   1661
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   50
         Left            =   1980
         TabIndex        =   1660
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   51
         Left            =   2220
         TabIndex        =   1659
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   52
         Left            =   2460
         TabIndex        =   1658
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   53
         Left            =   2700
         TabIndex        =   1657
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   54
         Left            =   2940
         TabIndex        =   1656
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   55
         Left            =   3180
         TabIndex        =   1655
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   56
         Left            =   3420
         TabIndex        =   1654
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   57
         Left            =   3660
         TabIndex        =   1653
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   58
         Left            =   3900
         TabIndex        =   1652
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   59
         Left            =   4140
         TabIndex        =   1651
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   60
         Left            =   4380
         TabIndex        =   1650
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   61
         Left            =   4620
         TabIndex        =   1649
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   62
         Left            =   4860
         TabIndex        =   1648
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   63
         Left            =   5100
         TabIndex        =   1647
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   64
         Left            =   5340
         TabIndex        =   1646
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   65
         Left            =   5580
         TabIndex        =   1645
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   66
         Left            =   5820
         TabIndex        =   1644
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   67
         Left            =   6060
         TabIndex        =   1643
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   68
         Left            =   6300
         TabIndex        =   1642
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   69
         Left            =   6540
         TabIndex        =   1641
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   70
         Left            =   6780
         TabIndex        =   1640
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   71
         Left            =   7020
         TabIndex        =   1639
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   72
         Left            =   7260
         TabIndex        =   1638
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   73
         Left            =   7500
         TabIndex        =   1637
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   74
         Left            =   7740
         TabIndex        =   1636
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   75
         Left            =   7980
         TabIndex        =   1635
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   76
         Left            =   8220
         TabIndex        =   1634
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   77
         Left            =   8460
         TabIndex        =   1633
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   78
         Left            =   8700
         TabIndex        =   1632
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   79
         Left            =   8940
         TabIndex        =   1631
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   80
         Left            =   9180
         TabIndex        =   1630
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   81
         Left            =   9420
         TabIndex        =   1629
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   82
         Left            =   9660
         TabIndex        =   1628
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   83
         Left            =   60
         TabIndex        =   1627
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   84
         Left            =   300
         TabIndex        =   1626
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   85
         Left            =   540
         TabIndex        =   1625
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   86
         Left            =   780
         TabIndex        =   1624
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   87
         Left            =   1020
         TabIndex        =   1623
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   88
         Left            =   1260
         TabIndex        =   1622
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   89
         Left            =   1500
         TabIndex        =   1621
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   90
         Left            =   1740
         TabIndex        =   1620
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   91
         Left            =   1980
         TabIndex        =   1619
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   92
         Left            =   2220
         TabIndex        =   1618
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   93
         Left            =   2460
         TabIndex        =   1617
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   94
         Left            =   2700
         TabIndex        =   1616
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   95
         Left            =   2940
         TabIndex        =   1615
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   96
         Left            =   3180
         TabIndex        =   1614
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   97
         Left            =   3420
         TabIndex        =   1613
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   98
         Left            =   3660
         TabIndex        =   1612
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   99
         Left            =   3900
         TabIndex        =   1611
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   100
         Left            =   4140
         TabIndex        =   1610
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   101
         Left            =   4380
         TabIndex        =   1609
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   102
         Left            =   4620
         TabIndex        =   1608
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   103
         Left            =   4860
         TabIndex        =   1607
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   104
         Left            =   5100
         TabIndex        =   1606
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   105
         Left            =   5340
         TabIndex        =   1605
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   106
         Left            =   5580
         TabIndex        =   1604
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   107
         Left            =   5820
         TabIndex        =   1603
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   108
         Left            =   6060
         TabIndex        =   1602
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   109
         Left            =   6300
         TabIndex        =   1601
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   110
         Left            =   6540
         TabIndex        =   1600
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   111
         Left            =   6780
         TabIndex        =   1599
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   112
         Left            =   7020
         TabIndex        =   1598
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   113
         Left            =   7260
         TabIndex        =   1597
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   114
         Left            =   7500
         TabIndex        =   1596
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   115
         Left            =   7740
         TabIndex        =   1595
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   116
         Left            =   7980
         TabIndex        =   1594
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   117
         Left            =   8220
         TabIndex        =   1593
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   118
         Left            =   8460
         TabIndex        =   1592
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   119
         Left            =   8700
         TabIndex        =   1591
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   120
         Left            =   8940
         TabIndex        =   1590
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   121
         Left            =   9180
         TabIndex        =   1589
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   122
         Left            =   9420
         TabIndex        =   1588
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   123
         Left            =   9660
         TabIndex        =   1587
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   124
         Left            =   60
         TabIndex        =   1586
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   125
         Left            =   300
         TabIndex        =   1585
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   126
         Left            =   540
         TabIndex        =   1584
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   127
         Left            =   780
         TabIndex        =   1583
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   128
         Left            =   1020
         TabIndex        =   1582
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   129
         Left            =   1260
         TabIndex        =   1581
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   130
         Left            =   1500
         TabIndex        =   1580
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   131
         Left            =   1740
         TabIndex        =   1579
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   132
         Left            =   1980
         TabIndex        =   1578
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   133
         Left            =   2220
         TabIndex        =   1577
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   134
         Left            =   2460
         TabIndex        =   1576
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   135
         Left            =   2700
         TabIndex        =   1575
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   136
         Left            =   2940
         TabIndex        =   1574
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   137
         Left            =   3180
         TabIndex        =   1573
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   138
         Left            =   3420
         TabIndex        =   1572
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   139
         Left            =   3660
         TabIndex        =   1571
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   140
         Left            =   3900
         TabIndex        =   1570
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   141
         Left            =   4140
         TabIndex        =   1569
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   142
         Left            =   4380
         TabIndex        =   1568
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   143
         Left            =   4620
         TabIndex        =   1567
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   144
         Left            =   4860
         TabIndex        =   1566
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   145
         Left            =   5100
         TabIndex        =   1565
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   146
         Left            =   5340
         TabIndex        =   1564
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   147
         Left            =   5580
         TabIndex        =   1563
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   148
         Left            =   5820
         TabIndex        =   1562
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   149
         Left            =   6060
         TabIndex        =   1561
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   150
         Left            =   6300
         TabIndex        =   1560
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   151
         Left            =   6540
         TabIndex        =   1559
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   152
         Left            =   6780
         TabIndex        =   1558
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   153
         Left            =   7020
         TabIndex        =   1557
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   154
         Left            =   7260
         TabIndex        =   1556
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   155
         Left            =   7500
         TabIndex        =   1555
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   156
         Left            =   7740
         TabIndex        =   1554
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   157
         Left            =   7980
         TabIndex        =   1553
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   158
         Left            =   8220
         TabIndex        =   1552
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   159
         Left            =   8460
         TabIndex        =   1551
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   160
         Left            =   8700
         TabIndex        =   1550
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   161
         Left            =   8940
         TabIndex        =   1549
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   162
         Left            =   9180
         TabIndex        =   1548
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   163
         Left            =   9420
         TabIndex        =   1547
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   164
         Left            =   9660
         TabIndex        =   1546
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   165
         Left            =   60
         TabIndex        =   1545
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   166
         Left            =   300
         TabIndex        =   1544
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   167
         Left            =   540
         TabIndex        =   1543
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   168
         Left            =   780
         TabIndex        =   1542
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   169
         Left            =   1020
         TabIndex        =   1541
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   170
         Left            =   1260
         TabIndex        =   1540
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   171
         Left            =   1500
         TabIndex        =   1539
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   172
         Left            =   1740
         TabIndex        =   1538
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   173
         Left            =   1980
         TabIndex        =   1537
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   174
         Left            =   2220
         TabIndex        =   1536
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   175
         Left            =   2460
         TabIndex        =   1535
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   176
         Left            =   2700
         TabIndex        =   1534
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   177
         Left            =   2940
         TabIndex        =   1533
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   178
         Left            =   3180
         TabIndex        =   1532
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   179
         Left            =   3420
         TabIndex        =   1531
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   180
         Left            =   3660
         TabIndex        =   1530
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   181
         Left            =   3900
         TabIndex        =   1529
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   182
         Left            =   4140
         TabIndex        =   1528
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   183
         Left            =   4380
         TabIndex        =   1527
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   184
         Left            =   4620
         TabIndex        =   1526
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   185
         Left            =   4860
         TabIndex        =   1525
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   186
         Left            =   5100
         TabIndex        =   1524
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   187
         Left            =   5340
         TabIndex        =   1523
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   188
         Left            =   5580
         TabIndex        =   1522
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   189
         Left            =   5820
         TabIndex        =   1521
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   190
         Left            =   6060
         TabIndex        =   1520
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   191
         Left            =   6300
         TabIndex        =   1519
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   192
         Left            =   6540
         TabIndex        =   1518
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   193
         Left            =   6780
         TabIndex        =   1517
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   194
         Left            =   7020
         TabIndex        =   1516
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   195
         Left            =   7260
         TabIndex        =   1515
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   196
         Left            =   7500
         TabIndex        =   1514
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   197
         Left            =   7740
         TabIndex        =   1513
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   198
         Left            =   7980
         TabIndex        =   1512
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   199
         Left            =   8220
         TabIndex        =   1511
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   200
         Left            =   8460
         TabIndex        =   1510
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   201
         Left            =   8700
         TabIndex        =   1509
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   202
         Left            =   8940
         TabIndex        =   1508
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   203
         Left            =   9180
         TabIndex        =   1507
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   204
         Left            =   9420
         TabIndex        =   1506
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   205
         Left            =   9660
         TabIndex        =   1505
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   206
         Left            =   60
         TabIndex        =   1504
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   207
         Left            =   300
         TabIndex        =   1503
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   208
         Left            =   540
         TabIndex        =   1502
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   209
         Left            =   780
         TabIndex        =   1501
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   210
         Left            =   1020
         TabIndex        =   1500
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   211
         Left            =   1260
         TabIndex        =   1499
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   212
         Left            =   1500
         TabIndex        =   1498
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   213
         Left            =   1740
         TabIndex        =   1497
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   214
         Left            =   1980
         TabIndex        =   1496
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   215
         Left            =   2220
         TabIndex        =   1495
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   216
         Left            =   2460
         TabIndex        =   1494
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   217
         Left            =   2700
         TabIndex        =   1493
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   218
         Left            =   2940
         TabIndex        =   1492
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   219
         Left            =   3180
         TabIndex        =   1491
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   220
         Left            =   3420
         TabIndex        =   1490
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   221
         Left            =   3660
         TabIndex        =   1489
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   222
         Left            =   3900
         TabIndex        =   1488
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   223
         Left            =   4140
         TabIndex        =   1487
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   224
         Left            =   4380
         TabIndex        =   1486
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   225
         Left            =   4620
         TabIndex        =   1485
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   226
         Left            =   4860
         TabIndex        =   1484
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   227
         Left            =   5100
         TabIndex        =   1483
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   228
         Left            =   5340
         TabIndex        =   1482
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   229
         Left            =   5580
         TabIndex        =   1481
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   230
         Left            =   5820
         TabIndex        =   1480
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   231
         Left            =   6060
         TabIndex        =   1479
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   232
         Left            =   6300
         TabIndex        =   1478
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   233
         Left            =   6540
         TabIndex        =   1477
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   234
         Left            =   6780
         TabIndex        =   1476
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   235
         Left            =   7020
         TabIndex        =   1475
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   236
         Left            =   7260
         TabIndex        =   1474
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   237
         Left            =   7500
         TabIndex        =   1473
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   238
         Left            =   7740
         TabIndex        =   1472
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   239
         Left            =   7980
         TabIndex        =   1471
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   240
         Left            =   8220
         TabIndex        =   1470
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   241
         Left            =   8460
         TabIndex        =   1469
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   242
         Left            =   8700
         TabIndex        =   1468
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   243
         Left            =   8940
         TabIndex        =   1467
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   244
         Left            =   9180
         TabIndex        =   1466
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   245
         Left            =   9420
         TabIndex        =   1465
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   246
         Left            =   9660
         TabIndex        =   1464
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   247
         Left            =   60
         TabIndex        =   1463
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   248
         Left            =   300
         TabIndex        =   1462
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   249
         Left            =   540
         TabIndex        =   1461
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   250
         Left            =   780
         TabIndex        =   1460
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   251
         Left            =   1020
         TabIndex        =   1459
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   252
         Left            =   1260
         TabIndex        =   1458
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   253
         Left            =   1500
         TabIndex        =   1457
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   254
         Left            =   1740
         TabIndex        =   1456
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   255
         Left            =   1980
         TabIndex        =   1455
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   256
         Left            =   2220
         TabIndex        =   1454
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   257
         Left            =   2460
         TabIndex        =   1453
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   258
         Left            =   2700
         TabIndex        =   1452
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   259
         Left            =   2940
         TabIndex        =   1451
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   260
         Left            =   3180
         TabIndex        =   1450
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   261
         Left            =   3420
         TabIndex        =   1449
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   262
         Left            =   3660
         TabIndex        =   1448
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   263
         Left            =   3900
         TabIndex        =   1447
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   264
         Left            =   4140
         TabIndex        =   1446
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   265
         Left            =   4380
         TabIndex        =   1445
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   266
         Left            =   4620
         TabIndex        =   1444
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   267
         Left            =   4860
         TabIndex        =   1443
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   268
         Left            =   5100
         TabIndex        =   1442
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   269
         Left            =   5340
         TabIndex        =   1441
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   270
         Left            =   5580
         TabIndex        =   1440
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   271
         Left            =   5820
         TabIndex        =   1439
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   272
         Left            =   6060
         TabIndex        =   1438
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   273
         Left            =   6300
         TabIndex        =   1437
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   274
         Left            =   6540
         TabIndex        =   1436
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   275
         Left            =   6780
         TabIndex        =   1435
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   276
         Left            =   7020
         TabIndex        =   1434
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   277
         Left            =   7260
         TabIndex        =   1433
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   278
         Left            =   7500
         TabIndex        =   1432
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   279
         Left            =   7740
         TabIndex        =   1431
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   280
         Left            =   7980
         TabIndex        =   1430
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   281
         Left            =   8220
         TabIndex        =   1429
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   282
         Left            =   8460
         TabIndex        =   1428
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   283
         Left            =   8700
         TabIndex        =   1427
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   284
         Left            =   8940
         TabIndex        =   1426
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   285
         Left            =   9180
         TabIndex        =   1425
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   286
         Left            =   9420
         TabIndex        =   1424
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   287
         Left            =   9660
         TabIndex        =   1423
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   288
         Left            =   60
         TabIndex        =   1422
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   289
         Left            =   300
         TabIndex        =   1421
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   290
         Left            =   540
         TabIndex        =   1420
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   291
         Left            =   780
         TabIndex        =   1419
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   292
         Left            =   1020
         TabIndex        =   1418
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   293
         Left            =   1260
         TabIndex        =   1417
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   294
         Left            =   1500
         TabIndex        =   1416
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   295
         Left            =   1740
         TabIndex        =   1415
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   296
         Left            =   1980
         TabIndex        =   1414
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   297
         Left            =   2220
         TabIndex        =   1413
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   298
         Left            =   2460
         TabIndex        =   1412
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   299
         Left            =   2700
         TabIndex        =   1411
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   300
         Left            =   2940
         TabIndex        =   1410
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   301
         Left            =   3180
         TabIndex        =   1409
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   302
         Left            =   3420
         TabIndex        =   1408
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   303
         Left            =   3660
         TabIndex        =   1407
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   304
         Left            =   3900
         TabIndex        =   1406
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   305
         Left            =   4140
         TabIndex        =   1405
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   306
         Left            =   4380
         TabIndex        =   1404
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   307
         Left            =   4620
         TabIndex        =   1403
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   308
         Left            =   4860
         TabIndex        =   1402
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   309
         Left            =   5100
         TabIndex        =   1401
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   310
         Left            =   5340
         TabIndex        =   1400
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   311
         Left            =   5580
         TabIndex        =   1399
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   312
         Left            =   5820
         TabIndex        =   1398
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   313
         Left            =   6060
         TabIndex        =   1397
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   314
         Left            =   6300
         TabIndex        =   1396
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   315
         Left            =   6540
         TabIndex        =   1395
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   316
         Left            =   6780
         TabIndex        =   1394
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   317
         Left            =   7020
         TabIndex        =   1393
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   318
         Left            =   7260
         TabIndex        =   1392
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   319
         Left            =   7500
         TabIndex        =   1391
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   320
         Left            =   7740
         TabIndex        =   1390
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   321
         Left            =   7980
         TabIndex        =   1389
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   322
         Left            =   8220
         TabIndex        =   1388
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   323
         Left            =   8460
         TabIndex        =   1387
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   324
         Left            =   8700
         TabIndex        =   1386
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   325
         Left            =   8940
         TabIndex        =   1385
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   326
         Left            =   9180
         TabIndex        =   1384
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   327
         Left            =   9420
         TabIndex        =   1383
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   328
         Left            =   9660
         TabIndex        =   1382
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   329
         Left            =   60
         TabIndex        =   1381
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   330
         Left            =   300
         TabIndex        =   1380
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   331
         Left            =   540
         TabIndex        =   1379
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   332
         Left            =   780
         TabIndex        =   1378
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   333
         Left            =   1020
         TabIndex        =   1377
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   334
         Left            =   1260
         TabIndex        =   1376
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   335
         Left            =   1500
         TabIndex        =   1375
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   336
         Left            =   1740
         TabIndex        =   1374
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   337
         Left            =   1980
         TabIndex        =   1373
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   338
         Left            =   2220
         TabIndex        =   1372
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   339
         Left            =   2460
         TabIndex        =   1371
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   340
         Left            =   2700
         TabIndex        =   1370
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   341
         Left            =   2940
         TabIndex        =   1369
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   342
         Left            =   3180
         TabIndex        =   1368
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   343
         Left            =   3420
         TabIndex        =   1367
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   344
         Left            =   3660
         TabIndex        =   1366
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   345
         Left            =   3900
         TabIndex        =   1365
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   346
         Left            =   4140
         TabIndex        =   1364
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   347
         Left            =   4380
         TabIndex        =   1363
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   348
         Left            =   4620
         TabIndex        =   1362
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   349
         Left            =   4860
         TabIndex        =   1361
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   350
         Left            =   5100
         TabIndex        =   1360
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   351
         Left            =   5340
         TabIndex        =   1359
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   352
         Left            =   5580
         TabIndex        =   1358
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   353
         Left            =   5820
         TabIndex        =   1357
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   354
         Left            =   6060
         TabIndex        =   1356
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   355
         Left            =   6300
         TabIndex        =   1355
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   356
         Left            =   6540
         TabIndex        =   1354
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   357
         Left            =   6780
         TabIndex        =   1353
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   358
         Left            =   7020
         TabIndex        =   1352
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   359
         Left            =   7260
         TabIndex        =   1351
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   360
         Left            =   7500
         TabIndex        =   1350
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   361
         Left            =   7740
         TabIndex        =   1349
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   362
         Left            =   7980
         TabIndex        =   1348
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   363
         Left            =   8220
         TabIndex        =   1347
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   364
         Left            =   8460
         TabIndex        =   1346
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   365
         Left            =   8700
         TabIndex        =   1345
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   366
         Left            =   8940
         TabIndex        =   1344
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   367
         Left            =   9180
         TabIndex        =   1343
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   368
         Left            =   9420
         TabIndex        =   1342
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   369
         Left            =   9660
         TabIndex        =   1341
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   370
         Left            =   60
         TabIndex        =   1340
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   371
         Left            =   300
         TabIndex        =   1339
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   372
         Left            =   540
         TabIndex        =   1338
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   373
         Left            =   780
         TabIndex        =   1337
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   374
         Left            =   1020
         TabIndex        =   1336
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   375
         Left            =   1260
         TabIndex        =   1335
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   376
         Left            =   1500
         TabIndex        =   1334
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   377
         Left            =   1740
         TabIndex        =   1333
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   378
         Left            =   1980
         TabIndex        =   1332
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   379
         Left            =   2220
         TabIndex        =   1331
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   380
         Left            =   2460
         TabIndex        =   1330
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   381
         Left            =   2700
         TabIndex        =   1329
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   382
         Left            =   2940
         TabIndex        =   1328
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   383
         Left            =   3180
         TabIndex        =   1327
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   384
         Left            =   3420
         TabIndex        =   1326
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   385
         Left            =   3660
         TabIndex        =   1325
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   386
         Left            =   3900
         TabIndex        =   1324
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   387
         Left            =   4140
         TabIndex        =   1323
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   388
         Left            =   4380
         TabIndex        =   1322
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   389
         Left            =   4620
         TabIndex        =   1321
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   390
         Left            =   4860
         TabIndex        =   1320
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   391
         Left            =   5100
         TabIndex        =   1319
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   392
         Left            =   5340
         TabIndex        =   1318
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   393
         Left            =   5580
         TabIndex        =   1317
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   394
         Left            =   5820
         TabIndex        =   1316
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   395
         Left            =   6060
         TabIndex        =   1315
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   396
         Left            =   6300
         TabIndex        =   1314
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   397
         Left            =   6540
         TabIndex        =   1313
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   398
         Left            =   6780
         TabIndex        =   1312
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   399
         Left            =   7020
         TabIndex        =   1311
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   400
         Left            =   7260
         TabIndex        =   1310
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   401
         Left            =   7500
         TabIndex        =   1309
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   402
         Left            =   7740
         TabIndex        =   1308
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   403
         Left            =   7980
         TabIndex        =   1307
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   404
         Left            =   8220
         TabIndex        =   1306
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   405
         Left            =   8460
         TabIndex        =   1305
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   406
         Left            =   8700
         TabIndex        =   1304
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   407
         Left            =   8940
         TabIndex        =   1303
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   408
         Left            =   9180
         TabIndex        =   1302
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   409
         Left            =   9420
         TabIndex        =   1301
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   410
         Left            =   9660
         TabIndex        =   1300
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   411
         Left            =   60
         TabIndex        =   1299
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   412
         Left            =   300
         TabIndex        =   1298
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   413
         Left            =   540
         TabIndex        =   1297
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   414
         Left            =   780
         TabIndex        =   1296
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   415
         Left            =   1020
         TabIndex        =   1295
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   416
         Left            =   1260
         TabIndex        =   1294
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   417
         Left            =   1500
         TabIndex        =   1293
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   418
         Left            =   1740
         TabIndex        =   1292
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   419
         Left            =   1980
         TabIndex        =   1291
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   420
         Left            =   2220
         TabIndex        =   1290
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   421
         Left            =   2460
         TabIndex        =   1289
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   422
         Left            =   2700
         TabIndex        =   1288
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   423
         Left            =   2940
         TabIndex        =   1287
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   424
         Left            =   3180
         TabIndex        =   1286
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   425
         Left            =   3420
         TabIndex        =   1285
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   426
         Left            =   3660
         TabIndex        =   1284
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   427
         Left            =   3900
         TabIndex        =   1283
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   428
         Left            =   4140
         TabIndex        =   1282
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   429
         Left            =   4380
         TabIndex        =   1281
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   430
         Left            =   4620
         TabIndex        =   1280
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   431
         Left            =   4860
         TabIndex        =   1279
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   432
         Left            =   5100
         TabIndex        =   1278
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   433
         Left            =   5340
         TabIndex        =   1277
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   434
         Left            =   5580
         TabIndex        =   1276
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   435
         Left            =   5820
         TabIndex        =   1275
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   436
         Left            =   6060
         TabIndex        =   1274
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   437
         Left            =   6300
         TabIndex        =   1273
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   438
         Left            =   6540
         TabIndex        =   1272
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   439
         Left            =   6780
         TabIndex        =   1271
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   440
         Left            =   7020
         TabIndex        =   1270
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   441
         Left            =   7260
         TabIndex        =   1269
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   442
         Left            =   7500
         TabIndex        =   1268
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   443
         Left            =   7740
         TabIndex        =   1267
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   444
         Left            =   7980
         TabIndex        =   1266
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   445
         Left            =   8220
         TabIndex        =   1265
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   446
         Left            =   8460
         TabIndex        =   1264
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   447
         Left            =   8700
         TabIndex        =   1263
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   448
         Left            =   8940
         TabIndex        =   1262
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   449
         Left            =   9180
         TabIndex        =   1261
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   450
         Left            =   9420
         TabIndex        =   1260
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   451
         Left            =   9660
         TabIndex        =   1259
         Top             =   2460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   452
         Left            =   60
         TabIndex        =   1258
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   453
         Left            =   300
         TabIndex        =   1257
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   454
         Left            =   540
         TabIndex        =   1256
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   455
         Left            =   780
         TabIndex        =   1255
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   456
         Left            =   1020
         TabIndex        =   1254
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   457
         Left            =   1260
         TabIndex        =   1253
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   458
         Left            =   1500
         TabIndex        =   1252
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   459
         Left            =   1740
         TabIndex        =   1251
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   460
         Left            =   1980
         TabIndex        =   1250
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   461
         Left            =   2220
         TabIndex        =   1249
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   462
         Left            =   2460
         TabIndex        =   1248
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   463
         Left            =   2700
         TabIndex        =   1247
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   464
         Left            =   2940
         TabIndex        =   1246
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   465
         Left            =   3180
         TabIndex        =   1245
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   466
         Left            =   3420
         TabIndex        =   1244
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   467
         Left            =   3660
         TabIndex        =   1243
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   468
         Left            =   3900
         TabIndex        =   1242
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   469
         Left            =   4140
         TabIndex        =   1241
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   470
         Left            =   4380
         TabIndex        =   1240
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   471
         Left            =   4620
         TabIndex        =   1239
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   472
         Left            =   4860
         TabIndex        =   1238
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   473
         Left            =   5100
         TabIndex        =   1237
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   474
         Left            =   5340
         TabIndex        =   1236
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   475
         Left            =   5580
         TabIndex        =   1235
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   476
         Left            =   5820
         TabIndex        =   1234
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   477
         Left            =   6060
         TabIndex        =   1233
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   478
         Left            =   6300
         TabIndex        =   1232
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   479
         Left            =   6540
         TabIndex        =   1231
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   480
         Left            =   6780
         TabIndex        =   1230
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   481
         Left            =   7020
         TabIndex        =   1229
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   482
         Left            =   7260
         TabIndex        =   1228
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   483
         Left            =   7500
         TabIndex        =   1227
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   484
         Left            =   7740
         TabIndex        =   1226
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   485
         Left            =   7980
         TabIndex        =   1225
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   486
         Left            =   8220
         TabIndex        =   1224
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   487
         Left            =   8460
         TabIndex        =   1223
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   488
         Left            =   8700
         TabIndex        =   1222
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   489
         Left            =   8940
         TabIndex        =   1221
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   490
         Left            =   9180
         TabIndex        =   1220
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   491
         Left            =   9420
         TabIndex        =   1219
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   492
         Left            =   9660
         TabIndex        =   1218
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   493
         Left            =   60
         TabIndex        =   1217
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   494
         Left            =   300
         TabIndex        =   1216
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   495
         Left            =   540
         TabIndex        =   1215
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   496
         Left            =   780
         TabIndex        =   1214
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   497
         Left            =   1020
         TabIndex        =   1213
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   498
         Left            =   1260
         TabIndex        =   1212
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   499
         Left            =   1500
         TabIndex        =   1211
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   500
         Left            =   1740
         TabIndex        =   1210
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   501
         Left            =   1980
         TabIndex        =   1209
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   502
         Left            =   2220
         TabIndex        =   1208
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   503
         Left            =   2460
         TabIndex        =   1207
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   504
         Left            =   2700
         TabIndex        =   1206
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   505
         Left            =   2940
         TabIndex        =   1205
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   506
         Left            =   3180
         TabIndex        =   1204
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   507
         Left            =   3420
         TabIndex        =   1203
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   508
         Left            =   3660
         TabIndex        =   1202
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   509
         Left            =   3900
         TabIndex        =   1201
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   510
         Left            =   4140
         TabIndex        =   1200
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   511
         Left            =   4380
         TabIndex        =   1199
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   512
         Left            =   4620
         TabIndex        =   1198
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   513
         Left            =   4860
         TabIndex        =   1197
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   514
         Left            =   5100
         TabIndex        =   1196
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   515
         Left            =   5340
         TabIndex        =   1195
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   516
         Left            =   5580
         TabIndex        =   1194
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   517
         Left            =   5820
         TabIndex        =   1193
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   518
         Left            =   6060
         TabIndex        =   1192
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   519
         Left            =   6300
         TabIndex        =   1191
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   520
         Left            =   6540
         TabIndex        =   1190
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   521
         Left            =   6780
         TabIndex        =   1189
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   522
         Left            =   7020
         TabIndex        =   1188
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   523
         Left            =   7260
         TabIndex        =   1187
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   524
         Left            =   7500
         TabIndex        =   1186
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   525
         Left            =   7740
         TabIndex        =   1185
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   526
         Left            =   7980
         TabIndex        =   1184
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   527
         Left            =   8220
         TabIndex        =   1183
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   528
         Left            =   8460
         TabIndex        =   1182
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   529
         Left            =   8700
         TabIndex        =   1181
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   530
         Left            =   8940
         TabIndex        =   1180
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   531
         Left            =   9180
         TabIndex        =   1179
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   532
         Left            =   9420
         TabIndex        =   1178
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   533
         Left            =   9660
         TabIndex        =   1177
         Top             =   2940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   534
         Left            =   60
         TabIndex        =   1176
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   535
         Left            =   300
         TabIndex        =   1175
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   536
         Left            =   540
         TabIndex        =   1174
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   537
         Left            =   780
         TabIndex        =   1173
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   538
         Left            =   1020
         TabIndex        =   1172
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   539
         Left            =   1260
         TabIndex        =   1171
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   540
         Left            =   1500
         TabIndex        =   1170
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   541
         Left            =   1740
         TabIndex        =   1169
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   542
         Left            =   1980
         TabIndex        =   1168
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   543
         Left            =   2220
         TabIndex        =   1167
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   544
         Left            =   2460
         TabIndex        =   1166
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   545
         Left            =   2700
         TabIndex        =   1165
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   546
         Left            =   2940
         TabIndex        =   1164
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   547
         Left            =   3180
         TabIndex        =   1163
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   548
         Left            =   3420
         TabIndex        =   1162
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   549
         Left            =   3660
         TabIndex        =   1161
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   550
         Left            =   3900
         TabIndex        =   1160
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   551
         Left            =   4140
         TabIndex        =   1159
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   552
         Left            =   4380
         TabIndex        =   1158
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   553
         Left            =   4620
         TabIndex        =   1157
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   554
         Left            =   4860
         TabIndex        =   1156
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   555
         Left            =   5100
         TabIndex        =   1155
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   556
         Left            =   5340
         TabIndex        =   1154
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   557
         Left            =   5580
         TabIndex        =   1153
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   558
         Left            =   5820
         TabIndex        =   1152
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   559
         Left            =   6060
         TabIndex        =   1151
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   560
         Left            =   6300
         TabIndex        =   1150
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   561
         Left            =   6540
         TabIndex        =   1149
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   562
         Left            =   6780
         TabIndex        =   1148
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   563
         Left            =   7020
         TabIndex        =   1147
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   564
         Left            =   7260
         TabIndex        =   1146
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   565
         Left            =   7500
         TabIndex        =   1145
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   566
         Left            =   7740
         TabIndex        =   1144
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   567
         Left            =   7980
         TabIndex        =   1143
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   568
         Left            =   8220
         TabIndex        =   1142
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   569
         Left            =   8460
         TabIndex        =   1141
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   570
         Left            =   8700
         TabIndex        =   1140
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   571
         Left            =   8940
         TabIndex        =   1139
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   572
         Left            =   9180
         TabIndex        =   1138
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   573
         Left            =   9420
         TabIndex        =   1137
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   574
         Left            =   9660
         TabIndex        =   1136
         Top             =   3180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   575
         Left            =   60
         TabIndex        =   1135
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   576
         Left            =   300
         TabIndex        =   1134
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   577
         Left            =   540
         TabIndex        =   1133
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   578
         Left            =   780
         TabIndex        =   1132
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   579
         Left            =   1020
         TabIndex        =   1131
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   580
         Left            =   1260
         TabIndex        =   1130
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   581
         Left            =   1500
         TabIndex        =   1129
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   582
         Left            =   1740
         TabIndex        =   1128
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   583
         Left            =   1980
         TabIndex        =   1127
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   584
         Left            =   2220
         TabIndex        =   1126
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   585
         Left            =   2460
         TabIndex        =   1125
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   586
         Left            =   2700
         TabIndex        =   1124
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   587
         Left            =   2940
         TabIndex        =   1123
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   588
         Left            =   3180
         TabIndex        =   1122
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   589
         Left            =   3420
         TabIndex        =   1121
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   590
         Left            =   3660
         TabIndex        =   1120
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   591
         Left            =   3900
         TabIndex        =   1119
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   592
         Left            =   4140
         TabIndex        =   1118
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   593
         Left            =   4380
         TabIndex        =   1117
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   594
         Left            =   4620
         TabIndex        =   1116
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   595
         Left            =   4860
         TabIndex        =   1115
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   596
         Left            =   5100
         TabIndex        =   1114
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   597
         Left            =   5340
         TabIndex        =   1113
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   598
         Left            =   5580
         TabIndex        =   1112
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   599
         Left            =   5820
         TabIndex        =   1111
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   600
         Left            =   6060
         TabIndex        =   1110
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   601
         Left            =   6300
         TabIndex        =   1109
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   602
         Left            =   6540
         TabIndex        =   1108
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   603
         Left            =   6780
         TabIndex        =   1107
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   604
         Left            =   7020
         TabIndex        =   1106
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   605
         Left            =   7260
         TabIndex        =   1105
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   606
         Left            =   7500
         TabIndex        =   1104
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   607
         Left            =   7740
         TabIndex        =   1103
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   608
         Left            =   7980
         TabIndex        =   1102
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   609
         Left            =   8220
         TabIndex        =   1101
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   610
         Left            =   8460
         TabIndex        =   1100
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   611
         Left            =   8700
         TabIndex        =   1099
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   612
         Left            =   8940
         TabIndex        =   1098
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   613
         Left            =   9180
         TabIndex        =   1097
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   614
         Left            =   9420
         TabIndex        =   1096
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   615
         Left            =   9660
         TabIndex        =   1095
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   616
         Left            =   60
         TabIndex        =   1094
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   617
         Left            =   300
         TabIndex        =   1093
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   618
         Left            =   540
         TabIndex        =   1092
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   619
         Left            =   780
         TabIndex        =   1091
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   620
         Left            =   1020
         TabIndex        =   1090
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   621
         Left            =   1260
         TabIndex        =   1089
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   622
         Left            =   1500
         TabIndex        =   1088
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   623
         Left            =   1740
         TabIndex        =   1087
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   624
         Left            =   1980
         TabIndex        =   1086
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   625
         Left            =   2220
         TabIndex        =   1085
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   626
         Left            =   2460
         TabIndex        =   1084
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   627
         Left            =   2700
         TabIndex        =   1083
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   628
         Left            =   2940
         TabIndex        =   1082
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   629
         Left            =   3180
         TabIndex        =   1081
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   630
         Left            =   3420
         TabIndex        =   1080
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   631
         Left            =   3660
         TabIndex        =   1079
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   632
         Left            =   3900
         TabIndex        =   1078
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   633
         Left            =   4140
         TabIndex        =   1077
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   634
         Left            =   4380
         TabIndex        =   1076
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   635
         Left            =   4620
         TabIndex        =   1075
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   636
         Left            =   4860
         TabIndex        =   1074
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   637
         Left            =   5100
         TabIndex        =   1073
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   638
         Left            =   5340
         TabIndex        =   1072
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   639
         Left            =   5580
         TabIndex        =   1071
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   640
         Left            =   5820
         TabIndex        =   1070
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   641
         Left            =   6060
         TabIndex        =   1069
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   642
         Left            =   6300
         TabIndex        =   1068
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   643
         Left            =   6540
         TabIndex        =   1067
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   644
         Left            =   6780
         TabIndex        =   1066
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   645
         Left            =   7020
         TabIndex        =   1065
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   646
         Left            =   7260
         TabIndex        =   1064
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   647
         Left            =   7500
         TabIndex        =   1063
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   648
         Left            =   7740
         TabIndex        =   1062
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   649
         Left            =   7980
         TabIndex        =   1061
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   650
         Left            =   8220
         TabIndex        =   1060
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   651
         Left            =   8460
         TabIndex        =   1059
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   652
         Left            =   8700
         TabIndex        =   1058
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   653
         Left            =   8940
         TabIndex        =   1057
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   654
         Left            =   9180
         TabIndex        =   1056
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   655
         Left            =   9420
         TabIndex        =   1055
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   656
         Left            =   9660
         TabIndex        =   1054
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   657
         Left            =   60
         TabIndex        =   1053
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   658
         Left            =   300
         TabIndex        =   1052
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   659
         Left            =   540
         TabIndex        =   1051
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   660
         Left            =   780
         TabIndex        =   1050
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   661
         Left            =   1020
         TabIndex        =   1049
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   662
         Left            =   1260
         TabIndex        =   1048
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   663
         Left            =   1500
         TabIndex        =   1047
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   664
         Left            =   1740
         TabIndex        =   1046
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   665
         Left            =   1980
         TabIndex        =   1045
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   666
         Left            =   2220
         TabIndex        =   1044
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   667
         Left            =   2460
         TabIndex        =   1043
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   668
         Left            =   2700
         TabIndex        =   1042
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   669
         Left            =   2940
         TabIndex        =   1041
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   670
         Left            =   3180
         TabIndex        =   1040
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   671
         Left            =   3420
         TabIndex        =   1039
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   672
         Left            =   3660
         TabIndex        =   1038
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   673
         Left            =   3900
         TabIndex        =   1037
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   674
         Left            =   4140
         TabIndex        =   1036
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   675
         Left            =   4380
         TabIndex        =   1035
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   676
         Left            =   4620
         TabIndex        =   1034
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   677
         Left            =   4860
         TabIndex        =   1033
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   678
         Left            =   5100
         TabIndex        =   1032
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   679
         Left            =   5340
         TabIndex        =   1031
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   680
         Left            =   5580
         TabIndex        =   1030
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   681
         Left            =   5820
         TabIndex        =   1029
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   682
         Left            =   6060
         TabIndex        =   1028
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   683
         Left            =   6300
         TabIndex        =   1027
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   684
         Left            =   6540
         TabIndex        =   1026
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   685
         Left            =   6780
         TabIndex        =   1025
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   686
         Left            =   7020
         TabIndex        =   1024
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   687
         Left            =   7260
         TabIndex        =   1023
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   688
         Left            =   7500
         TabIndex        =   1022
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   689
         Left            =   7740
         TabIndex        =   1021
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   690
         Left            =   7980
         TabIndex        =   1020
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   691
         Left            =   8220
         TabIndex        =   1019
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   692
         Left            =   8460
         TabIndex        =   1018
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   693
         Left            =   8700
         TabIndex        =   1017
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   694
         Left            =   8940
         TabIndex        =   1016
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   695
         Left            =   9180
         TabIndex        =   1015
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   696
         Left            =   9420
         TabIndex        =   1014
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   697
         Left            =   9660
         TabIndex        =   1013
         Top             =   3900
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   698
         Left            =   60
         TabIndex        =   1012
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   699
         Left            =   300
         TabIndex        =   1011
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   700
         Left            =   540
         TabIndex        =   1010
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   701
         Left            =   780
         TabIndex        =   1009
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   702
         Left            =   1020
         TabIndex        =   1008
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   703
         Left            =   1260
         TabIndex        =   1007
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   704
         Left            =   1500
         TabIndex        =   1006
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   705
         Left            =   1740
         TabIndex        =   1005
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   706
         Left            =   1980
         TabIndex        =   1004
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   707
         Left            =   2220
         TabIndex        =   1003
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   708
         Left            =   2460
         TabIndex        =   1002
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   709
         Left            =   2700
         TabIndex        =   1001
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   710
         Left            =   2940
         TabIndex        =   1000
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   711
         Left            =   3180
         TabIndex        =   999
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   712
         Left            =   3420
         TabIndex        =   998
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   713
         Left            =   3660
         TabIndex        =   997
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   714
         Left            =   3900
         TabIndex        =   996
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   715
         Left            =   4140
         TabIndex        =   995
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   716
         Left            =   4380
         TabIndex        =   994
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   717
         Left            =   4620
         TabIndex        =   993
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   718
         Left            =   4860
         TabIndex        =   992
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   719
         Left            =   5100
         TabIndex        =   991
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   720
         Left            =   5340
         TabIndex        =   990
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   721
         Left            =   5580
         TabIndex        =   989
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   722
         Left            =   5820
         TabIndex        =   988
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   723
         Left            =   6060
         TabIndex        =   987
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   724
         Left            =   6300
         TabIndex        =   986
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   725
         Left            =   6540
         TabIndex        =   985
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   726
         Left            =   6780
         TabIndex        =   984
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   727
         Left            =   7020
         TabIndex        =   983
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   728
         Left            =   7260
         TabIndex        =   982
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   729
         Left            =   7500
         TabIndex        =   981
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   730
         Left            =   7740
         TabIndex        =   980
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   731
         Left            =   7980
         TabIndex        =   979
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   732
         Left            =   8220
         TabIndex        =   978
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   733
         Left            =   8460
         TabIndex        =   977
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   734
         Left            =   8700
         TabIndex        =   976
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   735
         Left            =   8940
         TabIndex        =   975
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   736
         Left            =   9180
         TabIndex        =   974
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   737
         Left            =   9420
         TabIndex        =   973
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   738
         Left            =   9660
         TabIndex        =   972
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   739
         Left            =   60
         TabIndex        =   971
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   740
         Left            =   300
         TabIndex        =   970
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   741
         Left            =   540
         TabIndex        =   969
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   742
         Left            =   780
         TabIndex        =   968
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   743
         Left            =   1020
         TabIndex        =   967
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   744
         Left            =   1260
         TabIndex        =   966
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   745
         Left            =   1500
         TabIndex        =   965
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   746
         Left            =   1740
         TabIndex        =   964
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   747
         Left            =   1980
         TabIndex        =   963
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   748
         Left            =   2220
         TabIndex        =   962
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   749
         Left            =   2460
         TabIndex        =   961
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   750
         Left            =   2700
         TabIndex        =   960
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   751
         Left            =   2940
         TabIndex        =   959
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   752
         Left            =   3180
         TabIndex        =   958
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   753
         Left            =   3420
         TabIndex        =   957
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   754
         Left            =   3660
         TabIndex        =   956
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   755
         Left            =   3900
         TabIndex        =   955
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   756
         Left            =   4140
         TabIndex        =   954
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   757
         Left            =   4380
         TabIndex        =   953
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   758
         Left            =   4620
         TabIndex        =   952
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   759
         Left            =   4860
         TabIndex        =   951
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   760
         Left            =   5100
         TabIndex        =   950
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   761
         Left            =   5340
         TabIndex        =   949
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   762
         Left            =   5580
         TabIndex        =   948
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   763
         Left            =   5820
         TabIndex        =   947
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   764
         Left            =   6060
         TabIndex        =   946
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   765
         Left            =   6300
         TabIndex        =   945
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   766
         Left            =   6540
         TabIndex        =   944
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   767
         Left            =   6780
         TabIndex        =   943
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   768
         Left            =   7020
         TabIndex        =   942
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   769
         Left            =   7260
         TabIndex        =   941
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   770
         Left            =   7500
         TabIndex        =   940
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   771
         Left            =   7740
         TabIndex        =   939
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   772
         Left            =   7980
         TabIndex        =   938
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   773
         Left            =   8220
         TabIndex        =   937
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   774
         Left            =   8460
         TabIndex        =   936
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   775
         Left            =   8700
         TabIndex        =   935
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   776
         Left            =   8940
         TabIndex        =   934
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   777
         Left            =   9180
         TabIndex        =   933
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   778
         Left            =   9420
         TabIndex        =   932
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   779
         Left            =   9660
         TabIndex        =   931
         Top             =   4380
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   780
         Left            =   60
         TabIndex        =   930
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   781
         Left            =   300
         TabIndex        =   929
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   782
         Left            =   540
         TabIndex        =   928
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   783
         Left            =   780
         TabIndex        =   927
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   784
         Left            =   1020
         TabIndex        =   926
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   785
         Left            =   1260
         TabIndex        =   925
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   786
         Left            =   1500
         TabIndex        =   924
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   787
         Left            =   1740
         TabIndex        =   923
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   788
         Left            =   1980
         TabIndex        =   922
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   789
         Left            =   2220
         TabIndex        =   921
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   790
         Left            =   2460
         TabIndex        =   920
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   791
         Left            =   2700
         TabIndex        =   919
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   792
         Left            =   2940
         TabIndex        =   918
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   793
         Left            =   3180
         TabIndex        =   917
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   794
         Left            =   3420
         TabIndex        =   916
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   795
         Left            =   3660
         TabIndex        =   915
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   796
         Left            =   3900
         TabIndex        =   914
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   797
         Left            =   4140
         TabIndex        =   913
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   798
         Left            =   4380
         TabIndex        =   912
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   799
         Left            =   4620
         TabIndex        =   911
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   800
         Left            =   4860
         TabIndex        =   910
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   801
         Left            =   5100
         TabIndex        =   909
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   802
         Left            =   5340
         TabIndex        =   908
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   803
         Left            =   5580
         TabIndex        =   907
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   804
         Left            =   5820
         TabIndex        =   906
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   805
         Left            =   6060
         TabIndex        =   905
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   806
         Left            =   6300
         TabIndex        =   904
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   807
         Left            =   6540
         TabIndex        =   903
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   808
         Left            =   6780
         TabIndex        =   902
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   809
         Left            =   7020
         TabIndex        =   901
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   810
         Left            =   7260
         TabIndex        =   900
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   811
         Left            =   7500
         TabIndex        =   899
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   812
         Left            =   7740
         TabIndex        =   898
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   813
         Left            =   7980
         TabIndex        =   897
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   814
         Left            =   8220
         TabIndex        =   896
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   815
         Left            =   8460
         TabIndex        =   895
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   816
         Left            =   8700
         TabIndex        =   894
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   817
         Left            =   8940
         TabIndex        =   893
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   818
         Left            =   9180
         TabIndex        =   892
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   819
         Left            =   9420
         TabIndex        =   891
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   820
         Left            =   9660
         TabIndex        =   890
         Top             =   4620
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   821
         Left            =   60
         TabIndex        =   889
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   822
         Left            =   300
         TabIndex        =   888
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   823
         Left            =   540
         TabIndex        =   887
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   824
         Left            =   780
         TabIndex        =   886
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   825
         Left            =   1020
         TabIndex        =   885
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   826
         Left            =   1260
         TabIndex        =   884
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   827
         Left            =   1500
         TabIndex        =   883
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   828
         Left            =   1740
         TabIndex        =   882
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   829
         Left            =   1980
         TabIndex        =   881
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   830
         Left            =   2220
         TabIndex        =   880
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   831
         Left            =   2460
         TabIndex        =   879
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   832
         Left            =   2700
         TabIndex        =   878
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   833
         Left            =   2940
         TabIndex        =   877
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   834
         Left            =   3180
         TabIndex        =   876
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   835
         Left            =   3420
         TabIndex        =   875
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   836
         Left            =   3660
         TabIndex        =   874
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   837
         Left            =   3900
         TabIndex        =   873
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   838
         Left            =   4140
         TabIndex        =   872
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   839
         Left            =   4380
         TabIndex        =   871
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   840
         Left            =   4620
         TabIndex        =   870
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   841
         Left            =   4860
         TabIndex        =   869
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   842
         Left            =   5100
         TabIndex        =   868
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   843
         Left            =   5340
         TabIndex        =   867
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   844
         Left            =   5580
         TabIndex        =   866
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   845
         Left            =   5820
         TabIndex        =   865
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   846
         Left            =   6060
         TabIndex        =   864
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   847
         Left            =   6300
         TabIndex        =   863
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   848
         Left            =   6540
         TabIndex        =   862
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   849
         Left            =   6780
         TabIndex        =   861
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   850
         Left            =   7020
         TabIndex        =   860
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   851
         Left            =   7260
         TabIndex        =   859
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   852
         Left            =   7500
         TabIndex        =   858
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   853
         Left            =   7740
         TabIndex        =   857
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   854
         Left            =   7980
         TabIndex        =   856
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   855
         Left            =   8220
         TabIndex        =   855
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   856
         Left            =   8460
         TabIndex        =   854
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   857
         Left            =   8700
         TabIndex        =   853
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   858
         Left            =   8940
         TabIndex        =   852
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   859
         Left            =   9180
         TabIndex        =   851
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   860
         Left            =   9420
         TabIndex        =   850
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   861
         Left            =   9660
         TabIndex        =   849
         Top             =   4860
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   862
         Left            =   60
         TabIndex        =   848
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   863
         Left            =   300
         TabIndex        =   847
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   864
         Left            =   540
         TabIndex        =   846
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   865
         Left            =   780
         TabIndex        =   845
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   866
         Left            =   1020
         TabIndex        =   844
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   867
         Left            =   1260
         TabIndex        =   843
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   868
         Left            =   1500
         TabIndex        =   842
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   869
         Left            =   1740
         TabIndex        =   841
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   870
         Left            =   1980
         TabIndex        =   840
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   871
         Left            =   2220
         TabIndex        =   839
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   872
         Left            =   2460
         TabIndex        =   838
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   873
         Left            =   2700
         TabIndex        =   837
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   874
         Left            =   2940
         TabIndex        =   836
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   875
         Left            =   3180
         TabIndex        =   835
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   876
         Left            =   3420
         TabIndex        =   834
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   877
         Left            =   3660
         TabIndex        =   833
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   878
         Left            =   3900
         TabIndex        =   832
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   879
         Left            =   4140
         TabIndex        =   831
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   880
         Left            =   4380
         TabIndex        =   830
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   881
         Left            =   4620
         TabIndex        =   829
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   882
         Left            =   4860
         TabIndex        =   828
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   883
         Left            =   5100
         TabIndex        =   827
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   884
         Left            =   5340
         TabIndex        =   826
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   885
         Left            =   5580
         TabIndex        =   825
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   886
         Left            =   5820
         TabIndex        =   824
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   887
         Left            =   6060
         TabIndex        =   823
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   888
         Left            =   6300
         TabIndex        =   822
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   889
         Left            =   6540
         TabIndex        =   821
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   890
         Left            =   6780
         TabIndex        =   820
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   891
         Left            =   7020
         TabIndex        =   819
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   892
         Left            =   7260
         TabIndex        =   818
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   893
         Left            =   7500
         TabIndex        =   817
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   894
         Left            =   7740
         TabIndex        =   816
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   895
         Left            =   7980
         TabIndex        =   815
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   896
         Left            =   8220
         TabIndex        =   814
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   897
         Left            =   8460
         TabIndex        =   813
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   898
         Left            =   8700
         TabIndex        =   812
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   899
         Left            =   8940
         TabIndex        =   811
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   900
         Left            =   9180
         TabIndex        =   810
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   901
         Left            =   9420
         TabIndex        =   809
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   902
         Left            =   9660
         TabIndex        =   808
         Top             =   5100
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   903
         Left            =   60
         TabIndex        =   807
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   904
         Left            =   300
         TabIndex        =   806
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   905
         Left            =   540
         TabIndex        =   805
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   906
         Left            =   780
         TabIndex        =   804
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   907
         Left            =   1020
         TabIndex        =   803
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   908
         Left            =   1260
         TabIndex        =   802
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   909
         Left            =   1500
         TabIndex        =   801
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   910
         Left            =   1740
         TabIndex        =   800
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   911
         Left            =   1980
         TabIndex        =   799
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   912
         Left            =   2220
         TabIndex        =   798
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   913
         Left            =   2460
         TabIndex        =   797
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   914
         Left            =   2700
         TabIndex        =   796
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   915
         Left            =   2940
         TabIndex        =   795
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   916
         Left            =   3180
         TabIndex        =   794
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   917
         Left            =   3420
         TabIndex        =   793
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   918
         Left            =   3660
         TabIndex        =   792
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   919
         Left            =   3900
         TabIndex        =   791
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   920
         Left            =   4140
         TabIndex        =   790
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   921
         Left            =   4380
         TabIndex        =   789
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   922
         Left            =   4620
         TabIndex        =   788
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   923
         Left            =   4860
         TabIndex        =   787
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   924
         Left            =   5100
         TabIndex        =   786
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   925
         Left            =   5340
         TabIndex        =   785
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   926
         Left            =   5580
         TabIndex        =   784
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   927
         Left            =   5820
         TabIndex        =   783
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   928
         Left            =   6060
         TabIndex        =   782
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   929
         Left            =   6300
         TabIndex        =   781
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   930
         Left            =   6540
         TabIndex        =   780
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   931
         Left            =   6780
         TabIndex        =   779
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   932
         Left            =   7020
         TabIndex        =   778
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   933
         Left            =   7260
         TabIndex        =   777
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   934
         Left            =   7500
         TabIndex        =   776
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   935
         Left            =   7740
         TabIndex        =   775
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   936
         Left            =   7980
         TabIndex        =   774
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   937
         Left            =   8220
         TabIndex        =   773
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   938
         Left            =   8460
         TabIndex        =   772
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   939
         Left            =   8700
         TabIndex        =   771
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   940
         Left            =   8940
         TabIndex        =   770
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   941
         Left            =   9180
         TabIndex        =   769
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   942
         Left            =   9420
         TabIndex        =   768
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   943
         Left            =   9660
         TabIndex        =   767
         Top             =   5340
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   944
         Left            =   60
         TabIndex        =   766
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   945
         Left            =   300
         TabIndex        =   765
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   946
         Left            =   540
         TabIndex        =   764
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   947
         Left            =   780
         TabIndex        =   763
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   948
         Left            =   1020
         TabIndex        =   762
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   949
         Left            =   1260
         TabIndex        =   761
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   950
         Left            =   1500
         TabIndex        =   760
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   951
         Left            =   1740
         TabIndex        =   759
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   952
         Left            =   1980
         TabIndex        =   758
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   953
         Left            =   2220
         TabIndex        =   757
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   954
         Left            =   2460
         TabIndex        =   756
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   955
         Left            =   2700
         TabIndex        =   755
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   956
         Left            =   2940
         TabIndex        =   754
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   957
         Left            =   3180
         TabIndex        =   753
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   958
         Left            =   3420
         TabIndex        =   752
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   959
         Left            =   3660
         TabIndex        =   751
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   960
         Left            =   3900
         TabIndex        =   750
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   961
         Left            =   4140
         TabIndex        =   749
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   962
         Left            =   4380
         TabIndex        =   748
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   963
         Left            =   4620
         TabIndex        =   747
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   964
         Left            =   4860
         TabIndex        =   746
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   965
         Left            =   5100
         TabIndex        =   745
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   966
         Left            =   5340
         TabIndex        =   744
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   967
         Left            =   5580
         TabIndex        =   743
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   968
         Left            =   5820
         TabIndex        =   742
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   969
         Left            =   6060
         TabIndex        =   741
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   970
         Left            =   6300
         TabIndex        =   740
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   971
         Left            =   6540
         TabIndex        =   739
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   972
         Left            =   6780
         TabIndex        =   738
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   973
         Left            =   7020
         TabIndex        =   737
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   974
         Left            =   7260
         TabIndex        =   736
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   975
         Left            =   7500
         TabIndex        =   735
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   976
         Left            =   7740
         TabIndex        =   734
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   977
         Left            =   7980
         TabIndex        =   733
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   978
         Left            =   8220
         TabIndex        =   732
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   979
         Left            =   8460
         TabIndex        =   731
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   980
         Left            =   8700
         TabIndex        =   730
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   981
         Left            =   8940
         TabIndex        =   729
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   982
         Left            =   9180
         TabIndex        =   728
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   983
         Left            =   9420
         TabIndex        =   727
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   984
         Left            =   9660
         TabIndex        =   726
         Top             =   5580
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   985
         Left            =   60
         TabIndex        =   725
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   986
         Left            =   300
         TabIndex        =   724
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   987
         Left            =   540
         TabIndex        =   723
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   988
         Left            =   780
         TabIndex        =   722
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   989
         Left            =   1020
         TabIndex        =   721
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   990
         Left            =   1260
         TabIndex        =   720
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   991
         Left            =   1500
         TabIndex        =   719
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   992
         Left            =   1740
         TabIndex        =   718
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   993
         Left            =   1980
         TabIndex        =   717
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   994
         Left            =   2220
         TabIndex        =   716
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   995
         Left            =   2460
         TabIndex        =   715
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   996
         Left            =   2700
         TabIndex        =   714
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   997
         Left            =   2940
         TabIndex        =   713
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   998
         Left            =   3180
         TabIndex        =   712
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   999
         Left            =   3420
         TabIndex        =   711
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1000
         Left            =   3660
         TabIndex        =   710
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1001
         Left            =   3900
         TabIndex        =   709
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1002
         Left            =   4140
         TabIndex        =   708
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1003
         Left            =   4380
         TabIndex        =   707
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1004
         Left            =   4620
         TabIndex        =   706
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1005
         Left            =   4860
         TabIndex        =   705
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1006
         Left            =   5100
         TabIndex        =   704
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1007
         Left            =   5340
         TabIndex        =   703
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1008
         Left            =   5580
         TabIndex        =   702
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1009
         Left            =   5820
         TabIndex        =   701
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1010
         Left            =   6060
         TabIndex        =   700
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1011
         Left            =   6300
         TabIndex        =   699
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1012
         Left            =   6540
         TabIndex        =   698
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1013
         Left            =   6780
         TabIndex        =   697
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1014
         Left            =   7020
         TabIndex        =   696
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1015
         Left            =   7260
         TabIndex        =   695
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1016
         Left            =   7500
         TabIndex        =   694
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1017
         Left            =   7740
         TabIndex        =   693
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1018
         Left            =   7980
         TabIndex        =   692
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1019
         Left            =   8220
         TabIndex        =   691
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1020
         Left            =   8460
         TabIndex        =   690
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1021
         Left            =   8700
         TabIndex        =   689
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1022
         Left            =   8940
         TabIndex        =   688
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1023
         Left            =   9180
         TabIndex        =   687
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1024
         Left            =   9420
         TabIndex        =   686
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1025
         Left            =   9660
         TabIndex        =   685
         Top             =   5820
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1026
         Left            =   60
         TabIndex        =   684
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1027
         Left            =   300
         TabIndex        =   683
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1028
         Left            =   540
         TabIndex        =   682
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1029
         Left            =   780
         TabIndex        =   681
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1030
         Left            =   1020
         TabIndex        =   680
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1031
         Left            =   1260
         TabIndex        =   679
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1032
         Left            =   1500
         TabIndex        =   678
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1033
         Left            =   1740
         TabIndex        =   677
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1034
         Left            =   1980
         TabIndex        =   676
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1035
         Left            =   2220
         TabIndex        =   675
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1036
         Left            =   2460
         TabIndex        =   674
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1037
         Left            =   2700
         TabIndex        =   673
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1038
         Left            =   2940
         TabIndex        =   672
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1039
         Left            =   3180
         TabIndex        =   671
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1040
         Left            =   3420
         TabIndex        =   670
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1041
         Left            =   3660
         TabIndex        =   669
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1042
         Left            =   3900
         TabIndex        =   668
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1043
         Left            =   4140
         TabIndex        =   667
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1044
         Left            =   4380
         TabIndex        =   666
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1045
         Left            =   4620
         TabIndex        =   665
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1046
         Left            =   4860
         TabIndex        =   664
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1047
         Left            =   5100
         TabIndex        =   663
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1048
         Left            =   5340
         TabIndex        =   662
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1049
         Left            =   5580
         TabIndex        =   661
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1050
         Left            =   5820
         TabIndex        =   660
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1051
         Left            =   6060
         TabIndex        =   659
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1052
         Left            =   6300
         TabIndex        =   658
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1053
         Left            =   6540
         TabIndex        =   657
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1054
         Left            =   6780
         TabIndex        =   656
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1055
         Left            =   7020
         TabIndex        =   655
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1056
         Left            =   7260
         TabIndex        =   654
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1057
         Left            =   7500
         TabIndex        =   653
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1058
         Left            =   7740
         TabIndex        =   652
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1059
         Left            =   7980
         TabIndex        =   651
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1060
         Left            =   8220
         TabIndex        =   650
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1061
         Left            =   8460
         TabIndex        =   649
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1062
         Left            =   8700
         TabIndex        =   648
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1063
         Left            =   8940
         TabIndex        =   647
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1064
         Left            =   9180
         TabIndex        =   646
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1065
         Left            =   9420
         TabIndex        =   645
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1066
         Left            =   9660
         TabIndex        =   644
         Top             =   6060
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1067
         Left            =   60
         TabIndex        =   643
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1068
         Left            =   300
         TabIndex        =   642
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1069
         Left            =   540
         TabIndex        =   641
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1070
         Left            =   780
         TabIndex        =   640
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1071
         Left            =   1020
         TabIndex        =   639
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1072
         Left            =   1260
         TabIndex        =   638
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1073
         Left            =   1500
         TabIndex        =   637
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1074
         Left            =   1740
         TabIndex        =   636
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1075
         Left            =   1980
         TabIndex        =   635
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1076
         Left            =   2220
         TabIndex        =   634
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1077
         Left            =   2460
         TabIndex        =   633
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1078
         Left            =   2700
         TabIndex        =   632
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1079
         Left            =   2940
         TabIndex        =   631
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1080
         Left            =   3180
         TabIndex        =   630
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1081
         Left            =   3420
         TabIndex        =   629
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1082
         Left            =   3660
         TabIndex        =   628
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1083
         Left            =   3900
         TabIndex        =   627
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1084
         Left            =   4140
         TabIndex        =   626
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1085
         Left            =   4380
         TabIndex        =   625
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1086
         Left            =   4620
         TabIndex        =   624
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1087
         Left            =   4860
         TabIndex        =   623
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1088
         Left            =   5100
         TabIndex        =   622
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1089
         Left            =   5340
         TabIndex        =   621
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1090
         Left            =   5580
         TabIndex        =   620
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1091
         Left            =   5820
         TabIndex        =   619
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1092
         Left            =   6060
         TabIndex        =   618
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1093
         Left            =   6300
         TabIndex        =   617
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1094
         Left            =   6540
         TabIndex        =   616
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1095
         Left            =   6780
         TabIndex        =   615
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1096
         Left            =   7020
         TabIndex        =   614
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1097
         Left            =   7260
         TabIndex        =   613
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1098
         Left            =   7500
         TabIndex        =   612
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1099
         Left            =   7740
         TabIndex        =   611
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1100
         Left            =   7980
         TabIndex        =   610
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1101
         Left            =   8220
         TabIndex        =   609
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1102
         Left            =   8460
         TabIndex        =   608
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1103
         Left            =   8700
         TabIndex        =   607
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1104
         Left            =   8940
         TabIndex        =   606
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1105
         Left            =   9180
         TabIndex        =   605
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1106
         Left            =   9420
         TabIndex        =   604
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1107
         Left            =   9660
         TabIndex        =   603
         Top             =   6300
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1108
         Left            =   60
         TabIndex        =   602
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1109
         Left            =   300
         TabIndex        =   601
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1110
         Left            =   540
         TabIndex        =   600
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1111
         Left            =   780
         TabIndex        =   599
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1112
         Left            =   1020
         TabIndex        =   598
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1113
         Left            =   1260
         TabIndex        =   597
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1114
         Left            =   1500
         TabIndex        =   596
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1115
         Left            =   1740
         TabIndex        =   595
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1116
         Left            =   1980
         TabIndex        =   594
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1117
         Left            =   2220
         TabIndex        =   593
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1118
         Left            =   2460
         TabIndex        =   592
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1119
         Left            =   2700
         TabIndex        =   591
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1120
         Left            =   2940
         TabIndex        =   590
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1121
         Left            =   3180
         TabIndex        =   589
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1122
         Left            =   3420
         TabIndex        =   588
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1123
         Left            =   3660
         TabIndex        =   587
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1124
         Left            =   3900
         TabIndex        =   586
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1125
         Left            =   4140
         TabIndex        =   585
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1126
         Left            =   4380
         TabIndex        =   584
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1127
         Left            =   4620
         TabIndex        =   583
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1128
         Left            =   4860
         TabIndex        =   582
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1129
         Left            =   5100
         TabIndex        =   581
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1130
         Left            =   5340
         TabIndex        =   580
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1131
         Left            =   5580
         TabIndex        =   579
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1132
         Left            =   5820
         TabIndex        =   578
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1133
         Left            =   6060
         TabIndex        =   577
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1134
         Left            =   6300
         TabIndex        =   576
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1135
         Left            =   6540
         TabIndex        =   575
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1136
         Left            =   6780
         TabIndex        =   574
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1137
         Left            =   7020
         TabIndex        =   573
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1138
         Left            =   7260
         TabIndex        =   572
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1139
         Left            =   7500
         TabIndex        =   571
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1140
         Left            =   7740
         TabIndex        =   570
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1141
         Left            =   7980
         TabIndex        =   569
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1142
         Left            =   8220
         TabIndex        =   568
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1143
         Left            =   8460
         TabIndex        =   567
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1144
         Left            =   8700
         TabIndex        =   566
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1145
         Left            =   8940
         TabIndex        =   565
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1146
         Left            =   9180
         TabIndex        =   564
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1147
         Left            =   9420
         TabIndex        =   563
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1148
         Left            =   9660
         TabIndex        =   562
         Top             =   6540
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1149
         Left            =   60
         TabIndex        =   561
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1150
         Left            =   300
         TabIndex        =   560
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1151
         Left            =   540
         TabIndex        =   559
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1152
         Left            =   780
         TabIndex        =   558
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1153
         Left            =   1020
         TabIndex        =   557
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1154
         Left            =   1260
         TabIndex        =   556
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1155
         Left            =   1500
         TabIndex        =   555
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1156
         Left            =   1740
         TabIndex        =   554
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1157
         Left            =   1980
         TabIndex        =   553
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1158
         Left            =   2220
         TabIndex        =   552
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1159
         Left            =   2460
         TabIndex        =   551
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1160
         Left            =   2700
         TabIndex        =   550
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1161
         Left            =   2940
         TabIndex        =   549
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1162
         Left            =   3180
         TabIndex        =   548
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1163
         Left            =   3420
         TabIndex        =   547
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1164
         Left            =   3660
         TabIndex        =   546
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1165
         Left            =   3900
         TabIndex        =   545
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1166
         Left            =   4140
         TabIndex        =   544
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1167
         Left            =   4380
         TabIndex        =   543
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1168
         Left            =   4620
         TabIndex        =   542
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1169
         Left            =   4860
         TabIndex        =   541
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1170
         Left            =   5100
         TabIndex        =   540
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1171
         Left            =   5340
         TabIndex        =   539
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1172
         Left            =   5580
         TabIndex        =   538
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1173
         Left            =   5820
         TabIndex        =   537
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1174
         Left            =   6060
         TabIndex        =   536
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1175
         Left            =   6300
         TabIndex        =   535
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1176
         Left            =   6540
         TabIndex        =   534
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1177
         Left            =   6780
         TabIndex        =   533
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1178
         Left            =   7020
         TabIndex        =   532
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1179
         Left            =   7260
         TabIndex        =   531
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1180
         Left            =   7500
         TabIndex        =   530
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1181
         Left            =   7740
         TabIndex        =   529
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1182
         Left            =   7980
         TabIndex        =   528
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1183
         Left            =   8220
         TabIndex        =   527
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1184
         Left            =   8460
         TabIndex        =   526
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1185
         Left            =   8700
         TabIndex        =   525
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1186
         Left            =   8940
         TabIndex        =   524
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1187
         Left            =   9180
         TabIndex        =   523
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1188
         Left            =   9420
         TabIndex        =   522
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1189
         Left            =   9660
         TabIndex        =   521
         Top             =   6780
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1190
         Left            =   60
         TabIndex        =   520
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1191
         Left            =   300
         TabIndex        =   519
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1192
         Left            =   540
         TabIndex        =   518
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1193
         Left            =   780
         TabIndex        =   517
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1194
         Left            =   1020
         TabIndex        =   516
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1195
         Left            =   1260
         TabIndex        =   515
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1196
         Left            =   1500
         TabIndex        =   514
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1197
         Left            =   1740
         TabIndex        =   513
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1198
         Left            =   1980
         TabIndex        =   512
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1199
         Left            =   2220
         TabIndex        =   511
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1200
         Left            =   2460
         TabIndex        =   510
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1201
         Left            =   2700
         TabIndex        =   509
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1202
         Left            =   2940
         TabIndex        =   508
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1203
         Left            =   3180
         TabIndex        =   507
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1204
         Left            =   3420
         TabIndex        =   506
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1205
         Left            =   3660
         TabIndex        =   505
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1206
         Left            =   3900
         TabIndex        =   504
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1207
         Left            =   4140
         TabIndex        =   503
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1208
         Left            =   4380
         TabIndex        =   502
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1209
         Left            =   4620
         TabIndex        =   501
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1210
         Left            =   4860
         TabIndex        =   500
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1211
         Left            =   5100
         TabIndex        =   499
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1212
         Left            =   5340
         TabIndex        =   498
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1213
         Left            =   5580
         TabIndex        =   497
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1214
         Left            =   5820
         TabIndex        =   496
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1215
         Left            =   6060
         TabIndex        =   495
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1216
         Left            =   6300
         TabIndex        =   494
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1217
         Left            =   6540
         TabIndex        =   493
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1218
         Left            =   6780
         TabIndex        =   492
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1219
         Left            =   7020
         TabIndex        =   491
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1220
         Left            =   7260
         TabIndex        =   490
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1221
         Left            =   7500
         TabIndex        =   489
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1222
         Left            =   7740
         TabIndex        =   488
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1223
         Left            =   7980
         TabIndex        =   487
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1224
         Left            =   8220
         TabIndex        =   486
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1225
         Left            =   8460
         TabIndex        =   485
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1226
         Left            =   8700
         TabIndex        =   484
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1227
         Left            =   8940
         TabIndex        =   483
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1228
         Left            =   9180
         TabIndex        =   482
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1229
         Left            =   9420
         TabIndex        =   481
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1230
         Left            =   9660
         TabIndex        =   480
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1231
         Left            =   60
         TabIndex        =   479
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1232
         Left            =   300
         TabIndex        =   478
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1233
         Left            =   540
         TabIndex        =   477
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1234
         Left            =   780
         TabIndex        =   476
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1235
         Left            =   1020
         TabIndex        =   475
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1236
         Left            =   1260
         TabIndex        =   474
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1237
         Left            =   1500
         TabIndex        =   473
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1238
         Left            =   1740
         TabIndex        =   472
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1239
         Left            =   1980
         TabIndex        =   471
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1240
         Left            =   2220
         TabIndex        =   470
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1241
         Left            =   2460
         TabIndex        =   469
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1242
         Left            =   2700
         TabIndex        =   468
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1243
         Left            =   2940
         TabIndex        =   467
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1244
         Left            =   3180
         TabIndex        =   466
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1245
         Left            =   3420
         TabIndex        =   465
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1246
         Left            =   3660
         TabIndex        =   464
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1247
         Left            =   3900
         TabIndex        =   463
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1248
         Left            =   4140
         TabIndex        =   462
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1249
         Left            =   4380
         TabIndex        =   461
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1250
         Left            =   4620
         TabIndex        =   460
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1251
         Left            =   4860
         TabIndex        =   459
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1252
         Left            =   5100
         TabIndex        =   458
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1253
         Left            =   5340
         TabIndex        =   457
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1254
         Left            =   5580
         TabIndex        =   456
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1255
         Left            =   5820
         TabIndex        =   455
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1256
         Left            =   6060
         TabIndex        =   454
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1257
         Left            =   6300
         TabIndex        =   453
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1258
         Left            =   6540
         TabIndex        =   452
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1259
         Left            =   6780
         TabIndex        =   451
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1260
         Left            =   7020
         TabIndex        =   450
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1261
         Left            =   7260
         TabIndex        =   449
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1262
         Left            =   7500
         TabIndex        =   448
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1263
         Left            =   7740
         TabIndex        =   447
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1264
         Left            =   7980
         TabIndex        =   446
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1265
         Left            =   8220
         TabIndex        =   445
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1266
         Left            =   8460
         TabIndex        =   444
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1267
         Left            =   8700
         TabIndex        =   443
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1268
         Left            =   8940
         TabIndex        =   442
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1269
         Left            =   9180
         TabIndex        =   441
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1270
         Left            =   9420
         TabIndex        =   440
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1271
         Left            =   9660
         TabIndex        =   439
         Top             =   7260
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1272
         Left            =   60
         TabIndex        =   438
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1273
         Left            =   300
         TabIndex        =   437
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1274
         Left            =   540
         TabIndex        =   436
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1275
         Left            =   780
         TabIndex        =   435
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1276
         Left            =   1020
         TabIndex        =   434
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1277
         Left            =   1260
         TabIndex        =   433
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1278
         Left            =   1500
         TabIndex        =   432
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1279
         Left            =   1740
         TabIndex        =   431
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1280
         Left            =   1980
         TabIndex        =   430
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1281
         Left            =   2220
         TabIndex        =   429
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1282
         Left            =   2460
         TabIndex        =   428
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1283
         Left            =   2700
         TabIndex        =   427
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1284
         Left            =   2940
         TabIndex        =   426
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1285
         Left            =   3180
         TabIndex        =   425
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1286
         Left            =   3420
         TabIndex        =   424
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1287
         Left            =   3660
         TabIndex        =   423
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1288
         Left            =   3900
         TabIndex        =   422
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1289
         Left            =   4140
         TabIndex        =   421
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1290
         Left            =   4380
         TabIndex        =   420
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1291
         Left            =   4620
         TabIndex        =   419
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1292
         Left            =   4860
         TabIndex        =   418
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1293
         Left            =   5100
         TabIndex        =   417
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1294
         Left            =   5340
         TabIndex        =   416
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1295
         Left            =   5580
         TabIndex        =   415
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1296
         Left            =   5820
         TabIndex        =   414
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1297
         Left            =   6060
         TabIndex        =   413
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1298
         Left            =   6300
         TabIndex        =   412
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1299
         Left            =   6540
         TabIndex        =   411
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1300
         Left            =   6780
         TabIndex        =   410
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1301
         Left            =   7020
         TabIndex        =   409
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1302
         Left            =   7260
         TabIndex        =   408
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1303
         Left            =   7500
         TabIndex        =   407
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1304
         Left            =   7740
         TabIndex        =   406
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1305
         Left            =   7980
         TabIndex        =   405
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1306
         Left            =   8220
         TabIndex        =   404
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1307
         Left            =   8460
         TabIndex        =   403
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1308
         Left            =   8700
         TabIndex        =   402
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1309
         Left            =   8940
         TabIndex        =   401
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1310
         Left            =   9180
         TabIndex        =   400
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1311
         Left            =   9420
         TabIndex        =   399
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1312
         Left            =   9660
         TabIndex        =   398
         Top             =   7500
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1313
         Left            =   60
         TabIndex        =   397
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1314
         Left            =   300
         TabIndex        =   396
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1315
         Left            =   540
         TabIndex        =   395
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1316
         Left            =   780
         TabIndex        =   394
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1317
         Left            =   1020
         TabIndex        =   393
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1318
         Left            =   1260
         TabIndex        =   392
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1319
         Left            =   1500
         TabIndex        =   391
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1320
         Left            =   1740
         TabIndex        =   390
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1321
         Left            =   1980
         TabIndex        =   389
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1322
         Left            =   2220
         TabIndex        =   388
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1323
         Left            =   2460
         TabIndex        =   387
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1324
         Left            =   2700
         TabIndex        =   386
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1325
         Left            =   2940
         TabIndex        =   385
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1326
         Left            =   3180
         TabIndex        =   384
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1327
         Left            =   3420
         TabIndex        =   383
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1328
         Left            =   3660
         TabIndex        =   382
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1329
         Left            =   3900
         TabIndex        =   381
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1330
         Left            =   4140
         TabIndex        =   380
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1331
         Left            =   4380
         TabIndex        =   379
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1332
         Left            =   4620
         TabIndex        =   378
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1333
         Left            =   4860
         TabIndex        =   377
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1334
         Left            =   5100
         TabIndex        =   376
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1335
         Left            =   5340
         TabIndex        =   375
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1336
         Left            =   5580
         TabIndex        =   374
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1337
         Left            =   5820
         TabIndex        =   373
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1338
         Left            =   6060
         TabIndex        =   372
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1339
         Left            =   6300
         TabIndex        =   371
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1340
         Left            =   6540
         TabIndex        =   370
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1341
         Left            =   6780
         TabIndex        =   369
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1342
         Left            =   7020
         TabIndex        =   368
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1343
         Left            =   7260
         TabIndex        =   367
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1344
         Left            =   7500
         TabIndex        =   366
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1345
         Left            =   7740
         TabIndex        =   365
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1346
         Left            =   7980
         TabIndex        =   364
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1347
         Left            =   8220
         TabIndex        =   363
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1348
         Left            =   8460
         TabIndex        =   362
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1349
         Left            =   8700
         TabIndex        =   361
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1350
         Left            =   8940
         TabIndex        =   360
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1351
         Left            =   9180
         TabIndex        =   359
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1352
         Left            =   9420
         TabIndex        =   358
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1353
         Left            =   9660
         TabIndex        =   357
         Top             =   7740
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1354
         Left            =   60
         TabIndex        =   356
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1355
         Left            =   300
         TabIndex        =   355
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1356
         Left            =   540
         TabIndex        =   354
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1357
         Left            =   780
         TabIndex        =   353
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1358
         Left            =   1020
         TabIndex        =   352
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1359
         Left            =   1260
         TabIndex        =   351
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1360
         Left            =   1500
         TabIndex        =   350
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1361
         Left            =   1740
         TabIndex        =   349
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1362
         Left            =   1980
         TabIndex        =   348
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1363
         Left            =   2220
         TabIndex        =   347
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1364
         Left            =   2460
         TabIndex        =   346
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1365
         Left            =   2700
         TabIndex        =   345
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1366
         Left            =   2940
         TabIndex        =   344
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1367
         Left            =   3180
         TabIndex        =   343
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1368
         Left            =   3420
         TabIndex        =   342
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1369
         Left            =   3660
         TabIndex        =   341
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1370
         Left            =   3900
         TabIndex        =   340
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1371
         Left            =   4140
         TabIndex        =   339
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1372
         Left            =   4380
         TabIndex        =   338
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1373
         Left            =   4620
         TabIndex        =   337
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1374
         Left            =   4860
         TabIndex        =   336
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1375
         Left            =   5100
         TabIndex        =   335
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1376
         Left            =   5340
         TabIndex        =   334
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1377
         Left            =   5580
         TabIndex        =   333
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1378
         Left            =   5820
         TabIndex        =   332
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1379
         Left            =   6060
         TabIndex        =   331
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1380
         Left            =   6300
         TabIndex        =   330
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1381
         Left            =   6540
         TabIndex        =   329
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1382
         Left            =   6780
         TabIndex        =   328
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1383
         Left            =   7020
         TabIndex        =   327
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1384
         Left            =   7260
         TabIndex        =   326
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1385
         Left            =   7500
         TabIndex        =   325
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1386
         Left            =   7740
         TabIndex        =   324
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1387
         Left            =   7980
         TabIndex        =   323
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1388
         Left            =   8220
         TabIndex        =   322
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1389
         Left            =   8460
         TabIndex        =   321
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1390
         Left            =   8700
         TabIndex        =   320
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1391
         Left            =   8940
         TabIndex        =   319
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1392
         Left            =   9180
         TabIndex        =   318
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1393
         Left            =   9420
         TabIndex        =   317
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1394
         Left            =   9660
         TabIndex        =   316
         Top             =   7980
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1395
         Left            =   60
         TabIndex        =   315
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1396
         Left            =   300
         TabIndex        =   314
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1397
         Left            =   540
         TabIndex        =   313
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1398
         Left            =   780
         TabIndex        =   312
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1399
         Left            =   1020
         TabIndex        =   311
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1400
         Left            =   1260
         TabIndex        =   310
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1401
         Left            =   1500
         TabIndex        =   309
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1402
         Left            =   1740
         TabIndex        =   308
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1403
         Left            =   1980
         TabIndex        =   307
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1404
         Left            =   2220
         TabIndex        =   306
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1405
         Left            =   2460
         TabIndex        =   305
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1406
         Left            =   2700
         TabIndex        =   304
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1407
         Left            =   2940
         TabIndex        =   303
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1408
         Left            =   3180
         TabIndex        =   302
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1409
         Left            =   3420
         TabIndex        =   301
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1410
         Left            =   3660
         TabIndex        =   300
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1411
         Left            =   3900
         TabIndex        =   299
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1412
         Left            =   4140
         TabIndex        =   298
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1413
         Left            =   4380
         TabIndex        =   297
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1414
         Left            =   4620
         TabIndex        =   296
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1415
         Left            =   4860
         TabIndex        =   295
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1416
         Left            =   5100
         TabIndex        =   294
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1417
         Left            =   5340
         TabIndex        =   293
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1418
         Left            =   5580
         TabIndex        =   292
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1419
         Left            =   5820
         TabIndex        =   291
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1420
         Left            =   6060
         TabIndex        =   290
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1421
         Left            =   6300
         TabIndex        =   289
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1422
         Left            =   6540
         TabIndex        =   288
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1423
         Left            =   6780
         TabIndex        =   287
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1424
         Left            =   7020
         TabIndex        =   286
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1425
         Left            =   7260
         TabIndex        =   285
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1426
         Left            =   7500
         TabIndex        =   284
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1427
         Left            =   7740
         TabIndex        =   283
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1428
         Left            =   7980
         TabIndex        =   282
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1429
         Left            =   8220
         TabIndex        =   281
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1430
         Left            =   8460
         TabIndex        =   280
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1431
         Left            =   8700
         TabIndex        =   279
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1432
         Left            =   8940
         TabIndex        =   278
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1433
         Left            =   9180
         TabIndex        =   277
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1434
         Left            =   9420
         TabIndex        =   276
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1435
         Left            =   9660
         TabIndex        =   275
         Top             =   8220
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1436
         Left            =   60
         TabIndex        =   274
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1437
         Left            =   300
         TabIndex        =   273
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1438
         Left            =   540
         TabIndex        =   272
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1439
         Left            =   780
         TabIndex        =   271
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1440
         Left            =   1020
         TabIndex        =   270
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1441
         Left            =   1260
         TabIndex        =   269
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1442
         Left            =   1500
         TabIndex        =   268
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1443
         Left            =   1740
         TabIndex        =   267
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1444
         Left            =   1980
         TabIndex        =   266
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1445
         Left            =   2220
         TabIndex        =   265
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1446
         Left            =   2460
         TabIndex        =   264
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1447
         Left            =   2700
         TabIndex        =   263
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1448
         Left            =   2940
         TabIndex        =   262
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1449
         Left            =   3180
         TabIndex        =   261
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1450
         Left            =   3420
         TabIndex        =   260
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1451
         Left            =   3660
         TabIndex        =   259
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1452
         Left            =   3900
         TabIndex        =   258
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1453
         Left            =   4140
         TabIndex        =   257
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1454
         Left            =   4380
         TabIndex        =   256
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1455
         Left            =   4620
         TabIndex        =   255
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1456
         Left            =   4860
         TabIndex        =   254
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1457
         Left            =   5100
         TabIndex        =   253
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1458
         Left            =   5340
         TabIndex        =   252
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1459
         Left            =   5580
         TabIndex        =   251
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1460
         Left            =   5820
         TabIndex        =   250
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1461
         Left            =   6060
         TabIndex        =   249
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1462
         Left            =   6300
         TabIndex        =   248
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1463
         Left            =   6540
         TabIndex        =   247
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1464
         Left            =   6780
         TabIndex        =   246
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1465
         Left            =   7020
         TabIndex        =   245
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1466
         Left            =   7260
         TabIndex        =   244
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1467
         Left            =   7500
         TabIndex        =   243
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1468
         Left            =   7740
         TabIndex        =   242
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1469
         Left            =   7980
         TabIndex        =   241
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1470
         Left            =   8220
         TabIndex        =   240
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1471
         Left            =   8460
         TabIndex        =   239
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1472
         Left            =   8700
         TabIndex        =   238
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1473
         Left            =   8940
         TabIndex        =   237
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1474
         Left            =   9180
         TabIndex        =   236
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1475
         Left            =   9420
         TabIndex        =   235
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1476
         Left            =   9660
         TabIndex        =   234
         Top             =   8460
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1477
         Left            =   60
         TabIndex        =   233
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1478
         Left            =   300
         TabIndex        =   232
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1479
         Left            =   540
         TabIndex        =   231
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1480
         Left            =   780
         TabIndex        =   230
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1481
         Left            =   1020
         TabIndex        =   229
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1482
         Left            =   1260
         TabIndex        =   228
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1483
         Left            =   1500
         TabIndex        =   227
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1484
         Left            =   1740
         TabIndex        =   226
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1485
         Left            =   1980
         TabIndex        =   225
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1486
         Left            =   2220
         TabIndex        =   224
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1487
         Left            =   2460
         TabIndex        =   223
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1488
         Left            =   2700
         TabIndex        =   222
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1489
         Left            =   2940
         TabIndex        =   221
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1490
         Left            =   3180
         TabIndex        =   220
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1491
         Left            =   3420
         TabIndex        =   219
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1492
         Left            =   3660
         TabIndex        =   218
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1493
         Left            =   3900
         TabIndex        =   217
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1494
         Left            =   4140
         TabIndex        =   216
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1495
         Left            =   4380
         TabIndex        =   215
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1496
         Left            =   4620
         TabIndex        =   214
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1497
         Left            =   4860
         TabIndex        =   213
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1498
         Left            =   5100
         TabIndex        =   212
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1499
         Left            =   5340
         TabIndex        =   211
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1500
         Left            =   5580
         TabIndex        =   210
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1501
         Left            =   5820
         TabIndex        =   209
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1502
         Left            =   6060
         TabIndex        =   208
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1503
         Left            =   6300
         TabIndex        =   207
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1504
         Left            =   6540
         TabIndex        =   206
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1505
         Left            =   6780
         TabIndex        =   205
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1506
         Left            =   7020
         TabIndex        =   204
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1507
         Left            =   7260
         TabIndex        =   203
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1508
         Left            =   7500
         TabIndex        =   202
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1509
         Left            =   7740
         TabIndex        =   201
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1510
         Left            =   7980
         TabIndex        =   200
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1511
         Left            =   8220
         TabIndex        =   199
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1512
         Left            =   8460
         TabIndex        =   198
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1513
         Left            =   8700
         TabIndex        =   197
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1514
         Left            =   8940
         TabIndex        =   196
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1515
         Left            =   9180
         TabIndex        =   195
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1516
         Left            =   9420
         TabIndex        =   194
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1517
         Left            =   9660
         TabIndex        =   193
         Top             =   8700
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1518
         Left            =   60
         TabIndex        =   192
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1519
         Left            =   300
         TabIndex        =   191
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1520
         Left            =   540
         TabIndex        =   190
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1521
         Left            =   780
         TabIndex        =   189
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1522
         Left            =   1020
         TabIndex        =   188
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1523
         Left            =   1260
         TabIndex        =   187
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1524
         Left            =   1500
         TabIndex        =   186
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1525
         Left            =   1740
         TabIndex        =   185
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1526
         Left            =   1980
         TabIndex        =   184
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1527
         Left            =   2220
         TabIndex        =   183
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1528
         Left            =   2460
         TabIndex        =   182
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1529
         Left            =   2700
         TabIndex        =   181
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1530
         Left            =   2940
         TabIndex        =   180
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1531
         Left            =   3180
         TabIndex        =   179
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1532
         Left            =   3420
         TabIndex        =   178
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1533
         Left            =   3660
         TabIndex        =   177
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1534
         Left            =   3900
         TabIndex        =   176
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1535
         Left            =   4140
         TabIndex        =   175
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1536
         Left            =   4380
         TabIndex        =   174
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1537
         Left            =   4620
         TabIndex        =   173
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1538
         Left            =   4860
         TabIndex        =   172
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1539
         Left            =   5100
         TabIndex        =   171
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1540
         Left            =   5340
         TabIndex        =   170
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1541
         Left            =   5580
         TabIndex        =   169
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1542
         Left            =   5820
         TabIndex        =   168
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1543
         Left            =   6060
         TabIndex        =   167
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1544
         Left            =   6300
         TabIndex        =   166
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1545
         Left            =   6540
         TabIndex        =   165
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1546
         Left            =   6780
         TabIndex        =   164
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1547
         Left            =   7020
         TabIndex        =   163
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1548
         Left            =   7260
         TabIndex        =   162
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1549
         Left            =   7500
         TabIndex        =   161
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1550
         Left            =   7740
         TabIndex        =   160
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1551
         Left            =   7980
         TabIndex        =   159
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1552
         Left            =   8220
         TabIndex        =   158
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1553
         Left            =   8460
         TabIndex        =   157
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1554
         Left            =   8700
         TabIndex        =   156
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1555
         Left            =   8940
         TabIndex        =   155
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1556
         Left            =   9180
         TabIndex        =   154
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1557
         Left            =   9420
         TabIndex        =   153
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1558
         Left            =   9660
         TabIndex        =   152
         Top             =   8940
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1559
         Left            =   60
         TabIndex        =   151
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1560
         Left            =   300
         TabIndex        =   150
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1561
         Left            =   540
         TabIndex        =   149
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1562
         Left            =   780
         TabIndex        =   148
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1563
         Left            =   1020
         TabIndex        =   147
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1564
         Left            =   1260
         TabIndex        =   146
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1565
         Left            =   1500
         TabIndex        =   145
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1566
         Left            =   1740
         TabIndex        =   144
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1567
         Left            =   1980
         TabIndex        =   143
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1568
         Left            =   2220
         TabIndex        =   142
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1569
         Left            =   2460
         TabIndex        =   141
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1570
         Left            =   2700
         TabIndex        =   140
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1571
         Left            =   2940
         TabIndex        =   139
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1572
         Left            =   3180
         TabIndex        =   138
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1573
         Left            =   3420
         TabIndex        =   137
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1574
         Left            =   3660
         TabIndex        =   136
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1575
         Left            =   3900
         TabIndex        =   135
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1576
         Left            =   4140
         TabIndex        =   134
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1577
         Left            =   4380
         TabIndex        =   133
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1578
         Left            =   4620
         TabIndex        =   132
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1579
         Left            =   4860
         TabIndex        =   131
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1580
         Left            =   5100
         TabIndex        =   130
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1581
         Left            =   5340
         TabIndex        =   129
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1582
         Left            =   5580
         TabIndex        =   128
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1583
         Left            =   5820
         TabIndex        =   127
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1584
         Left            =   6060
         TabIndex        =   126
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1585
         Left            =   6300
         TabIndex        =   125
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1586
         Left            =   6540
         TabIndex        =   124
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1587
         Left            =   6780
         TabIndex        =   123
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1588
         Left            =   7020
         TabIndex        =   122
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1589
         Left            =   7260
         TabIndex        =   121
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1590
         Left            =   7500
         TabIndex        =   120
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1591
         Left            =   7740
         TabIndex        =   119
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1592
         Left            =   7980
         TabIndex        =   118
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1593
         Left            =   8220
         TabIndex        =   117
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1594
         Left            =   8460
         TabIndex        =   116
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1595
         Left            =   8700
         TabIndex        =   115
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1596
         Left            =   8940
         TabIndex        =   114
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1597
         Left            =   9180
         TabIndex        =   113
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1598
         Left            =   9420
         TabIndex        =   112
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1599
         Left            =   9660
         TabIndex        =   111
         Top             =   9180
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1600
         Left            =   60
         TabIndex        =   110
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1601
         Left            =   300
         TabIndex        =   109
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1602
         Left            =   540
         TabIndex        =   108
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1603
         Left            =   780
         TabIndex        =   107
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1604
         Left            =   1020
         TabIndex        =   106
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1605
         Left            =   1260
         TabIndex        =   105
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1606
         Left            =   1500
         TabIndex        =   104
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1607
         Left            =   1740
         TabIndex        =   103
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1608
         Left            =   1980
         TabIndex        =   102
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1609
         Left            =   2220
         TabIndex        =   101
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1610
         Left            =   2460
         TabIndex        =   100
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1611
         Left            =   2700
         TabIndex        =   99
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1612
         Left            =   2940
         TabIndex        =   98
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1613
         Left            =   3180
         TabIndex        =   97
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1614
         Left            =   3420
         TabIndex        =   96
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1615
         Left            =   3660
         TabIndex        =   95
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1616
         Left            =   3900
         TabIndex        =   94
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1617
         Left            =   4140
         TabIndex        =   93
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1618
         Left            =   4380
         TabIndex        =   92
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1619
         Left            =   4620
         TabIndex        =   91
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1620
         Left            =   4860
         TabIndex        =   90
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1621
         Left            =   5100
         TabIndex        =   89
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1622
         Left            =   5340
         TabIndex        =   88
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1623
         Left            =   5580
         TabIndex        =   87
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1624
         Left            =   5820
         TabIndex        =   86
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1625
         Left            =   6060
         TabIndex        =   85
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1626
         Left            =   6300
         TabIndex        =   84
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1627
         Left            =   6540
         TabIndex        =   83
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1628
         Left            =   6780
         TabIndex        =   82
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1629
         Left            =   7020
         TabIndex        =   81
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1630
         Left            =   7260
         TabIndex        =   80
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1631
         Left            =   7500
         TabIndex        =   79
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1632
         Left            =   7740
         TabIndex        =   78
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1633
         Left            =   7980
         TabIndex        =   77
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1634
         Left            =   8220
         TabIndex        =   76
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1635
         Left            =   8460
         TabIndex        =   75
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1636
         Left            =   8700
         TabIndex        =   74
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1637
         Left            =   8940
         TabIndex        =   73
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1638
         Left            =   9180
         TabIndex        =   72
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1639
         Left            =   9420
         TabIndex        =   71
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1640
         Left            =   9660
         TabIndex        =   70
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1641
         Left            =   60
         TabIndex        =   69
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1642
         Left            =   300
         TabIndex        =   68
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1643
         Left            =   540
         TabIndex        =   67
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1644
         Left            =   780
         TabIndex        =   66
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1645
         Left            =   1020
         TabIndex        =   65
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1646
         Left            =   1260
         TabIndex        =   64
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1647
         Left            =   1500
         TabIndex        =   63
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1648
         Left            =   1740
         TabIndex        =   62
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1649
         Left            =   1980
         TabIndex        =   61
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1650
         Left            =   2220
         TabIndex        =   60
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1651
         Left            =   2460
         TabIndex        =   59
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1652
         Left            =   2700
         TabIndex        =   58
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1653
         Left            =   2940
         TabIndex        =   57
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1654
         Left            =   3180
         TabIndex        =   56
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1655
         Left            =   3420
         TabIndex        =   55
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1656
         Left            =   3660
         TabIndex        =   54
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1657
         Left            =   3900
         TabIndex        =   53
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1658
         Left            =   4140
         TabIndex        =   52
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1659
         Left            =   4380
         TabIndex        =   51
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1660
         Left            =   4620
         TabIndex        =   50
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1661
         Left            =   4860
         TabIndex        =   49
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1662
         Left            =   5100
         TabIndex        =   48
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1663
         Left            =   5340
         TabIndex        =   47
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1664
         Left            =   5580
         TabIndex        =   46
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1665
         Left            =   5820
         TabIndex        =   45
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1666
         Left            =   6060
         TabIndex        =   44
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1667
         Left            =   6300
         TabIndex        =   43
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1668
         Left            =   6540
         TabIndex        =   42
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1669
         Left            =   6780
         TabIndex        =   41
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1670
         Left            =   7020
         TabIndex        =   40
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1671
         Left            =   7260
         TabIndex        =   39
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1672
         Left            =   7500
         TabIndex        =   38
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1673
         Left            =   7740
         TabIndex        =   37
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1674
         Left            =   7980
         TabIndex        =   36
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1675
         Left            =   8220
         TabIndex        =   35
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1676
         Left            =   8460
         TabIndex        =   34
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1677
         Left            =   8700
         TabIndex        =   33
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1678
         Left            =   8940
         TabIndex        =   32
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1679
         Left            =   9180
         TabIndex        =   31
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1680
         Left            =   9420
         TabIndex        =   30
         Top             =   9660
         Width           =   135
      End
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1681
         Left            =   9660
         TabIndex        =   29
         Top             =   9660
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1320
      TabIndex        =   1
      ToolTipText     =   "Options"
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdExportRooms 
      Caption         =   "Export Rooms"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3900
      TabIndex        =   3
      ToolTipText     =   "Export Rooms"
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdReload 
      Caption         =   "&Reload Map"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "Reload Map"
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdHideUnused 
      Caption         =   "Show &Blocks"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2580
      TabIndex        =   2
      ToolTipText     =   "Hide/Show Blocks"
      Top             =   0
      Width           =   1155
   End
   Begin VB.CommandButton cmdLegend 
      Caption         =   "&Legend / Help"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5160
      TabIndex        =   4
      ToolTipText     =   "Show Legend Window"
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Private Enum DrawRoomEnum
    drSquare = 0
    drStar = 1
    drOpenCircle = 2
    drUp = 3
    drDown = 4
    drCircle = 5
    drLineN = 6
    drLineS = 7
    drLineE = 8
    drLineW = 9
    drLineNE = 10
    drLineNW = 11
    drLineSE = 12
    drLineSW = 13
End Enum

Public nLastMapClick As Long
Public nLastRoomClick As Long

Dim SECorner As Integer
Dim RowLength As Integer

Dim DB As Database
Dim tabRooms As Recordset
Dim tabInfo As Recordset
Dim bUpdateExistingADB As Boolean
Dim sDataSource As String
Dim sExportPath As String

Dim nBackcolor As Long
Dim StartRoom As Long
Dim StartMap As Long
Dim CenterCell As Integer
Dim objTooltip As clsToolTip

Dim CellRoom(1 To 1681, 1 To 2) As Long
'Dim CellLabel(1681) As String
Dim UnchartedCells(1 To 1681) As Integer
Dim StopBuild As Boolean

Private Sub Check1_Click()

End Sub

Private Sub chkMarkAux_Click()
If chkMarkAux.Value = 1 Then
    optMarkAux(0).Enabled = True
    optMarkAux(1).Enabled = True
    optMarkAux(2).Enabled = True
    optMarkAux(3).Enabled = True
Else
    optMarkAux(0).Enabled = False
    optMarkAux(1).Enabled = False
    optMarkAux(2).Enabled = False
    optMarkAux(3).Enabled = False
End If

Call optMarkAux_Click(0)

End Sub

Private Sub cmdBuildControlRoomList_Click()
On Error GoTo error:
Dim nStatus As Integer, nYesNo As Integer

nYesNo = MsgBox("Control rooms will be found as maps are drawn and rooms are pulled via other means.  If you would like a complete accurate list click yes to continue and scan all of the rooms.  Otherwise click no and just redraw an area a couple times to let it find the control rooms (in the immediate area).", vbInformation + vbYesNo)
If Not nYesNo = vbYes Then Exit Sub

frmMain.Enabled = False
Call BuildControlRoomList

kill:
frmMain.Enabled = True
If Not bStopControlBuild Then Call cmdReload_Click
Exit Sub
error:
Call HandleError
bStopControlBuild = True
Resume kill:
End Sub

Private Sub Form_Load()
On Error GoTo error:

Set objTooltip = New clsToolTip
With objTooltip
    .DelayTime = 20
    .VisibleTime = 20000
    .BkColor = &HC0FFFF
    .txtColor = &H0
    .Style = 1 'ttStyleBalloon
    '.Style = ttStyleStandard
End With

Me.Top = ReadINI("Windows", "MapTop")
Me.Left = ReadINI("Windows", "MapLeft")

'Me.ScaleMode = vbPixels
'Me.ScaleHeight = 672
'Me.ScaleWidth = 656
'
picMap.ScaleMode = vbPixels
picMap.ScaleHeight = 657
picMap.ScaleWidth = 657

cmbMapSize.ListIndex = ReadINI("Options", "MapSize")
chkNoColors.Value = ReadINI("Options", "MapNoColors")
chkNoLineColors.Value = ReadINI("Options", "MapNoLineColors")
chkDontMarkStart.Value = ReadINI("Options", "MapDontMarkStart")
chkFollowMapChanges.Value = ReadINI("Options", "MapFollowMapChanges")
chkDontFollowHidden.Value = ReadINI("Options", "MapDontFollowHidden")
chkMarkLair.Value = ReadINI("Options", "MapMarkLair")
chkMarkCMD.Value = ReadINI("Options", "MapMarkCMD")
chkMarkNPC.Value = ReadINI("Options", "MapMarkNPC")
chkUseLastExport.Value = ReadINI("Options", "MapUseSameExport")
chkMonsterRegen.Value = ReadINI("Options", "MapLookUpMonsterRegen")
chkUseWhiteBG.Value = ReadINI("Options", "MapWhiteBG")
chkNoTooltips.Value = ReadINI("Options", "MapNoTooltips")

chkMarkAux.Value = ReadINI("Options", "MapMarkAux")
If Val(ReadINI("Options", "MapMarkAuxSpells")) > 0 Then
    optMarkAux(0).Value = True
ElseIf Val(ReadINI("Options", "MapMarkAuxShops")) > 0 Then
    optMarkAux(1).Value = True
ElseIf Val(ReadINI("Options", "MapMarkAuxRooms")) > 0 Then
    optMarkAux(2).Value = True
ElseIf Val(ReadINI("Options", "MapMarkAuxControl")) > 0 Then
    optMarkAux(3).Value = True
End If

Call cmbMapSize_Change

Call StartMapping

If StopBuild = True Then GoTo Cancel:

Me.Show
Me.SetFocus

cmdReload.SetFocus
'Call cmdHideUnused_Click

Exit Sub

Cancel:

Exit Sub

error:
Call HandleError
Resume Next
End Sub

Private Sub cmbMapSize_Change()

Select Case cmbMapSize.ListIndex
    Case 0: '41x41
        CenterCell = 881
        
    Case 1: '41x19
        CenterCell = 389
        
    Case 2: '30x30
        CenterCell = 630
    
    Case 3: '21x21
        CenterCell = 421
        
End Select
End Sub

Private Sub cmbMapSize_Click()
Call cmbMapSize_Change
End Sub

Private Sub cmdHideUnused_Click()
Dim x As Integer

If cmdHideUnused.Caption = "Hide &Blocks" Then
    For x = 1 To 1681
        If CellRoom(x, 1) = 0 Then lblRoomCell(x).Visible = False
    Next x
    cmdHideUnused.Caption = "Show &Blocks"
Else
    For x = 1 To 1681
        lblRoomCell(x).Visible = True
    Next x
    cmdHideUnused.Caption = "Hide &Blocks"
End If

On Error Resume Next
If Me.Visible Then cmdReload.SetFocus

End Sub


Private Sub cmdExportRooms_Click()
On Error GoTo error:
Dim nTemp As Integer, sNewPath() As String, x As Integer

frmMain.Enabled = False
ProgressBar.Value = 0
ProgressBar.Max = SECorner

nTemp = CreateDatabase

bUpdateExistingADB = False

Select Case nTemp
    Case 3: 'cancel
        GoTo ReEnable:
    Case 2: 'yes (update existing)
        bUpdateExistingADB = True
    Case 1: 'no (create new)
        
        lblExportCount.Caption = "Creating Tables..."
        If eDatFileVersion >= v111j Then
            nTemp = CreateAccessTables(sDataSource, True)
        Else
            nTemp = CreateAccessTables(sDataSource, False)
        End If
        
        If nTemp = False Then
            MsgBox "Access Database was not created successfully."
            GoTo ReEnable:
        End If
    Case Else: 'else
        MsgBox "Access Database was not created successfully."
        GoTo ReEnable:
End Select

If InStr(1, sDataSource, "\") > 0 Then
    sNewPath = Split(sDataSource, "\")
    sExportPath = sNewPath(LBound(sNewPath()))
    For x = LBound(sNewPath()) + 1 To UBound(sNewPath()) - 1
        sExportPath = sExportPath & "\" & sNewPath(x)
    Next x
    MsgBox sExportPath
    Call WriteINI("Options", "ExportPath", sExportPath)
End If
Erase sNewPath()

If cmbMapSize.ListIndex = 3 Then '21x21
    framExporting.Left = 40
    framExporting.Top = 72
Else
    framExporting.Left = 192
    framExporting.Top = 80
End If

lblExportCount.Caption = 1 & " / " & SECorner

frmMain.Enabled = False
ProgressBar.Value = 0
ProgressBar.Max = SECorner
framOptions.Visible = False
framExporting.Visible = True
DoEvents

Set DB = OpenDatabase(sDataSource)
Set tabRooms = DB.OpenRecordset("Rooms")
Set tabInfo = DB.OpenRecordset("Info")

If bUpdateExistingADB = True Then
    If CheckVersion = False Then
        Call CloseAll(True)
        GoTo ReEnable:
    End If
End If

DoEvents

Call ExportRooms
Call ExportVersionInfo

ProgressBar.Value = ProgressBar.Max

Call CloseAll

MsgBox "Export Complete.", vbInformation

ReEnable:
On Error Resume Next
Erase sNewPath()
frmMain.Enabled = True
framExporting.Visible = False
Exit Sub

error:
Call HandleError
Call CloseAll(True)
Resume ReEnable:

End Sub

Private Sub ExportRooms()
Dim nStatus As Integer, x As Integer, nCell As Integer, nYesNo As Integer, bSkipErrors As Boolean

nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Could not get first room record, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If


For nCell = 1 To SECorner
    
    lblExportCount.Caption = nCell & " / " & SECorner
    DoEvents
    If CellRoom(nCell, 1) = 0 Or CellRoom(nCell, 2) = 0 Then GoTo NextRoom:
    
    RoomKeyStruct.MapNum = CellRoom(nCell, 1)
    RoomKeyStruct.RoomNum = CellRoom(nCell, 2)

    nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), RoomKeyStruct, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If bSkipErrors Then GoTo NextRoom:
        
        nYesNo = MsgBox("Error retrieving room " & CellRoom(nCell, 1) & "/" & CellRoom(nCell, 2) & ": " & BtrieveErrorCode(nStatus, True) _
            & vbCrLf & "Keep exporting and skip remaining errors?", vbOKCancel + vbQuestion)
        
        If nYesNo = vbOK Then
            bSkipErrors = True
            GoTo NextRoom:
        Else
            Exit Sub
        End If
    End If

    Call RoomRowToStruct(Roomdatabuf.buf)

    If bUpdateExistingADB = True Then
        If tabRooms.RecordCount = 0 Then
            tabRooms.AddNew
        Else
            tabRooms.Index = "idxRooms"
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
    tabRooms.Fields("Spell") = Roomrec.Spell
    tabRooms.Fields("Exit Room") = Roomrec.ExitRoom
    tabRooms.Fields("Attributes") = Roomrec.Attributes
    For x = 0 To 6
        tabRooms.Fields("Desc " & x) = Roomrec.Desc(x)
    Next
    For x = 0 To 16
        tabRooms.Fields("Room Item " & x) = Roomrec.RoomItems(x)
        tabRooms.Fields("Room Item " & x & " QTY") = Roomrec.RoomItemQty(x)
        tabRooms.Fields("Room Item " & x & " USES") = Roomrec.RoomItemUses(x)
    Next
    For x = 0 To 14
        tabRooms.Fields("Hidden Item " & x) = Roomrec.InvisItems(x)
        tabRooms.Fields("Hidden Item " & x & " QTY") = Roomrec.InvisItemQty(x)
        tabRooms.Fields("Hidden Item " & x & " USES") = Roomrec.InvisItemUses(x)
        tabRooms.Fields("CurrentRoomMon " & x) = Roomrec.CurrentRoomMon(x)
    Next
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

NextRoom:
    ProgressBar.Value = nCell
    If Not bUseCPU Then DoEvents
Next

End Sub

Private Sub ExportVersionInfo()
On Error GoTo error:

tabInfo.AddNew
tabInfo.Fields("NMR Version") = sAppVersion
tabInfo.Fields("Dat File Version") = FriendlyDatVersion(eDatFileVersion)
tabInfo.Fields("Date") = Date
tabInfo.Fields("Time") = Time
tabInfo.Fields("Custom") = ""
tabInfo.Update

Exit Sub
error:
Call HandleError
End Sub

Private Sub CloseAll(Optional DontCompact As Boolean)
On Error Resume Next
Dim temp As String
Dim fso As FileSystemObject

tabRooms.Close
tabInfo.Close

DB.Close

Set tabRooms = Nothing
Set tabInfo = Nothing

Set DB = Nothing

If DontCompact Then GoTo finish:

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

finish:
Set fso = Nothing
DoEvents

End Sub

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
    nYesNo = MsgBox("Warning, current NMR Version/Dat File Version does not match the export file versions." & vbCrLf _
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

Private Function CreateDatabase() As Integer
On Error GoTo error:
Dim S As String, nYesNo As Integer, catDB As ADOX.Catalog
Dim fso As FileSystemObject, x As Integer, y As Integer, nTemp As Integer

'0=not created
'1=created ok
'2=update existing
'3=cancel

CreateDatabase = 0

Set fso = CreateObject("Scripting.FileSystemObject")

If chkUseLastExport.Value = 1 Then
    If fso.FileExists(sDataSource) Then
        CreateDatabase = 2
        Set fso = Nothing
        Exit Function
    End If
End If

sExportPath = ReadINI("Options", "ExportPath")
If Not fso.FolderExists(sExportPath) Then sExportPath = App.Path

CommonDialog1.Filter = "MDB Files (*.mdb)|*.mdb"
CommonDialog1.DialogTitle = "Select Export File/Enter New File Name"
CommonDialog1.FileName = "NMR-DataExport.mdb"
CommonDialog1.InitDir = sExportPath

On Error GoTo canceled:
CommonDialog1.ShowSave
If CommonDialog1.FileName = "" Then GoTo canceled:

On Error GoTo error:
sDataSource = CommonDialog1.FileName

If Not LCase(Right(sDataSource, 4)) = ".mdb" Then sDataSource = sDataSource & ".mdb"

If fso.FileExists(sDataSource) = True Then
    nYesNo = MsgBox("'" & sDataSource & "' already exists." & vbCrLf & vbCrLf _
        & "Attempt to add to and/or update it?" & vbCrLf & vbCrLf _
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

'get just the file name from the end of path
'nTemp = 0
'For x = 1 To Len(sDataSource)
'     y = InStr(x, sDataSource, "\")
'     If Not y = 0 Then nTemp = y
'Next x
'If Not nTemp = 0 Then
'    sExportPath = Left(sDataSource, nTemp - 1)
'    Call WriteINI("Options", "ExportPath", sExportPath)
'End If


bUpdateExistingADB = False

'create database
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
Resume canceled:
End Function

Private Sub cmdOptions_Click()
If framOptions.Visible = True Then
    framOptions.Visible = False
    Exit Sub
End If

framOptions.Visible = True

cmdReload.SetFocus
End Sub

Private Sub cmdLegend_Click()
cmdReload.SetFocus
If FormIsLoaded("frmMapLegend") Then
    Unload frmMapLegend
Else
    frmMapLegend.Show
    frmMapLegend.SetFocus
End If
End Sub

Public Sub cmdReload_Click()
On Error Resume Next
frmRoom.Show
cmdReload.SetFocus
Call StartMapping
Me.SetFocus
End Sub
Private Sub SetTitleBar()
'Dim TitleInfo As TITLEBARINFO, temp As cnWin32Ver
'
'temp = Win32Ver
'If temp = win95 Then GoTo win95:
'
'TitleInfo.cbSize = Len(TitleInfo)
'GetTitleBarInfo Me.hWnd, TitleInfo
'
'If chkSmallMap.value = 1 Then
'    If CenterCell > 778 Then CenterCell = 389
'    With TitleInfo.rcTitleBar
'        Me.Height = (CStr(.Bottom) * 15) - (CStr(.Top) * 15) + 4845
'    End With
'    shpBorder(1).Visible = True
'    shpBorder(0).Visible = False
'Else
'    With TitleInfo.rcTitleBar
'        Me.Height = (CStr(.Bottom) * 15) - (CStr(.Top) * 15) + 10125
'    End With
'    shpBorder(1).Visible = False
'    shpBorder(0).Visible = True
'End If
'
'Exit Sub
'
'win95:
'If chkSmallMap.value = 1 Then
'    If CenterCell > 778 Then CenterCell = 389
'    Me.Height = 5145
'    shpBorder(1).Visible = True
'    shpBorder(0).Visible = False
'Else
'    Me.Height = 10425
'    shpBorder(1).Visible = False
'    shpBorder(0).Visible = True
'End If

End Sub

Public Sub StartMapping()
On Error GoTo error:
Dim x As Integer, nMapSize As Integer, bCheckAgain As Boolean, y As Integer

picMap.Cls
picMap.Visible = False
frmMain.Enabled = False

If chkUseWhiteBG.Value = 1 Then
    nBackcolor = &H808080
    Me.BackColor = &HFFFFFF
    picMap.BackColor = &HFFFFFF
Else
    nBackcolor = &HC0C0C0
    Me.BackColor = &H0
    picMap.BackColor = &H0
End If

For x = 1 To 1681
    objTooltip.DelToolTip picMap.hwnd, 0
    lblRoomCell(x).BackColor = nBackcolor
    lblRoomCell(x).Visible = True
    lblRoomCell(x).Tag = 0
    UnchartedCells(x) = 0
    CellRoom(x, 1) = 0
    CellRoom(x, 2) = 0
Next x


Call SetTitleBar

framOptions.Visible = False

Select Case cmbMapSize.ListIndex
    Case 0: '41x41
        SECorner = 1681
        RowLength = 41
        If CenterCell > SECorner Then CenterCell = 881
        Me.Width = 9930
        Me.Height = 10440 + TITLEBAR_OFFSET
        
    Case 1: '41x19
        SECorner = 779
        RowLength = 41
        If CenterCell > SECorner Then CenterCell = 389
        Me.Width = 9930
        Me.Height = 5160 + TITLEBAR_OFFSET
        
    Case 2: '30x30
        SECorner = 1219
        RowLength = 30
        If CenterCell > SECorner Then CenterCell = 630
        Me.Height = 7815 + TITLEBAR_OFFSET
        Me.Width = 7275
    
    Case 3: '21x21
        SECorner = 841
        RowLength = 21
        If CenterCell > SECorner Then CenterCell = 421
        Me.Height = 5635 + TITLEBAR_OFFSET
        Me.Width = 5130
        
    Case Else: '30x30
        SECorner = 1219
        RowLength = 30
        If CenterCell > SECorner Then CenterCell = 630
        Me.Height = 7800 + TITLEBAR_OFFSET
        Me.Width = 7290
        
End Select

StopBuild = False

StartRoom = frmRoom.txtRoom
StartMap = frmRoom.txtMap

'CellLabel(CenterCell) = StartMap & "," & StartRoom
CellRoom(CenterCell, 1) = StartMap
CellRoom(CenterCell, 2) = StartRoom

Call MapExits(CenterCell, StartRoom, StartMap)

frmMain.Enabled = False
frmProgressBar.sCaption = "Map Builder"
frmProgressBar.lblCaption = "Please wait, " & vbCrLf & "Creating Map..."
frmProgressBar.cmdCancel.Enabled = True
Call frmProgressBar.SetRange(SECorner)
frmProgressBar.lblPanel(0).Caption = "Current Map,Room:"
frmProgressBar.Show

DoEvents
again:
bCheckAgain = False
For x = 1 To SECorner
    If StopBuild = True Then GoTo Cancel:
    If UnchartedCells(x) = 1 Then
        For y = 1 To SECorner
        
            If Not CellRoom(x, 1) = 0 Then
            
                If Not x = y Then
                
                    If CellRoom(y, 2) = CellRoom(x, 2) Then
                    
                        If CellRoom(y, 1) = CellRoom(x, 1) Then
                            CellRoom(x, 2) = 0
                            CellRoom(x, 1) = 0
                            UnchartedCells(x) = 0
                            GoTo skiproom:
                        End If
                    End If
                End If
            End If
            
        Next y
        Call MapExits(x, CellRoom(x, 2), CellRoom(x, 1))
skiproom:
        bCheckAgain = True
    End If
    If Not bUseCPU Then DoEvents
Next x
If bCheckAgain Then GoTo again:

If chkNoColors.Value = 0 And chkDontMarkStart.Value = 0 Then
   Call DrawOnRoom(lblRoomCell(CenterCell), drSquare, 4, BrightBlue)
End If

frmMain.Enabled = True
Unload frmProgressBar

DoEvents
If cmdHideUnused.Caption = "Show &Blocks" Then
    For x = 1 To 1681
        If CellRoom(x, 1) = 0 Then lblRoomCell(x).Visible = False
    Next x
    DoEvents
End If

'Me.Show
picMap.Visible = True

Exit Sub

Cancel:
frmMain.Enabled = True
Unload frmProgressBar
'Unload Me

Exit Sub
error:
Call HandleError
Unload frmProgressBar
'Me.Show
picMap.Visible = True
DoEvents
End Sub
Public Sub ToggleStopBuild()
StopBuild = True
End Sub
Private Sub MapExits(ByVal Cell As Integer, ByVal Room As Long, ByVal Map As Long)
Dim ActivatedCell As Integer, nStatus As Integer, x As Integer, y As Integer
Dim rc As RECT, ToolTipString As String, sText As String, sMonsters As String
Dim sRemote As String, sCommand As String, nRemote As Integer, sRemoteEffect As String
Dim sExits As String

CellRoom(Cell, 1) = Map
CellRoom(Cell, 2) = Room

Call frmProgressBar.IncreaseProgress

RoomKeyStruct.MapNum = Map
RoomKeyStruct.RoomNum = Room
nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then

    UnchartedCells(Cell) = 2
    Call DrawOnRoom(lblRoomCell(Cell), drSquare, 8, BrightRed)
    If chkNoTooltips.Value = 0 Then
        ToolTipString = "Map " & Map & " Room " & Room
        rc.Left = lblRoomCell(Cell).Left
        rc.Top = lblRoomCell(Cell).Top
        rc.Bottom = (lblRoomCell(Cell).Top + lblRoomCell(Cell).Height)
        rc.Right = (lblRoomCell(Cell).Left + lblRoomCell(Cell).Width)
        objTooltip.SetToolTipItem picMap.hwnd, 0, rc.Left, rc.Top, rc.Right, rc.Bottom, ToolTipString, False
    End If
    Exit Sub
End If
RoomRowToStruct Roomdatabuf.buf

If chkMarkCMD.Value = 1 And Roomrec.CmdText > 0 Then
    sCommand = vbCrLf & vbCrLf & "Room commands: " & GetTextblockCMDS(Roomrec.CmdText)
    Call DrawOnRoom(lblRoomCell(Cell), drSquare, 6, BrightGreen)
End If

If Roomrec.ControlRoom > 0 Then Call AddControlRoom(Roomrec.MapNumber, Roomrec.ControlRoom, Roomrec.RoomNumber)

If chkMarkAux.Value = 1 Then
    If optMarkAux(0).Value = True Then 'room spells
        If Roomrec.Spell > 0 Then Call DrawOnRoom(lblRoomCell(Cell), drStar, 2, BrightCyan)
    ElseIf optMarkAux(1).Value = True Then 'shops
        If Roomrec.ShopNum > 0 Then Call DrawOnRoom(lblRoomCell(Cell), drStar, 2, BrightCyan)
    ElseIf optMarkAux(2).Value = True Then 'exit/death
        If Roomrec.ExitRoom > 0 Then
            Call DrawOnRoom(lblRoomCell(Cell), drStar, 2, BrightCyan)
        ElseIf Roomrec.DeathRoom > 0 Then
            Call DrawOnRoom(lblRoomCell(Cell), drStar, 2, BrightCyan)
        End If
    ElseIf optMarkAux(3).Value = True Then 'control room
        If ControlRoomList.Exists(Roomrec.MapNumber & "/" & Roomrec.RoomNumber) Then
            Call DrawOnRoom(lblRoomCell(Cell), drStar, 2, BrightCyan)
        End If
    End If
End If

'map exits
For x = 0 To 9
    If Not Roomrec.RoomExit(x) = 0 Then
        Select Case Roomrec.RoomType(x)
            Case 8: 'map change
                sExits = sExits & IIf(sExits = "", "", vbCrLf) _
                    & GetRoomExits(x, False) & ":" & "Map Change to map " & Roomrec.Para1(x)
                ActivatedCell = ActivateCell(Cell, x, Roomrec.RoomType(x))
                If ActivatedCell = -1 Then GoTo Skip:
                If chkFollowMapChanges.Value = 1 Then
                    CellRoom(ActivatedCell, 1) = Roomrec.Para1(x)
                    CellRoom(ActivatedCell, 2) = Roomrec.RoomExit(x)
                    If UnchartedCells(ActivatedCell) = 0 Then UnchartedCells(ActivatedCell) = 1
                End If
                
            Case 6: 'hidden
                sExits = sExits & IIf(sExits = "", "", vbCrLf) _
                    & GetRoomExits(x, False) & ": "
                Select Case Roomrec.Para1(x)
                    Case 1: 'passable
                        sExits = sExits & " (Hidden/Passable)"
                    Case 2: 'searchable
                        sExits = sExits & " (Hidden/Searchable)"
                    Case Else:
                        If Abs(Roomrec.Para2(x)) > 1 Then 'does it require more than 1 action?
                            If Roomrec.Para2(x) < 0 Then
                                sExits = sExits & " (Hidden/Needs " & Abs(Roomrec.Para2(x)) & " Actions, any order)"
                            Else
                                sExits = sExits & " (Hidden/Needs " & Roomrec.Para2(x) & " Actions, specific order)"
                            End If
                        Else
                            sExits = sExits & " (Hidden/Needs Action)"
                        End If
                End Select
                ActivatedCell = ActivateCell(Cell, x, Roomrec.RoomType(x))
                If ActivatedCell = -1 Then GoTo Skip:
                If chkDontFollowHidden.Value = 0 Then
                    CellRoom(ActivatedCell, 1) = Map
                    CellRoom(ActivatedCell, 2) = Roomrec.RoomExit(x)
                    If UnchartedCells(ActivatedCell) = 0 Then UnchartedCells(ActivatedCell) = 1
                End If
            
            Case 10: 'text
                sExits = sExits & IIf(sExits = "", "", vbCrLf) _
                    & GetRoomExits(x, False) & ": " & GetMessages(Roomrec.Para1(x), -1)
                
                ActivatedCell = ActivateCell(Cell, x, Roomrec.RoomType(x))
                If ActivatedCell = -1 Then GoTo Skip:
                CellRoom(ActivatedCell, 1) = Map
                CellRoom(ActivatedCell, 2) = Roomrec.RoomExit(x)
                If UnchartedCells(ActivatedCell) = 0 Then UnchartedCells(ActivatedCell) = 1
                
            Case 12: 'remote action
                sExits = sExits & IIf(sExits = "", "", vbCrLf) _
                    & GetRoomExits(x, False) & ": Remote Action"
                If chkMarkCMD.Value = 1 Then
                    If sRemote = "" Then
                        sRemote = vbCrLf & vbCrLf & "Remote Actions: " & vbCrLf
                    Else
                        sRemote = sRemote & vbCrLf
                    End If
                    
                    nRemote = Roomrec.Para2(x)
                    If nRemote > 9 Then
                        Select Case nRemote
                            Case Is > 89:
                                sRemoteEffect = ", action #9"
                            Case Is > 79:
                                sRemoteEffect = ", action #8"
                            Case Is > 69:
                                sRemoteEffect = ", action #7"
                            Case Is > 59:
                                sRemoteEffect = ", action #6"
                            Case Is > 49:
                                sRemoteEffect = ", action #5"
                            Case Is > 39:
                                sRemoteEffect = ", action #4"
                            Case Is > 29:
                                sRemoteEffect = ", action #3"
                            Case Is > 19:
                                sRemoteEffect = ", action #2"
                            Case Is > 9:
                                sRemoteEffect = ", action #1"
                        End Select
                        sRemoteEffect = GetRoomExits(nRemote Mod 10, True) & sRemoteEffect
                    Else
                        sRemoteEffect = GetRoomExits(nRemote, True)
                    End If
                    
                    
                    If Roomrec.RoomNumber = Roomrec.RoomExit(x) Then
                        sRemote = sRemote & "[on this room: "
                    Else
                        sRemote = sRemote & "[on room " & Roomrec.MapNumber & "/" & Roomrec.RoomExit(x) & ": "
                    End If
                    'sRemote = sRemote & "[on room " & Roomrec.RoomExit(x) & ", " & sRemoteEffect & ": "
                    sRemote = sRemote & GetMessages(Roomrec.Para1(x), -1) & "]"
                    If Not Roomrec.Para4(x) = 0 Then sRemote = sRemote & " (Item: " & Roomrec.Para4(x) & ")"
                    
                    Call DrawOnRoom(lblRoomCell(Cell), drSquare, 6, BrightGreen)
                End If
            
            Case Else:
                If Roomrec.RoomType(x) > 0 Then
                    sExits = sExits & IIf(sExits = "", "", vbCrLf) _
                        & GetRoomExits(x, False) & ": " & GetRoomExitType(Roomrec.RoomType(x))
                End If
                ActivatedCell = ActivateCell(Cell, x, Roomrec.RoomType(x))
                If ActivatedCell = -1 Then GoTo Skip:
                CellRoom(ActivatedCell, 1) = Map
                CellRoom(ActivatedCell, 2) = Roomrec.RoomExit(x)
                If UnchartedCells(ActivatedCell) = 0 Then UnchartedCells(ActivatedCell) = 1

        End Select
    End If
Skip:
Next x

If chkMarkLair.Value = 1 And Roomrec.Type = 3 Then
    Call DrawOnRoom(lblRoomCell(Cell), drCircle, 5, BrightMagenta)
End If
    
If chkMarkNPC.Value = 1 And Roomrec.PermNPC > 0 Then
    Call DrawOnRoom(lblRoomCell(Cell), drOpenCircle, 2, BrightRed)
End If

If chkNoColors.Value = 1 Then
    lblRoomCell(Cell).BackColor = nBackcolor '-- nothing
Else
    'set color of this room
    If Roomrec.RoomExit(8) = 0 And Roomrec.RoomExit(9) = 0 Then
        lblRoomCell(Cell).BackColor = nBackcolor '&H0& '-- nothing
    ElseIf Not Roomrec.RoomExit(8) = 0 And Roomrec.RoomExit(9) = 0 Then
        lblRoomCell(Cell).BackColor = &HFF00& '-- up
    ElseIf Roomrec.RoomExit(8) = 0 And Not Roomrec.RoomExit(9) = 0 Then
        lblRoomCell(Cell).BackColor = &HFFFF& '-- down
    Else
        lblRoomCell(Cell).BackColor = &HFFFF00 '-- both
    End If
End If

'& " - " & Mid(Roomrec.Name, 1, InStr(1, Roomrec.Name, Chr(0)) - 1) & vbCrLf & vbCrLf _

If chkNoTooltips.Value = 0 Then
    ToolTipString = Map & "/" & Room & " - " & ClipNull(Roomrec.Name) & vbCrLf
    
    If Roomrec.PermNPC > 0 Then ToolTipString = ToolTipString & "NPC: " & GetMonsterName(Roomrec.PermNPC) & vbCrLf
    
    If Not sExits = "" Then ToolTipString = ToolTipString & sExits & vbCrLf
    
    ToolTipString = ToolTipString & vbCrLf & "Type: " & GetRoomType(Roomrec.Type) & ", " & GetMonType(Roomrec.MonsterType) & vbCrLf _
        & "Index: " & Roomrec.MinIndex & " to " & Roomrec.MaxIndex & vbCrLf _
        & "Control Room: " & Roomrec.ControlRoom & vbCrLf _
        & "Max Regen: " & Roomrec.MaxRegen & ", Delay: " & Roomrec.Delay
    
    
    ToolTipString = ToolTipString & sText & sRemote & sCommand
    
    If chkMonsterRegen.Value = 1 Then
        If Not Roomrec.MaxIndex = 0 And Roomrec.Type = 3 Then '3=lair
            If Not UBound(MGIL(), 1) = 1 Then
                
                sMonsters = ""
                If UBound(MGIL(), 2) < Roomrec.MinIndex Then ReDim Preserve MGIL(UBound(MGIL(), 1), Roomrec.MinIndex)
                If UBound(MGIL(), 2) < Roomrec.MaxIndex Then ReDim Preserve MGIL(UBound(MGIL(), 1), Roomrec.MaxIndex)
                
                For x = Roomrec.MinIndex To Roomrec.MaxIndex
                    For y = 0 To 10 '20
                        If Not MGIL(Roomrec.MonsterType, x).nNumber(y) = 0 Then
                            If sMonsters = "" Then
                                sMonsters = "Also here: "
                            Else
                                sMonsters = sMonsters & ", "
                            End If
                            
                            sMonsters = sMonsters & GetMonsterName(MGIL(Roomrec.MonsterType, x).nNumber(y)) _
                                & "(" & MGIL(Roomrec.MonsterType, x).nNumber(y) & ")"
                        End If
                    Next y
                Next x
                
                'If Len(sMonsters) > 350 Then sMonsters = Left(sMonsters, 325) & " +More"
                
                If Not sMonsters = "" Then ToolTipString = ToolTipString & vbCrLf & vbCrLf & sMonsters
            Else
                ToolTipString = ToolTipString & vbCrLf & vbCrLf & "Also Here: <create monster index to see (Control+D)>"
            End If
        End If
    End If
    
    If ControlRoomList.Exists(Roomrec.MapNumber & "/" & Roomrec.RoomNumber) Then
        ToolTipString = ToolTipString & vbCrLf & vbCrLf & "Control Room From: " & GetControlRoomListByRoom(Roomrec.MapNumber, Roomrec.RoomNumber)
    End If
    
    rc.Left = lblRoomCell(Cell).Left
    rc.Top = lblRoomCell(Cell).Top
    rc.Bottom = (lblRoomCell(Cell).Top + lblRoomCell(Cell).Height)
    rc.Right = (lblRoomCell(Cell).Left + lblRoomCell(Cell).Width)
    objTooltip.SetToolTipItem picMap.hwnd, 0, rc.Left, rc.Top, rc.Right, rc.Bottom, ToolTipString, False
End If

'Call frmProgressBar.IncreaseProgress
'frmProgressBar.lblPanel(1).Caption = CellLabel(x)
frmProgressBar.lblPanel(1).Caption = Map & " / " & Room
        
UnchartedCells(Cell) = 2
End Sub
Private Function ActivateCell(FromCell As Integer, direction As Integer, ExitType As Integer) As Integer
Dim temp As Integer, LineColor As Long
'Dim x1 As Long, x2 As Long, y1 As Long, y2 As Long

'0 = N = -41
'1 = S = +41
'2 = E = +1
'3 = W = -1
'4 = NE = -40
'5 = NW = -42
'6 = SE = +42
'7 = SW = +40

'figure out which cell is to be activated
Select Case direction
    Case 0: 'north
        ActivateCell = (FromCell - 41)
        'checking to see if it's on the north edge
        If ActivateCell < 1 Then
            Call DrawOnRoom(lblRoomCell(FromCell), drLineN, 4, Grey)
            GoTo DontActivate
        End If

    Case 1: 'south
        ActivateCell = (FromCell + 41)
        'checking to see if it's on the south edge
        If ActivateCell > SECorner Then
            Call DrawOnRoom(lblRoomCell(FromCell), drLineS, 4, Grey)
            GoTo DontActivate
        End If

    Case 2: 'east
        ActivateCell = (FromCell + 1)
        'checking to see if it's on the east edge
        For temp = RowLength To SECorner Step 41
            If FromCell = temp Then
                Call DrawOnRoom(lblRoomCell(FromCell), drLineE, 4, Grey)
                GoTo DontActivate
            End If
        Next
        
    Case 3: 'west
        ActivateCell = (FromCell - 1)
        'checking to see if it's on the west edge
        For temp = 1 To SECorner Step 41
            If FromCell = temp Then
                Call DrawOnRoom(lblRoomCell(FromCell), drLineW, 4, Grey)
                GoTo DontActivate
            End If
        Next

    Case 4: 'northeast
        ActivateCell = (FromCell - 40)
        'checking to see if it's on the north edge
        If ActivateCell < 1 Then
            Call DrawOnRoom(lblRoomCell(FromCell), drLineNE, 4, Grey)
            GoTo DontActivate
        End If
        'checking to see if it's on the east edge
        For temp = RowLength To SECorner Step 41
            If FromCell = temp Then
                Call DrawOnRoom(lblRoomCell(FromCell), drLineNE, 4, Grey)
                GoTo DontActivate
            End If
        Next

    Case 5: 'northwest
        ActivateCell = (FromCell - 42)
        'checking to see if it's on the north edge
        If ActivateCell < 1 Then
            Call DrawOnRoom(lblRoomCell(FromCell), drLineNW, 4, Grey)
            GoTo DontActivate:
        End If
        'checking to see if it's on the west edge
        For temp = 1 To SECorner Step 41
            If FromCell = temp Then
                Call DrawOnRoom(lblRoomCell(FromCell), drLineNW, 4, Grey)
                GoTo DontActivate
            End If
        Next

    Case 6: 'southeast
        ActivateCell = (FromCell + 42)
        'checking to see if it's on the south edge
        If ActivateCell > SECorner Then
            Call DrawOnRoom(lblRoomCell(FromCell), drLineSE, 4, Grey)
            GoTo DontActivate
        End If
        'checking to see if it's on the east edge
        For temp = RowLength To SECorner Step 41
            If FromCell = temp Then
                Call DrawOnRoom(lblRoomCell(FromCell), drLineSE, 4, Grey)
                GoTo DontActivate
            End If
        Next

    Case 7: 'southwest
        ActivateCell = (FromCell + 40)
        'checking to see if it's on the south edge
        If ActivateCell > SECorner Then
            Call DrawOnRoom(lblRoomCell(FromCell), drLineSW, 4, Grey)
            GoTo DontActivate:
        End If
        'checking to see if it's on the west edge
        For temp = 1 To SECorner Step 41
            If FromCell = temp Then
                Call DrawOnRoom(lblRoomCell(FromCell), drLineSW, 4, Grey)
                GoTo DontActivate
            End If
        Next

    Case 8:
        'If chkNoColors.value = 1 Then GoTo DontActivate:
'        Select Case lblRoomCell(FromCell).Tag
'            Case 1: 'up
'            Case 2: 'down
'                Call DrawOnRoom(lblRoomCell(FromCell), drCircle, BrightCyan)
'                lblRoomCell(FromCell).Tag = 3
'            Case 3: 'up and down
'            Case Else:
'                Call DrawOnRoom(lblRoomCell(FromCell), drCircle, BrightGreen)
'                lblRoomCell(FromCell).Tag = 1
'        End Select
'        If lblRoomCell(FromCell).BackColor = &H0 Then lblRoomCell(FromCell).BackColor = &HFF00&
'        If lblRoomCell(FromCell).BackColor = &HFFFFFF Then lblRoomCell(FromCell).BackColor = &HFF00&
'        If lblRoomCell(FromCell).BackColor = &HFFFF& Then lblRoomCell(FromCell).BackColor = &HFFFF00
        GoTo DontActivate:

    Case 9:
        'If chkNoColors.value = 1 Then GoTo DontActivate:
'        Select Case lblRoomCell(FromCell).Tag
'            Case 1: 'up
'                Call DrawOnRoom(lblRoomCell(FromCell), drCircle, BrightCyan)
'                lblRoomCell(FromCell).Tag = 3
'            Case 2: 'down
'            Case 3: 'up and down
'            Case Else:
'                Call DrawOnRoom(lblRoomCell(FromCell), drCircle, Yellow)
'                lblRoomCell(FromCell).Tag = 2
'        End Select
        
'        If lblRoomCell(FromCell).BackColor = &H0 Then lblRoomCell(FromCell).BackColor = &HFFFF&
'        If lblRoomCell(FromCell).BackColor = &HFFFFFF Then lblRoomCell(FromCell).BackColor = &HFFFF&
'        If lblRoomCell(FromCell).BackColor = &HFF00& Then lblRoomCell(FromCell).BackColor = &HFFFF00
        GoTo DontActivate:
    
    Case Else:
        GoTo DontActivate:
        
End Select

If ActivateCell < 1 Or ActivateCell > SECorner Then GoTo DontActivate:

'set line mode
'ScaleMode = vbPixels
DrawWidth = 4

'pick line color
Select Case ExitType
    Case 2: LineColor = 10    'l green - key
    Case 3: LineColor = 10    'l green - item
    Case 4: LineColor = 10    'l green - toll
    Case 5: LineColor = 11    'l cyan - action
    Case 6: LineColor = 5     'd magenta - hidden
    Case 7: LineColor = 9     'l blue - door/gate
    Case 8: LineColor = 13    'l magenta - map change
    Case 9: LineColor = 12    'l red - trap/spell trap
    Case 10: LineColor = 14   'l yellow - text
    Case 11: LineColor = 9    'l blue - door/gate
    Case 12: LineColor = 11   'l cyan - remote action
    Case 13: LineColor = 1    'd blue - class
    Case 14: LineColor = 1    'd blue - race
    Case 15: LineColor = 1    'd blue - level
    Case 16: LineColor = 2    'green (was gray) - timed
    Case 20: LineColor = 1    'd blue - alignment
    Case 23: LineColor = 1    'd blue - ability
    Case 24: LineColor = 12   'l red - trap/spell trap
    Case Else:
        If chkUseWhiteBG.Value = 1 Then
            LineColor = 0 'black - anything else
        Else
            LineColor = 8 'l grey - anything else
        End If
End Select

If chkNoColors.Value = 1 Or chkNoLineColors.Value = 1 Then
    If chkUseWhiteBG.Value = 1 Then
        LineColor = 0 'black
    Else
        LineColor = 8 'd grey
    End If
End If

'draw the line
Select Case direction
    Case 0:
'        x1 = lblRoomCell(FromCell).Left + 4
'        y1 = lblRoomCell(FromCell).Top + 4
'        x2 = lblRoomCell(ActivateCell).Left + 4
'        y2 = Abs((lblRoomCell(ActivateCell).Top + 4 + lblRoomCell(FromCell).Top + 4) / 2)
'        Line (x1, y1)-(x2, y2), QBColor(LineColor), BF
        Call DrawOnRoom(lblRoomCell(FromCell), drLineN, 4, LineColor)
    Case 1:
'        x1 = lblRoomCell(FromCell).Left + 4
'        y1 = lblRoomCell(FromCell).Top + 4
'        x2 = lblRoomCell(ActivateCell).Left + 4
'        y2 = Abs(((lblRoomCell(ActivateCell).Top + 4 + lblRoomCell(FromCell).Top + 4) / 2)) + 1
'        Line (x1, y1)-(x2, y2), QBColor(LineColor), BF
        Call DrawOnRoom(lblRoomCell(FromCell), drLineS, 4, LineColor)
    Case 2:
'        x1 = lblRoomCell(FromCell).Left + 4
'        y1 = lblRoomCell(FromCell).Top + 4
'        x2 = Abs(((lblRoomCell(ActivateCell).Left + 4 + lblRoomCell(FromCell).Left + 4) / 2))
'        y2 = lblRoomCell(ActivateCell).Top + 4
'        Line (x1, y1)-(x2, y2), QBColor(LineColor), BF
        Call DrawOnRoom(lblRoomCell(FromCell), drLineE, 4, LineColor)
    Case 3:
'        x1 = lblRoomCell(FromCell).Left + 4
'        y1 = lblRoomCell(FromCell).Top + 4
'        x2 = Abs(((lblRoomCell(ActivateCell).Left + 4 + lblRoomCell(FromCell).Left + 4) / 2))
'        y2 = lblRoomCell(ActivateCell).Top + 4
'        Line (x1, y1)-(x2, y2), QBColor(LineColor), BF
        Call DrawOnRoom(lblRoomCell(FromCell), drLineW, 4, LineColor)
    Case 4:
'        x1 = lblRoomCell(FromCell).Left + 4
'        y1 = lblRoomCell(FromCell).Top + 4
'        x2 = Abs(((lblRoomCell(ActivateCell).Left + 4 + lblRoomCell(FromCell).Left + 4) / 2))
'        y2 = Abs(((lblRoomCell(ActivateCell).Top + 4 + lblRoomCell(FromCell).Top + 4) / 2))
'        Line (x1, y1)-(x2, y2), QBColor(LineColor)
        Call DrawOnRoom(lblRoomCell(FromCell), drLineNE, 4, LineColor)
    Case 5:
'        x1 = lblRoomCell(FromCell).Left + 4
'        y1 = lblRoomCell(FromCell).Top + 4
'        x2 = Abs(((lblRoomCell(ActivateCell).Left + 4 + lblRoomCell(FromCell).Left + 4) / 2))
'        y2 = Abs(((lblRoomCell(ActivateCell).Top + 4 + lblRoomCell(FromCell).Top + 4) / 2))
'        Line (x1, y1)-(x2, y2), QBColor(LineColor)
        Call DrawOnRoom(lblRoomCell(FromCell), drLineNW, 4, LineColor)
    Case 6:
'        x1 = lblRoomCell(FromCell).Left + 4
'        y1 = lblRoomCell(FromCell).Top + 4
'        x2 = Abs(((lblRoomCell(ActivateCell).Left + 4 + lblRoomCell(FromCell).Left + 4) / 2))
'        y2 = Abs(((lblRoomCell(ActivateCell).Top + 4 + lblRoomCell(FromCell).Top + 4) / 2))
'        Line (x1, y1)-(x2, y2), QBColor(LineColor)
        Call DrawOnRoom(lblRoomCell(FromCell), drLineSE, 4, LineColor)
    Case 7:
'        x1 = lblRoomCell(FromCell).Left + 4
'        y1 = lblRoomCell(FromCell).Top + 4
'        x2 = Abs(((lblRoomCell(ActivateCell).Left + 4 + lblRoomCell(FromCell).Left + 4) / 2))
'        y2 = Abs(((lblRoomCell(ActivateCell).Top + 4 + lblRoomCell(FromCell).Top + 4) / 2))
'        Line (x1, y1)-(x2, y2), QBColor(LineColor)
        Call DrawOnRoom(lblRoomCell(FromCell), drLineSW, 4, LineColor)
        
End Select

'if the cell to be activated has already been mapped, dont map it again
If UnchartedCells(ActivateCell) = 2 Then GoTo DontActivate:

Select Case ExitType
    Case 12: ActivateCell = -1 'if it's a remote action, dont map it
    Case 8: 'if it's a map change, check to see if it should be mapped
        If chkFollowMapChanges.Value = 1 Then
            lblRoomCell(ActivateCell).BackColor = &H0
        Else
            ActivateCell = -1
        End If
    Case Else: lblRoomCell(ActivateCell).BackColor = &H0
End Select

Exit Function
DontActivate:
ActivateCell = -1

End Function

Private Sub lblRoomCell_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo error:

If CellRoom(Index, 1) = 0 Then
    If Button = 2 And Shift = 1 Then
        CenterCell = Index
        Call StartMapping
        Me.SetFocus
        Exit Sub
    Else
        Exit Sub
    End If
End If

nLastMapClick = CellRoom(Index, 1)
nLastRoomClick = CellRoom(Index, 2)
    
If Button = 1 Then
    frmRoom.Show
    
    Call frmRoom.GotoRoom(CellRoom(Index, 1), CellRoom(Index, 2), True)
    If FormIsLoaded("frmMapEditor") Then frmMapEditor.SetFocus Else frmRoom.SetFocus
    Exit Sub

ElseIf Button = 2 Then
    If Shift = 1 Then CenterCell = Index
    
    If lblRoomCell(Index).BackColor = &HFF00& Then '-- up
        PopupMenu frmMain.mnuMapUp
    ElseIf lblRoomCell(Index).BackColor = &HFFFF& Then '-- down
        PopupMenu frmMain.mnuMapDown
    ElseIf lblRoomCell(Index).BackColor = &HFFFF00 Then '-- both
        PopupMenu frmMain.mnuMapUpDown
    Else
        frmRoom.Show
        Call frmRoom.GotoRoom(CellRoom(Index, 1), CellRoom(Index, 2), True)
        
        Call StartMapping
        Me.SetFocus
    End If
End If

out:
Exit Sub
error:
Call HandleError("lblRoomCell_MouseDown")
Resume out:

End Sub


Private Sub Form_Unload(Cancel As Integer)
Set objTooltip = Nothing

If FormIsLoaded("frmMapLegend") Then Unload frmMapLegend
Call WriteINI("Options", "MapSize", cmbMapSize.ListIndex)
Call WriteINI("Options", "MapNoColors", chkNoColors.Value)
Call WriteINI("Options", "MapNoLineColors", chkNoLineColors.Value)
Call WriteINI("Options", "MapDontMarkStart", chkDontMarkStart.Value)
Call WriteINI("Options", "MapFollowMapChanges", chkFollowMapChanges.Value)
Call WriteINI("Options", "MapDontFollowHidden", chkDontFollowHidden.Value)
Call WriteINI("Options", "MapMarkLair", chkMarkLair.Value)
Call WriteINI("Options", "MapMarkCMD", chkMarkCMD.Value)
Call WriteINI("Options", "MapMarkNPC", chkMarkNPC.Value)
Call WriteINI("Options", "MapUseSameExport", chkUseLastExport.Value)
Call WriteINI("Options", "MapLookUpMonsterRegen", chkMonsterRegen.Value)
Call WriteINI("Options", "MapWhiteBG", chkUseWhiteBG.Value)
Call WriteINI("Options", "MapNoTooltips", chkNoTooltips.Value)

Call WriteINI("Options", "MapMarkAux", chkMarkAux.Value)
If optMarkAux(0).Value = True Then
    Call WriteINI("Options", "MapMarkAuxSpells", 1)
    Call WriteINI("Options", "MapMarkAuxShops", 0)
    Call WriteINI("Options", "MapMarkAuxRooms", 0)
    Call WriteINI("Options", "MapMarkAuxControl", 0)
ElseIf optMarkAux(1).Value = True Then
    Call WriteINI("Options", "MapMarkAuxSpells", 0)
    Call WriteINI("Options", "MapMarkAuxShops", 1)
    Call WriteINI("Options", "MapMarkAuxRooms", 0)
    Call WriteINI("Options", "MapMarkAuxControl", 0)
ElseIf optMarkAux(2).Value = True Then
    Call WriteINI("Options", "MapMarkAuxSpells", 0)
    Call WriteINI("Options", "MapMarkAuxShops", 0)
    Call WriteINI("Options", "MapMarkAuxRooms", 1)
    Call WriteINI("Options", "MapMarkAuxControl", 0)
ElseIf optMarkAux(3).Value = True Then
    Call WriteINI("Options", "MapMarkAuxSpells", 0)
    Call WriteINI("Options", "MapMarkAuxShops", 0)
    Call WriteINI("Options", "MapMarkAuxRooms", 0)
    Call WriteINI("Options", "MapMarkAuxControl", 1)
End If

If Not Me.WindowState = vbMinimized Then
    Call WriteINI("Windows", "MapTop", Me.Top)
    Call WriteINI("Windows", "MapLeft", Me.Left)
End If

End Sub

Private Sub DrawOnRoom(ByRef oLabel As label, ByVal drDrawType As DrawRoomEnum, ByVal nSize As Integer, ByVal nColor As QBColorCode)
Dim x1 As Integer, x2 As Integer, y1 As Integer, y2 As Integer
Dim nTemp As Integer, nLineColor As QBColorCode

nTemp = picMap.DrawWidth

If chkNoColors.Value = 1 Then
    If chkUseWhiteBG.Value = 1 Then
        nColor = black
    Else
        nColor = Grey
    End If
End If

nLineColor = nColor
If chkNoLineColors.Value = 1 Then
    If chkUseWhiteBG.Value = 1 Then
        nLineColor = black
    Else
        nLineColor = Grey
    End If
End If

Select Case drDrawType
    Case 0: 'square
        picMap.DrawWidth = nSize
        x1 = oLabel.Left
        y1 = oLabel.Top
        x2 = oLabel.Left + oLabel.Width
        y2 = oLabel.Top + oLabel.Height
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor), BF
        
    Case 1: 'star
        picMap.DrawWidth = nSize
        '/
        x1 = oLabel.Left - 2
        y1 = oLabel.Top + oLabel.Height + 2
        x2 = oLabel.Left + 4
        y2 = oLabel.Top - 4
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
        
        '\
        x2 = oLabel.Left + oLabel.Width + 2
        y2 = oLabel.Top + oLabel.Height + 2
        picMap.Line -(x2, y2), QBColor(nColor) '(x1, y1)-(x2, y2), QBColor(nColor)
        
        '\
        x2 = oLabel.Left - 4
        y2 = oLabel.Top
        picMap.Line -(x2, y2), QBColor(nColor)
        
        '-
        x2 = oLabel.Left + oLabel.Width + 3
        y2 = oLabel.Top
        picMap.Line -(x2, y2), QBColor(nColor)
        
        '/
        x2 = oLabel.Left - 2
        y2 = oLabel.Top + oLabel.Height + 2
        picMap.Line -(x2, y2), QBColor(nColor)
        
    Case 2: 'open circle
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        picMap.Circle (x1, y1), 8, QBColor(nColor)
      
     Case 3: 'up
        picMap.DrawWidth = nSize
        x1 = oLabel.Left
        y1 = oLabel.Top
        x2 = oLabel.Left + oLabel.Width
        y2 = oLabel.Top + 2
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor), B
        
     Case 4: 'down
        picMap.DrawWidth = nSize
        x1 = oLabel.Left - 1
        y1 = oLabel.Top + oLabel.Height - 1
        x2 = oLabel.Left + oLabel.Width
        y2 = y1 + 2
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor), B
    
    Case 5: 'circle
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        picMap.Circle (x1, y1), 5, QBColor(nColor)
    
    Case 6: 'LineN
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        x2 = x1
        y2 = y1 - 8
        picMap.Line (x1, y1)-(x2, y2), QBColor(nLineColor), BF
        
    Case 7: 'LineS
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        x2 = x1
        y2 = y1 + 8
        picMap.Line (x1, y1)-(x2, y2), QBColor(nLineColor), BF
        
    Case 8: 'LineE
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        x2 = x1 + 8
        y2 = y1
        picMap.Line (x1, y1)-(x2, y2), QBColor(nLineColor), BF
        
    Case 9: 'LineW
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        x2 = x1 - 8
        y2 = y1
        picMap.Line (x1, y1)-(x2, y2), QBColor(nLineColor), BF
        
    Case 10: 'LineNE
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        x2 = x1 + 8
        y2 = y1 - 8
        picMap.Line (x1, y1)-(x2, y2), QBColor(nLineColor)
        
    Case 11: 'LineNW
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        x2 = x1 - 8
        y2 = y1 - 8
        picMap.Line (x1, y1)-(x2, y2), QBColor(nLineColor)
        
    Case 12: 'LineSE
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        x2 = x1 + 8
        y2 = y1 + 8
        picMap.Line (x1, y1)-(x2, y2), QBColor(nLineColor)
    
    Case 13: 'LineSW
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        x2 = x1 - 8
        y2 = y1 + 8
        picMap.Line (x1, y1)-(x2, y2), QBColor(nLineColor)
        
End Select

picMap.DrawWidth = nTemp
End Sub


Private Sub optMarkAux_Click(Index As Integer)
    If optMarkAux(3).Value = False Or optMarkAux(3).Enabled = False Then
        cmdBuildControlRoomList.Enabled = False
    Else
        cmdBuildControlRoomList.Enabled = True
    End If
End Sub
