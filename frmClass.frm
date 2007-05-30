VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20D5284F-7B23-4F0A-B8B1-6C9D18B64F1C}#1.0#0"; "exlimiter.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmClass 
   Caption         =   "Class Editor"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6330
   Icon            =   "frmClass.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5205
   ScaleWidth      =   6330
   Begin MSComctlLib.ListView lvDatabase 
      Height          =   4695
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   8281
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
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   180
      Width           =   2535
   End
   Begin VB.Frame framNav 
      BorderStyle     =   0  'None
      Height          =   5115
      Left            =   2580
      TabIndex        =   3
      Top             =   60
      Width           =   3675
      Begin VB.CommandButton cmdDiscard 
         Caption         =   "Dis&card"
         Height          =   315
         Left            =   2820
         TabIndex        =   7
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   315
         Left            =   660
         TabIndex        =   5
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "&Insert"
         Height          =   315
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   675
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   315
         Left            =   1980
         TabIndex        =   6
         Top             =   0
         Width           =   855
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4575
         Left            =   60
         TabIndex        =   8
         Top             =   540
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   8070
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "frmClass.frx":08CA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label(15)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label(14)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label(13)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label(12)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label(11)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label(10)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label(9)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label(8)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Line1"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label(1)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label(0)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "cmdEditTitleText"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "txtNumber"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "txtName"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "txtMagicType"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "txtWeapon"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "txtArmour"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "txtTitleText"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "txtMagicLvL"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "txtMaxHP"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "txtCombat"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "txtMinHP"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "txtExp"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "chkAutoSave"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).ControlCount=   24
         TabCaption(1)   =   "Abilities"
         TabPicture(1)   =   "frmClass.frx":08E6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frmAbilities"
         Tab(1).ControlCount=   1
         Begin VB.CheckBox chkAutoSave 
            Caption         =   "Auto-Save"
            Height          =   195
            Left            =   2400
            TabIndex        =   66
            Top             =   420
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.TextBox txtExp 
            Height          =   315
            Left            =   960
            TabIndex        =   14
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txtMinHP 
            Height          =   315
            Left            =   960
            TabIndex        =   16
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox txtCombat 
            Height          =   315
            Left            =   960
            TabIndex        =   20
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox txtMaxHP 
            Height          =   315
            Left            =   2100
            TabIndex        =   18
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox txtMagicLvL 
            Height          =   315
            Left            =   2040
            TabIndex        =   27
            Top             =   3240
            Width           =   1095
         End
         Begin VB.TextBox txtTitleText 
            Height          =   315
            Left            =   960
            TabIndex        =   29
            Top             =   3600
            Width           =   975
         End
         Begin VB.ComboBox txtArmour 
            Height          =   315
            ItemData        =   "frmClass.frx":0902
            Left            =   960
            List            =   "frmClass.frx":0924
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   2520
            Width           =   2175
         End
         Begin VB.ComboBox txtWeapon 
            Height          =   315
            ItemData        =   "frmClass.frx":09AB
            Left            =   960
            List            =   "frmClass.frx":09CD
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   2880
            Width           =   2175
         End
         Begin VB.ComboBox txtMagicType 
            Height          =   315
            ItemData        =   "frmClass.frx":0A3D
            Left            =   960
            List            =   "frmClass.frx":0A53
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   3240
            Width           =   975
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   960
            MaxLength       =   29
            TabIndex        =   12
            Top             =   900
            Width           =   2535
         End
         Begin VB.TextBox txtNumber 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   540
            Width           =   615
         End
         Begin VB.Frame frmAbilities 
            Caption         =   "Abilities"
            Height          =   4095
            Left            =   -74880
            TabIndex        =   31
            Top             =   360
            Width           =   3375
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
               Height          =   195
               Left            =   2640
               TabIndex        =   32
               Top             =   0
               Width           =   615
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   0
               Left            =   120
               TabIndex        =   36
               Top             =   480
               Width           =   495
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   4
               Left            =   2640
               TabIndex        =   50
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   1920
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   3
               Left            =   2640
               TabIndex        =   47
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   1560
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   2
               Left            =   2640
               TabIndex        =   44
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   1
               Left            =   2640
               TabIndex        =   41
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   840
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   0
               Left            =   2640
               TabIndex        =   38
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   9
               Left            =   2640
               TabIndex        =   65
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   3720
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   8
               Left            =   2640
               TabIndex        =   62
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   3360
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   7
               Left            =   2640
               TabIndex        =   59
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   3000
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   6
               Left            =   2640
               TabIndex        =   56
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   2640
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   5
               Left            =   2640
               TabIndex        =   53
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   2280
               Width           =   615
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   1
               Left            =   120
               TabIndex        =   39
               Top             =   840
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   2
               Left            =   120
               TabIndex        =   42
               Top             =   1200
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   3
               Left            =   120
               TabIndex        =   45
               Top             =   1560
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   4
               Left            =   120
               TabIndex        =   48
               Top             =   1920
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   5
               Left            =   120
               TabIndex        =   51
               Top             =   2280
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   6
               Left            =   120
               TabIndex        =   54
               Top             =   2640
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   7
               Left            =   120
               TabIndex        =   57
               Top             =   3000
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   8
               Left            =   120
               TabIndex        =   60
               Top             =   3360
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   9
               Left            =   120
               TabIndex        =   63
               Top             =   3720
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
               TabIndex        =   37
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
               TabIndex        =   40
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
               TabIndex        =   43
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
               TabIndex        =   46
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
               TabIndex        =   49
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
               TabIndex        =   52
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
               TabIndex        =   55
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
               TabIndex        =   58
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
               TabIndex        =   61
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
               TabIndex        =   64
               Text            =   "empty"
               Top             =   3720
               Width           =   1815
            End
            Begin VB.Label Label3 
               Caption         =   "#"
               Height          =   255
               Left            =   120
               TabIndex        =   34
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "Name"
               Height          =   255
               Left            =   720
               TabIndex        =   33
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label2 
               Caption         =   "Value"
               Height          =   255
               Left            =   2640
               TabIndex        =   35
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.CommandButton cmdEditTitleText 
            Height          =   195
            Left            =   2040
            TabIndex        =   30
            Top             =   3660
            Width           =   195
         End
         Begin VB.Label Label 
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
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   540
            Width           =   855
         End
         Begin VB.Label Label 
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
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   900
            Width           =   855
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   3480
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Label Label 
            Caption         =   "Exp %"
            Height          =   315
            Index           =   8
            Left            =   120
            TabIndex        =   13
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label 
            Caption         =   "HP"
            Height          =   315
            Index           =   9
            Left            =   120
            TabIndex        =   15
            Top             =   1800
            Width           =   675
         End
         Begin VB.Label Label 
            Caption         =   "Armour"
            Height          =   315
            Index           =   10
            Left            =   120
            TabIndex        =   21
            Top             =   2520
            Width           =   735
         End
         Begin VB.Label Label 
            Caption         =   "Weapons"
            Height          =   315
            Index           =   11
            Left            =   120
            TabIndex        =   23
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label Label 
            Caption         =   "Combat"
            Height          =   315
            Index           =   12
            Left            =   120
            TabIndex        =   19
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label 
            Caption         =   "Magery"
            Height          =   315
            Index           =   13
            Left            =   120
            TabIndex        =   25
            Top             =   3240
            Width           =   735
         End
         Begin VB.Label Label 
            Caption         =   "Title Text"
            Height          =   315
            Index           =   14
            Left            =   120
            TabIndex        =   28
            Top             =   3600
            Width           =   735
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   15
            Left            =   1920
            TabIndex        =   17
            Top             =   1740
            Width           =   195
         End
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   3660
         Y1              =   420
         Y2              =   420
      End
   End
   Begin exlimiter.EL EL1 
      Left            =   4140
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
   End
   Begin VB.Label Label5 
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim bLoaded As Boolean
Dim nCurrentRecord As Integer

Private Sub cmdAbilsClear_Click()
Dim x As Integer
On Error GoTo Error:

For x = 0 To 9
    txtAbilityA(x).Text = 0
    txtAbilityB(x).Text = 0
Next x

out:
Exit Sub
Error:
Call HandleError("cmdAbilsClear_Click")
Resume out:

End Sub

Private Sub Form_Load()
On Error Resume Next
bLoaded = False

With EL1
    .FormInQuestion = Me
    .MINHEIGHT = 375 + (TITLEBAR_OFFSET / 10)
    .MINWIDTH = 445
    .CenterOnLoad = False
    .EnableLimiter = True
End With

Me.Top = ReadINI("Windows", "ClassTop")
Me.Left = ReadINI("Windows", "ClassLeft")
Me.Width = ReadINI("Windows", "ClassWidth")
Me.Height = ReadINI("Windows", "ClassHeight")

lvDatabase.ListItems.clear

Call LoadClasses

Me.Show
Me.SetFocus
txtSearch.SetFocus
If ReadINI("Windows", "ClassMaxed") = "1" Then Me.WindowState = vbMaximized
End Sub

Private Sub cmdDiscard_Click()
Dim nStatus As Integer

If lvDatabase.SelectedItem Is Nothing Or nCurrentRecord = 0 Then
    MsgBox "No current record."
    Exit Sub
End If

nStatus = BTRCALL(BGETEQUAL, ClassPosBlock, Classdatabuf, Len(Classdatabuf), nCurrentRecord, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error on BGETEQUAL: " & BtrieveErrorCode(nStatus)
Else
    DispClassInfo Classdatabuf.buf
End If

End Sub

Private Sub cmdEditTitleText_Click()
    Call frmTextblock.GotoTB(Val(txtTitleText.Text))
End Sub

Private Sub cmdSave_Click()
On Error GoTo Error:

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
If lvDatabase.SelectedItem Is Nothing Then Exit Sub

Call saverecord(nCurrentRecord)
'Call lvDatabase_ItemClick(lvDatabase.SelectedItem)

Dim oLI As ListItem
Set oLI = lvDatabase.FindItem(Classrec.Number, lvwText, , 0)
If Not oLI Is Nothing Then
    oLI.ListSubItems(1).Text = ClipNull(Classrec.Name)
    oLI.ListSubItems(2).Text = Classrec.Exp + 100 & "%"
    oLI.ListSubItems(3).Text = GetClassWeaponType(Classrec.Weapon)
    oLI.ListSubItems(4).Text = GetArmourType(Classrec.Armour)
    oLI.ListSubItems(5).Text = GetMagery(Classrec.MagicType, Classrec.MagicLvL)
    oLI.ListSubItems(6).Text = Classrec.Combat - 2
    oLI.ListSubItems(7).Text = Classrec.MinHp & "-" & (Classrec.MinHp + Classrec.MaxHP)
End If
Set oLI = Nothing

out:
Exit Sub
Error:
Call HandleError("cmdSave_Click")
Resume out:

End Sub

Private Sub LoadClasses()
On Error GoTo Error:
Dim nStatus As Integer

lvDatabase.ColumnHeaders.clear
lvDatabase.ColumnHeaders.add 1, "Number", "#", 400, lvwColumnLeft
lvDatabase.ColumnHeaders.add 2, "Name", "Name", 1300, lvwColumnCenter
lvDatabase.ColumnHeaders.add 3, "Exp%", "Exp", 700, lvwColumnCenter
lvDatabase.ColumnHeaders.add 4, "Weapon", "Weapon", 1400, lvwColumnCenter
lvDatabase.ColumnHeaders.add 5, "Armour", "Armour", 1000, lvwColumnCenter
lvDatabase.ColumnHeaders.add 6, "Magic", "Magic", 900, lvwColumnCenter
lvDatabase.ColumnHeaders.add 7, "Cmbt", "Cmbt", 600, lvwColumnCenter
lvDatabase.ColumnHeaders.add 8, "HP", "HP", 600, lvwColumnCenter

nStatus = BTRCALL(BGETFIRST, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "LoadClasses, BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

Do While nStatus = 0
    ClassRowToStruct Classdatabuf.buf
    
    Call AddClassToLV(Classrec.Number)
    
    nStatus = BTRCALL(BGETNEXT, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "LoadClasses, Error: " & BtrieveErrorCode(nStatus)
End If

Call modMain.LoadClassArray

If lvDatabase.ListItems.Count >= 1 Then Call lvDatabase_ItemClick(lvDatabase.ListItems(1))

lvDatabase.refresh
SortListView lvDatabase, 1, ldtNumber, True
bLoaded = True

Exit Sub
Error:
Call HandleError
End Sub

Private Sub AddClassToLV(ByVal nNumber As Integer)
Dim nStatus As Integer, oLI As ListItem
On Error GoTo Error:

If Not nNumber = Classrec.Number Then
    nStatus = BTRCALL(BGETEQUAL, ClassPosBlock, Classdatabuf, Len(Classdatabuf), nNumber, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then MsgBox "Error getting record " & nNumber & ": " & BtrieveErrorCode(nStatus)
    bLoaded = False
    Exit Sub
End If

Set oLI = lvDatabase.ListItems.add()
oLI.Text = Classrec.Number

oLI.ListSubItems.add (1), "Name", ClipNull(Classrec.Name)
oLI.ListSubItems.add (2), "Exp%", Classrec.Exp + 100 & "%"
oLI.ListSubItems.add (3), "Weapon", GetClassWeaponType(Classrec.Weapon)
oLI.ListSubItems.add (4), "Armour", GetArmourType(Classrec.Armour)
oLI.ListSubItems.add (5), "Magic", GetMagery(Classrec.MagicType, Classrec.MagicLvL)
oLI.ListSubItems.add (6), "Cmbt", Classrec.Combat - 2
oLI.ListSubItems.add (7), "HP", Classrec.MinHp & "-" & (Classrec.MinHp + Classrec.MaxHP)

Set oLI = Nothing
Exit Sub
Error:
Call HandleError
Set oLI = Nothing
End Sub
Private Sub DispClassInfo(row() As Byte)
On Error GoTo Error:
Dim x As Integer

Call ClassRowToStruct(row())

Me.Caption = "Class Editor -- " & ClipNull(Classrec.Name)

txtNumber.Text = SInt2UInt(Classrec.Number)
txtName.Text = Classrec.Name
txtMinHP.Text = SInt2UInt(Classrec.MinHp)
txtMaxHP.Text = (SInt2UInt(Classrec.MinHp) + SInt2UInt(Classrec.MaxHP))
txtExp.Text = Classrec.Exp + 100
txtMagicLvL.Text = Classrec.MagicLvL
txtCombat.Text = Classrec.Combat - 2
txtTitleText.Text = SLong2ULong(Classrec.TitleText)
txtMagicType.ListIndex = SInt2UInt(Classrec.MagicType)
txtWeapon.ListIndex = SInt2UInt(Classrec.Weapon)
txtArmour.ListIndex = SInt2UInt(Classrec.Armour)
For x = 0 To 9
    txtAbilityA(x).Text = Classrec.AbilityA(x)
    txtAbilityB(x).Text = Classrec.AbilityB(x)
Next x

Exit Sub
Error:
Call HandleError
MsgBox "Warning, record was not completely displayed." & vbCrLf _
    & "Previous records stats may still be in memory.  Select 'Disable DB Writing'" & vbCrLf _
    & "from the file menu and then reload the editor.", vbExclamation
End Sub


Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = vbMinimized Then Exit Sub
framNav.Left = Me.Width - framNav.Width - 220
lvDatabase.Width = framNav.Left - 100
lvDatabase.Height = Me.Height - 925 - TITLEBAR_OFFSET
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
If bLoaded = True Then
    saverecord (nCurrentRecord)
    Call modMain.LoadClassArray
End If

If Me.WindowState = vbMinimized Then Exit Sub

If Me.WindowState = vbMaximized Then
    Call WriteINI("Windows", "ClassMaxed", 1)
Else
    Call WriteINI("Windows", "ClassMaxed", 0)
    Call WriteINI("Windows", "ClassTop", Me.Top)
    Call WriteINI("Windows", "ClassLeft", Me.Left)
    Call WriteINI("Windows", "ClassWidth", Me.Width)
    Call WriteINI("Windows", "ClassHeight", Me.Height)
End If

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

Private Sub saverecord(ByVal nRecord As Integer)
On Error GoTo Error:
Dim nStatus As Integer, x As Integer

If nRecord = 0 Then Exit Sub

nStatus = BTRCALL(BGETEQUAL, ClassPosBlock, Classdatabuf, Len(Classdatabuf), nRecord, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Save Error on BGETEQUAL: " & BtrieveErrorCode(nStatus)
    Exit Sub
Else
    ClassRowToStruct Classdatabuf.buf
End If
'DoEvents
Classrec.Name = txtName.Text & Chr(0)
Classrec.MinHp = UInt2SInt(Val(txtMinHP.Text))
Classrec.MaxHP = UInt2SInt((Val(txtMaxHP.Text) - Val(txtMinHP.Text)))
Classrec.Exp = (Val(txtExp.Text) - 100)
Classrec.MagicLvL = Val(txtMagicLvL.Text)
Classrec.Combat = (Val(txtCombat.Text) + 2)
Classrec.TitleText = ULong2SLong(Val(txtTitleText.Text))
Classrec.MagicType = Val(txtMagicType.ListIndex)
Classrec.Weapon = Val(txtWeapon.ListIndex)
Classrec.Armour = Val(txtArmour.ListIndex)

For x = 0 To 9
    Classrec.AbilityA(x) = Val(txtAbilityA(x).Text)
    Classrec.AbilityB(x) = Val(txtAbilityB(x).Text)
Next x

nStatus = UpdateClass
If Not nStatus = 0 Then
    MsgBox "SaveRecord, Error: " & BtrieveErrorCode(nStatus)
Else
    DispClassInfo Classdatabuf.buf
End If

Exit Sub
Error:
Call HandleError
End Sub

Private Sub lvDatabase_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim nSort As ListDataType
Select Case ColumnHeader.Index
    Case 1, 3, 7, 8: nSort = ldtNumber
    Case Else: nSort = ldtString
End Select
SortListView lvDatabase, ColumnHeader.Index, nSort, lvDatabase.SortOrder
End Sub

Private Sub lvDatabase_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim temp As Integer, nStatus As Integer

If bLoaded = True And chkAutoSave.Value = 1 Then Call saverecord(nCurrentRecord)

temp = Val(Item.Text)
nStatus = BTRCALL(BGETEQUAL, ClassPosBlock, Classdatabuf, Len(Classdatabuf), temp, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error on BGETEQUAL: " & BtrieveErrorCode(nStatus)
    bLoaded = False
Else
    nCurrentRecord = temp
    DispClassInfo Classdatabuf.buf
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

Private Sub txtAbilityB_GotFocus(Index As Integer)
Call SelectAll(txtAbilityB(Index))

End Sub

Private Sub txtCombat_GotFocus()
Call SelectAll(txtCombat)

End Sub

Private Sub txtExp_GotFocus()
Call SelectAll(txtExp)

End Sub

Private Sub txtMagicLvL_GotFocus()
Call SelectAll(txtMagicLvL)

End Sub

Private Sub txtMaxHP_GotFocus()
Call SelectAll(txtMaxHP)

End Sub

Private Sub txtMinHP_GotFocus()
Call SelectAll(txtMinHP)

End Sub

Private Sub txtName_GotFocus()
Call SelectAll(txtName)

End Sub

Private Sub txtNumber_GotFocus()
Call SelectAll(txtNumber)

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
On Error GoTo Error:
Dim nStatus As Integer
Dim nDelete As Integer, temp As Long

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

nDelete = MsgBox("Delete this record from database?", vbYesNo, "Delete Record?")

If bLoaded Then Call saverecord(nCurrentRecord)

If Not nDelete = vbYes Then Exit Sub
    
nCurrentRecord = Val(lvDatabase.SelectedItem.Text)
temp = lvDatabase.SelectedItem.Index

nStatus = BTRCALL(BGETEQUAL, ClassPosBlock, Classdatabuf, Len(Classdatabuf), nCurrentRecord, KEY_BUF_LEN, 0)
If nStatus = 0 Then
    nStatus = BTRCALL(BDELETE, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
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
            
            Call modMain.LoadClassArray
        Else
            Call Form_Unload(1)
            Call Form_Load
        End If
    End If
Else
    MsgBox "Couldn't get record, Error: " & BtrieveErrorCode(nStatus)
End If

Exit Sub
Error:
Call HandleError
End Sub

Private Sub cmdInsert_Click()
On Error GoTo Error:
Dim nStatus As Integer
Dim nNewClassNumber As String, oLI As ListItem

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

If bLoaded = True Then Call saverecord(nCurrentRecord)

nNewClassNumber = InputBox("New Class Number:" & vbCrLf & vbCrLf & "Enter 0 for the next highest number.", "Insert", "0")
If nNewClassNumber = "" Then Exit Sub

Classrec.Number = Val(nNewClassNumber)
'Classrec.Name = "New Class" & Chr(0)
Call ClassStructToRow(Classdatabuf.buf)

nStatus = BTRCALL(BINSERT, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdInsert, BINSERT, Error: " & BtrieveErrorCode(nStatus)
    bLoaded = False
Else
    ClassRowToStruct Classdatabuf.buf
    
    Call AddClassToLV(Classrec.Number)
    
    SortListView lvDatabase, 1, ldtNumber, True
    
    nCurrentRecord = Classrec.Number
    DispClassInfo Classdatabuf.buf
    
    Set oLI = lvDatabase.FindItem(Classrec.Number, lvwText, , 0)
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
    Call modMain.LoadClassArray
End If

Set oLI = Nothing
Exit Sub
Error:
Call HandleError
Set oLI = Nothing
End Sub

Private Sub txtTitleText_GotFocus()
Call SelectAll(txtTitleText)

End Sub
