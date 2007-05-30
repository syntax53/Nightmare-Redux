VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20D5284F-7B23-4F0A-B8B1-6C9D18B64F1C}#1.0#0"; "exlimiter.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmRace 
   Caption         =   "Race Editor"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   Icon            =   "frmRace.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6375
   Begin VB.Frame framNav 
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   2700
      TabIndex        =   3
      Top             =   60
      Width           =   3615
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   315
         Left            =   840
         TabIndex        =   5
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "&Insert"
         Height          =   315
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdDiscard 
         Caption         =   "Dis&card"
         Height          =   315
         Left            =   2760
         TabIndex        =   7
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         Top             =   0
         Width           =   855
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5055
         Left            =   0
         TabIndex        =   8
         Top             =   480
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   8916
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "frmRace.frx":08CA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Line1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label(10)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label(9)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label(8)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label(1)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label(0)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label(7)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label(6)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label(5)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label(4)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Label(3)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label(2)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "txtHpBonus"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "txtCp"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "txtExpChart"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "txtName"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "txtNumber"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "txtMaxCharm"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "txtMaxWillpower"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "txtMaxHealth"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "txtMaxIntellect"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "txtMaxAgility"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "txtMaxStrength"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "txtMinCharm"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "txtMinWillpower"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "txtMinHealth"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "txtMinIntellect"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "txtMinAgility"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "txtMinStrength"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "chkAutoSave"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).ControlCount=   32
         TabCaption(1)   =   "Abilities"
         TabPicture(1)   =   "frmRace.frx":08E6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdAbilsClear"
         Tab(1).Control(1)=   "frmAbilities"
         Tab(1).ControlCount=   2
         Begin VB.CheckBox chkAutoSave 
            Caption         =   "Auto-Save"
            Height          =   195
            Left            =   2400
            TabIndex        =   71
            Top             =   420
            Value           =   1  'Checked
            Width           =   1155
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
            Left            =   -72240
            TabIndex        =   39
            Top             =   420
            Width           =   675
         End
         Begin VB.TextBox txtMinStrength 
            Height          =   315
            Left            =   1200
            TabIndex        =   22
            Top             =   2700
            Width           =   615
         End
         Begin VB.TextBox txtMinAgility 
            Height          =   315
            Left            =   1200
            TabIndex        =   25
            Top             =   3060
            Width           =   615
         End
         Begin VB.TextBox txtMinIntellect 
            Height          =   315
            Left            =   1200
            TabIndex        =   28
            Top             =   3420
            Width           =   615
         End
         Begin VB.TextBox txtMinHealth 
            Height          =   315
            Left            =   1200
            TabIndex        =   31
            Top             =   3780
            Width           =   615
         End
         Begin VB.TextBox txtMinWillpower 
            Height          =   315
            Left            =   1200
            TabIndex        =   34
            Top             =   4140
            Width           =   615
         End
         Begin VB.TextBox txtMinCharm 
            Height          =   315
            Left            =   1200
            TabIndex        =   37
            Top             =   4500
            Width           =   615
         End
         Begin VB.TextBox txtMaxStrength 
            Height          =   315
            Left            =   1920
            TabIndex        =   23
            Top             =   2700
            Width           =   615
         End
         Begin VB.TextBox txtMaxAgility 
            Height          =   315
            Left            =   1920
            TabIndex        =   26
            Top             =   3060
            Width           =   615
         End
         Begin VB.TextBox txtMaxIntellect 
            Height          =   315
            Left            =   1920
            TabIndex        =   29
            Top             =   3420
            Width           =   615
         End
         Begin VB.TextBox txtMaxHealth 
            Height          =   315
            Left            =   1920
            TabIndex        =   32
            Top             =   3780
            Width           =   615
         End
         Begin VB.TextBox txtMaxWillpower 
            Height          =   315
            Left            =   1920
            TabIndex        =   35
            Top             =   4140
            Width           =   615
         End
         Begin VB.TextBox txtMaxCharm 
            Height          =   315
            Left            =   1920
            TabIndex        =   38
            Top             =   4500
            Width           =   615
         End
         Begin VB.TextBox txtNumber 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   420
            Width           =   615
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   960
            MaxLength       =   29
            TabIndex        =   12
            Top             =   780
            Width           =   2535
         End
         Begin VB.TextBox txtExpChart 
            Height          =   315
            Left            =   1200
            TabIndex        =   14
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox txtCp 
            Height          =   315
            Left            =   1200
            TabIndex        =   16
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox txtHpBonus 
            Height          =   315
            Left            =   1200
            TabIndex        =   18
            Top             =   2040
            Width           =   615
         End
         Begin VB.Frame frmAbilities 
            Caption         =   "Abilities"
            Height          =   4095
            Left            =   -74880
            TabIndex        =   40
            Top             =   720
            Width           =   3375
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   9
               Left            =   120
               TabIndex        =   68
               Top             =   3600
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   8
               Left            =   120
               TabIndex        =   65
               Top             =   3240
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   7
               Left            =   120
               TabIndex        =   62
               Top             =   2880
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   6
               Left            =   120
               TabIndex        =   59
               Top             =   2520
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   5
               Left            =   120
               TabIndex        =   56
               Top             =   2160
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   4
               Left            =   120
               TabIndex        =   53
               Top             =   1800
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   3
               Left            =   120
               TabIndex        =   50
               Top             =   1440
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   2
               Left            =   120
               TabIndex        =   47
               Top             =   1080
               Width           =   495
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   1
               Left            =   120
               TabIndex        =   44
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   5
               Left            =   2640
               TabIndex        =   58
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   6
               Left            =   2640
               TabIndex        =   61
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   2520
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   7
               Left            =   2640
               TabIndex        =   64
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   2880
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   8
               Left            =   2640
               TabIndex        =   67
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   3240
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   9
               Left            =   2640
               TabIndex        =   70
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   3600
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   285
               Index           =   0
               Left            =   2640
               TabIndex        =   43
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   1
               Left            =   2640
               TabIndex        =   46
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   2
               Left            =   2640
               TabIndex        =   49
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   1080
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   3
               Left            =   2640
               TabIndex        =   52
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   1440
               Width           =   615
            End
            Begin VB.TextBox txtAbilityB 
               Height          =   315
               Index           =   4
               Left            =   2640
               TabIndex        =   55
               ToolTipText     =   "Enter the value for the ability here."
               Top             =   1800
               Width           =   615
            End
            Begin VB.TextBox txtAbilityA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   0
               Left            =   120
               TabIndex        =   41
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
               TabIndex        =   42
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
               TabIndex        =   45
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
               TabIndex        =   48
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
               TabIndex        =   51
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
               TabIndex        =   54
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
               TabIndex        =   57
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
               TabIndex        =   60
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
               TabIndex        =   63
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
               TabIndex        =   66
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
               TabIndex        =   69
               Text            =   "empty"
               Top             =   3600
               Width           =   1815
            End
         End
         Begin VB.Label Label 
            Caption         =   "Strength"
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   21
            Top             =   2700
            Width           =   735
         End
         Begin VB.Label Label 
            Caption         =   "Agility"
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   24
            Top             =   3060
            Width           =   735
         End
         Begin VB.Label Label 
            Caption         =   "Intellect"
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   27
            Top             =   3420
            Width           =   735
         End
         Begin VB.Label Label 
            Caption         =   "Health"
            Height          =   315
            Index           =   5
            Left            =   120
            TabIndex        =   30
            Top             =   3780
            Width           =   735
         End
         Begin VB.Label Label 
            Caption         =   "Willpower"
            Height          =   315
            Index           =   6
            Left            =   120
            TabIndex        =   33
            Top             =   4140
            Width           =   735
         End
         Begin VB.Label Label 
            Caption         =   "Charm"
            Height          =   315
            Index           =   7
            Left            =   120
            TabIndex        =   36
            Top             =   4500
            Width           =   735
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
            Top             =   420
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
            Top             =   780
            Width           =   855
         End
         Begin VB.Label Label 
            Caption         =   "Exp %"
            Height          =   315
            Index           =   8
            Left            =   120
            TabIndex        =   13
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label 
            Caption         =   "Starting Cp"
            Height          =   315
            Index           =   9
            Left            =   120
            TabIndex        =   15
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label 
            Caption         =   "Hp Bonus"
            Height          =   315
            Index           =   10
            Left            =   120
            TabIndex        =   17
            Top             =   2040
            Width           =   975
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   3480
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Min"
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
            Left            =   1200
            TabIndex        =   19
            Top             =   2520
            Width           =   615
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Max"
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
            Left            =   1920
            TabIndex        =   20
            Top             =   2520
            Width           =   615
         End
      End
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   60
      TabIndex        =   1
      Top             =   180
      Width           =   2535
   End
   Begin MSComctlLib.ListView lvDatabase 
      Height          =   5115
      Left            =   60
      TabIndex        =   2
      Top             =   480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   9022
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
      Left            =   5580
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
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmRace"
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
    .MINHEIGHT = 405 + (TITLEBAR_OFFSET / 10)
    .MINWIDTH = 435
    .CenterOnLoad = False
    .EnableLimiter = True
End With

Me.Top = ReadINI("Windows", "RaceTop")
Me.Left = ReadINI("Windows", "RaceLeft")
Me.Width = ReadINI("Windows", "RaceWidth")
Me.Height = ReadINI("Windows", "RaceHeight")

lvDatabase.ListItems.clear

Call LoadRaces

Me.Show
Me.SetFocus
txtSearch.SetFocus
If ReadINI("Windows", "RaceMaxed") = "1" Then Me.WindowState = vbMaximized
End Sub
Private Sub cmdDiscard_Click()
Dim nStatus As Integer

If lvDatabase.SelectedItem Is Nothing Or nCurrentRecord = 0 Then
    MsgBox "No current record."
    Exit Sub
End If

nStatus = BTRCALL(BGETEQUAL, RacePosBlock, Racedatabuf, Len(Racedatabuf), nCurrentRecord, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error on BGETEQUAL: " & BtrieveErrorCode(nStatus)
Else
    DispRaceInfo Racedatabuf.buf
End If
End Sub

Private Sub cmdSave_Click()
On Error GoTo Error:

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
If lvDatabase.SelectedItem Is Nothing Then Exit Sub

Call saverecord(nCurrentRecord)
'Call lvDatabase_ItemClick(lvDatabase.SelectedItem)

Dim oLI As ListItem
Set oLI = lvDatabase.FindItem(Racerec.Number, lvwText, , 0)
If Not oLI Is Nothing Then
    oLI.ListSubItems(1).Text = ClipNull(Racerec.Name)
    oLI.ListSubItems(2).Text = Racerec.ExpChart & "%"
    oLI.ListSubItems(3).Text = Racerec.MinStr & "-" & Racerec.MaxStr
    oLI.ListSubItems(4).Text = Racerec.MinAgl & "-" & Racerec.MaxAgl
    oLI.ListSubItems(5).Text = Racerec.MinInt & "-" & Racerec.MaxInt
    oLI.ListSubItems(6).Text = Racerec.MinHea & "-" & Racerec.MaxHea
    oLI.ListSubItems(7).Text = Racerec.MinWil & "-" & Racerec.MaxWil
    oLI.ListSubItems(8).Text = Racerec.MinChm & "-" & Racerec.MaxChm
End If
Set oLI = Nothing

out:
Exit Sub
Error:
Call HandleError("cmdSave_Click")
Resume out:

End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = vbMinimized Then Exit Sub
framNav.Left = Me.Width - framNav.Width - 220
lvDatabase.Width = framNav.Left - 100
lvDatabase.Height = Me.Height - 925 - TITLEBAR_OFFSET
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
If bLoaded = True Then
    saverecord (nCurrentRecord)
    Call modMain.LoadRaceArray
End If
If Me.WindowState = vbMinimized Then Exit Sub

If Me.WindowState = vbMaximized Then
    Call WriteINI("Windows", "RaceMaxed", 1)
Else
    Call WriteINI("Windows", "RaceMaxed", 0)
    Call WriteINI("Windows", "RaceTop", Me.Top)
    Call WriteINI("Windows", "RaceLeft", Me.Left)
    Call WriteINI("Windows", "RaceWidth", Me.Width)
    Call WriteINI("Windows", "RaceHeight", Me.Height)
End If
End Sub

Private Sub lvDatabase_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim nSort As ListDataType
Select Case ColumnHeader.Index
    Case 1, 3 To 9: nSort = ldtNumber
    Case Else: nSort = ldtString
End Select
SortListView lvDatabase, ColumnHeader.Index, nSort, lvDatabase.SortOrder
End Sub

Private Sub lvDatabase_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim temp As Integer, nStatus As Integer

If bLoaded = True And chkAutoSave.Value = 1 Then Call saverecord(nCurrentRecord)

temp = Val(Item.Text)
nStatus = BTRCALL(BGETEQUAL, RacePosBlock, Racedatabuf, Len(Racedatabuf), temp, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error on BGETEQUAL: " & BtrieveErrorCode(nStatus)
    bLoaded = False
Else
    nCurrentRecord = temp
    DispRaceInfo Racedatabuf.buf
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

Private Sub LoadRaces()
On Error GoTo Error:
Dim nStatus As Integer

lvDatabase.ColumnHeaders.clear
lvDatabase.ColumnHeaders.add 1, "Number", "#", 400, lvwColumnLeft
lvDatabase.ColumnHeaders.add 2, "Name", "Name", 1300, lvwColumnCenter
lvDatabase.ColumnHeaders.add 3, "Exp%", "Exp", 700, lvwColumnCenter
lvDatabase.ColumnHeaders.add 4, "STR", "STR", 800, lvwColumnCenter
lvDatabase.ColumnHeaders.add 5, "AGI", "AGI", 800, lvwColumnCenter
lvDatabase.ColumnHeaders.add 6, "INT", "INT", 800, lvwColumnCenter
lvDatabase.ColumnHeaders.add 7, "HEA", "HEA", 800, lvwColumnCenter
lvDatabase.ColumnHeaders.add 8, "WIS", "WIS", 800, lvwColumnCenter
lvDatabase.ColumnHeaders.add 9, "CHM", "CHM", 800, lvwColumnCenter

nStatus = BTRCALL(BGETFIRST, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "LoadRaces, BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

Do While nStatus = 0
    RaceRowToStruct Racedatabuf.buf
    
    Call AddRaceToLV(Racerec.Number)
    
    nStatus = BTRCALL(BGETNEXT, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "LoadRaces, Error: " & BtrieveErrorCode(nStatus)
End If

Call modMain.LoadRaceArray

If lvDatabase.ListItems.Count >= 1 Then Call lvDatabase_ItemClick(lvDatabase.ListItems(1))

lvDatabase.refresh
SortListView lvDatabase, 1, ldtNumber, True
bLoaded = True

Exit Sub
Error:
Call HandleError
End Sub

Private Sub AddRaceToLV(ByVal nNumber As Integer)
Dim nStatus As Integer, oLI As ListItem
On Error GoTo Error:

If Not nNumber = Racerec.Number Then
    nStatus = BTRCALL(BGETEQUAL, RacePosBlock, Racedatabuf, Len(Racedatabuf), nNumber, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then MsgBox "Error getting record " & nNumber & ": " & BtrieveErrorCode(nStatus)
    bLoaded = False
    Exit Sub
End If

Set oLI = lvDatabase.ListItems.add()
oLI.Text = Racerec.Number

oLI.ListSubItems.add (1), "Name", ClipNull(Racerec.Name)
oLI.ListSubItems.add (2), "Exp%", Racerec.ExpChart & "%"
oLI.ListSubItems.add (3), "STR", Racerec.MinStr & "-" & Racerec.MaxStr
oLI.ListSubItems.add (4), "AGI", Racerec.MinAgl & "-" & Racerec.MaxAgl
oLI.ListSubItems.add (5), "INT", Racerec.MinInt & "-" & Racerec.MaxInt
oLI.ListSubItems.add (6), "HEA", Racerec.MinHea & "-" & Racerec.MaxHea
oLI.ListSubItems.add (7), "WIS", Racerec.MinWil & "-" & Racerec.MaxWil
oLI.ListSubItems.add (8), "CHM", Racerec.MinChm & "-" & Racerec.MaxChm

Set oLI = Nothing
Exit Sub
Error:
Call HandleError
Set oLI = Nothing
End Sub

Private Sub DispRaceInfo(row() As Byte)
On Error GoTo Error:
Dim x As Integer

Call RaceRowToStruct(row())

Me.Caption = "Race Editor -- " & ClipNull(Racerec.Name)

txtNumber.Text = Racerec.Number
txtName.Text = Racerec.Name
txtExpChart.Text = Racerec.ExpChart
txtHpBonus.Text = Racerec.HPBonus
txtCp.Text = Racerec.CP
txtMinIntellect.Text = Racerec.MinInt
txtMinAgility.Text = Racerec.MinAgl
txtMinStrength.Text = Racerec.MinStr
txtMinWillpower.Text = Racerec.MinWil
txtMinCharm.Text = Racerec.MinChm
txtMinHealth.Text = Racerec.MinHea
txtMaxIntellect.Text = Racerec.MaxInt
txtMaxAgility.Text = Racerec.MaxAgl
txtMaxStrength.Text = Racerec.MaxStr
txtMaxWillpower.Text = Racerec.MaxWil
txtMaxCharm.Text = Racerec.MaxChm
txtMaxHealth.Text = Racerec.MaxHea

For x = 0 To 9
    txtAbilityA(x).Text = Racerec.AbilityA(x)
    txtAbilityB(x).Text = Racerec.AbilityB(x)
Next

Exit Sub
Error:
Call HandleError
MsgBox "Warning, record was not completely displayed." & vbCrLf _
    & "Previous records stats may still be in memory.  Select 'Disable DB Writing'" & vbCrLf _
    & "from the file menu and then reload the editor.", vbExclamation
End Sub

Private Sub saverecord(ByVal nRecord As Integer)
On Error GoTo Error:
Dim nStatus As Integer, x As Integer

If nRecord = 0 Then Exit Sub

nStatus = BTRCALL(BGETEQUAL, RacePosBlock, Racedatabuf, Len(Racedatabuf), nRecord, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error on BGETEQUAL: " & BtrieveErrorCode(nStatus)
    Exit Sub
Else
    RaceRowToStruct Racedatabuf.buf
End If

'DoEvents
Racerec.Name = txtName.Text & Chr(0)
Racerec.ExpChart = Val(txtExpChart.Text)
Racerec.HPBonus = Val(txtHpBonus.Text)
Racerec.CP = Val(txtCp.Text)
Racerec.MinInt = Val(txtMinIntellect.Text)
Racerec.MinAgl = Val(txtMinAgility.Text)
Racerec.MinStr = Val(txtMinStrength.Text)
Racerec.MinWil = Val(txtMinWillpower.Text)
Racerec.MinChm = Val(txtMinCharm.Text)
Racerec.MinHea = Val(txtMinHealth.Text)
Racerec.MaxInt = Val(txtMaxIntellect.Text)
Racerec.MaxAgl = Val(txtMaxAgility.Text)
Racerec.MaxStr = Val(txtMaxStrength.Text)
Racerec.MaxWil = Val(txtMaxWillpower.Text)
Racerec.MaxChm = Val(txtMaxCharm.Text)
Racerec.MaxHea = Val(txtMaxHealth.Text)

For x = 0 To 9
    Racerec.AbilityA(x) = Val(txtAbilityA(x).Text)
    Racerec.AbilityB(x) = Val(txtAbilityB(x).Text)
Next

nStatus = UpdateRace
If Not nStatus = 0 Then
    MsgBox "SaveRecord, Error: " & BtrieveErrorCode(nStatus)
Else
    DispRaceInfo Racedatabuf.buf
End If

Exit Sub
Error:
Call HandleError
End Sub


Private Sub txtAbilityB_GotFocus(Index As Integer)
Call SelectAll(txtAbilityB(Index))

End Sub

Private Sub txtCP_GotFocus()
Call SelectAll(txtCp)

End Sub

Private Sub txtExpChart_GotFocus()
Call SelectAll(txtExpChart)

End Sub

Private Sub txtHpBonus_GotFocus()
Call SelectAll(txtHpBonus)

End Sub

Private Sub txtMaxAgility_GotFocus()
Call SelectAll(txtMaxAgility)

End Sub

Private Sub txtMaxCharm_GotFocus()
Call SelectAll(txtMaxCharm)

End Sub

Private Sub txtMaxHealth_GotFocus()
Call SelectAll(txtMaxHealth)

End Sub

Private Sub txtMaxIntellect_GotFocus()
Call SelectAll(txtMaxIntellect)

End Sub

Private Sub txtMaxStrength_GotFocus()
Call SelectAll(txtMaxStrength)

End Sub

Private Sub txtMaxWillpower_GotFocus()
Call SelectAll(txtMaxWillpower)

End Sub

Private Sub txtMinAgility_GotFocus()
Call SelectAll(txtMinAgility)

End Sub

Private Sub txtMinCharm_GotFocus()
Call SelectAll(txtMinCharm)

End Sub

Private Sub txtMinHealth_GotFocus()
Call SelectAll(txtMinHealth)

End Sub

Private Sub txtMinIntellect_GotFocus()
Call SelectAll(txtMinIntellect)

End Sub

Private Sub txtMinStrength_GotFocus()
Call SelectAll(txtMinStrength)

End Sub

Private Sub txtMinWillpower_GotFocus()
Call SelectAll(txtMinWillpower)

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

nStatus = BTRCALL(BGETEQUAL, RacePosBlock, Racedatabuf, Len(Racedatabuf), nCurrentRecord, KEY_BUF_LEN, 0)
If nStatus = 0 Then
    nStatus = BTRCALL(BDELETE, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
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
            
            Call modMain.LoadRaceArray
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
Dim nNewRaceNumber As String, oLI As ListItem

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

If bLoaded = True Then Call saverecord(nCurrentRecord)

nNewRaceNumber = InputBox("New Race Number:" & vbCrLf & vbCrLf & "Enter 0 for the next highest number.", "Insert", "0")
If nNewRaceNumber = "" Then Exit Sub

Racerec.Number = Val(nNewRaceNumber)
'Racerec.Name = "New Race" & Chr(0)
Call RaceStructToRow(Racedatabuf.buf)

nStatus = BTRCALL(BINSERT, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdInsert, BINSERT, Error: " & BtrieveErrorCode(nStatus)
    bLoaded = False
Else
    RaceRowToStruct Racedatabuf.buf
    
    Call AddRaceToLV(Racerec.Number)
    
    SortListView lvDatabase, 1, ldtNumber, True
    
    nCurrentRecord = Racerec.Number
    DispRaceInfo Racedatabuf.buf
    
    Set oLI = lvDatabase.FindItem(Racerec.Number, lvwText, , 0)
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
    Call modMain.LoadRaceArray
End If

Set oLI = Nothing
Exit Sub
Error:
Call HandleError
Set oLI = Nothing
End Sub

