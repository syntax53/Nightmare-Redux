VERSION 5.00
Begin VB.Form frmMapEditor 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Editor"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   8730
   Icon            =   "frmMapEditor.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   8730
   Begin VB.Frame framTools 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   6720
      TabIndex        =   247
      Top             =   0
      Width           =   1935
      Begin VB.Frame framOptions 
         BackColor       =   &H00404040&
         Caption         =   "Options"
         ForeColor       =   &H00FFFFFF&
         Height          =   3315
         Left            =   60
         TabIndex        =   262
         Top             =   2940
         Visible         =   0   'False
         Width           =   1815
         Begin VB.CheckBox chkSmallForm 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Small Form"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   225
            Left            =   120
            TabIndex        =   270
            Top             =   2880
            Width           =   1335
         End
         Begin VB.CheckBox chkFollowMapChanges 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
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
            ForeColor       =   &H00E0E0E0&
            Height          =   225
            Left            =   120
            TabIndex        =   263
            Top             =   360
            Width           =   1635
         End
         Begin VB.CheckBox chkDontFollowHidden 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Don't Follow Hidden"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   225
            Left            =   120
            TabIndex        =   268
            Top             =   2160
            Width           =   1575
         End
         Begin VB.CheckBox chkMarkLair 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
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
            ForeColor       =   &H00E0E0E0&
            Height          =   225
            Left            =   120
            TabIndex        =   264
            Top             =   720
            Width           =   1515
         End
         Begin VB.CheckBox chkMarkCMD 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Mark Commands"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   225
            Left            =   120
            TabIndex        =   265
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CheckBox chkMarkNPC 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Mark Perm NPCs"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   225
            Left            =   120
            TabIndex        =   266
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CheckBox chkNoTooltips 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
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
            ForeColor       =   &H00E0E0E0&
            Height          =   225
            Left            =   120
            TabIndex        =   269
            Top             =   2520
            Width           =   1635
         End
         Begin VB.CheckBox chkDisplayNumbers 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Display Room #s"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   225
            Left            =   120
            TabIndex        =   267
            Top             =   1800
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00404040&
         Cancel          =   -1  'True
         Caption         =   "&Close"
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
         Left            =   60
         TabIndex        =   280
         Top             =   60
         Width           =   1860
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "&Options"
         Height          =   315
         Left            =   60
         TabIndex        =   261
         ToolTipText     =   "Show Legend Window"
         Top             =   2580
         Width           =   1875
      End
      Begin VB.TextBox txtIncrement 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   240
         TabIndex        =   277
         Top             =   5880
         Width           =   1395
      End
      Begin VB.CheckBox chkAutoUseSameMap 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Auto Use Same Map"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   60
         MaskColor       =   &H00404040&
         TabIndex        =   275
         Top             =   4620
         Width           =   1875
      End
      Begin VB.CheckBox chkDontAsk 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Dont Ask Anything"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   60
         MaskColor       =   &H00000000&
         TabIndex        =   271
         Top             =   3000
         Width           =   1875
      End
      Begin VB.CheckBox chkAutoUseRoom 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Auto Use Next Room"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   60
         MaskColor       =   &H00000000&
         TabIndex        =   276
         Top             =   5160
         Width           =   1875
      End
      Begin VB.CheckBox chkAutoCreate 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Auto Create"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   60
         MaskColor       =   &H00000000&
         TabIndex        =   272
         Top             =   3540
         Width           =   1875
      End
      Begin VB.CheckBox chkAutoCreateNew 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Auto New (5)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         MaskColor       =   &H00000000&
         TabIndex        =   273
         Top             =   3840
         Width           =   1635
      End
      Begin VB.CheckBox chkAutoSelectExisting 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Auto Exist (5)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         MaskColor       =   &H00000000&
         TabIndex        =   274
         Top             =   4140
         Width           =   1635
      End
      Begin VB.CommandButton cmdKeypad 
         BackColor       =   &H00808080&
         Caption         =   "(Keypad Enabled)"
         Height          =   315
         Left            =   60
         TabIndex        =   278
         Top             =   6360
         Width           =   1875
      End
      Begin VB.Frame framMove 
         BackColor       =   &H00404040&
         Caption         =   "Move"
         ForeColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   60
         TabIndex        =   248
         Top             =   360
         Width           =   1815
         Begin VB.CommandButton cmdReload 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   720
            Picture         =   "frmMapEditor.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   259
            ToolTipText     =   "Reload"
            Top             =   660
            Width           =   435
         End
         Begin VB.CommandButton cmdMove 
            BackColor       =   &H00C0C0C0&
            Caption         =   "D"
            Height          =   315
            Index           =   9
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   258
            Top             =   1380
            Width           =   435
         End
         Begin VB.CommandButton cmdMove 
            BackColor       =   &H00C0C0C0&
            Caption         =   "U"
            Height          =   315
            Index           =   8
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   257
            Top             =   1380
            Width           =   435
         End
         Begin VB.CommandButton cmdMove 
            BackColor       =   &H00C0C0C0&
            Caption         =   "SW"
            Height          =   315
            Index           =   7
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   256
            Top             =   1020
            Width           =   435
         End
         Begin VB.CommandButton cmdMove 
            BackColor       =   &H00C0C0C0&
            Caption         =   "SE"
            Height          =   315
            Index           =   6
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   255
            Top             =   1020
            Width           =   435
         End
         Begin VB.CommandButton cmdMove 
            BackColor       =   &H00C0C0C0&
            Caption         =   "NW"
            Height          =   315
            Index           =   5
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   254
            Top             =   300
            Width           =   435
         End
         Begin VB.CommandButton cmdMove 
            BackColor       =   &H00C0C0C0&
            Caption         =   "NE"
            Height          =   315
            Index           =   4
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   253
            Top             =   300
            Width           =   435
         End
         Begin VB.CommandButton cmdMove 
            BackColor       =   &H00C0C0C0&
            Caption         =   "W"
            Height          =   315
            Index           =   3
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   252
            Top             =   660
            Width           =   435
         End
         Begin VB.CommandButton cmdMove 
            BackColor       =   &H00C0C0C0&
            Caption         =   "E"
            Height          =   315
            Index           =   2
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   251
            Top             =   660
            Width           =   435
         End
         Begin VB.CommandButton cmdMove 
            BackColor       =   &H00C0C0C0&
            Caption         =   "S"
            Height          =   315
            Index           =   1
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   250
            Top             =   1020
            Width           =   435
         End
         Begin VB.CommandButton cmdMove 
            BackColor       =   &H00C0C0C0&
            Caption         =   "N"
            Height          =   315
            Index           =   0
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   249
            Top             =   300
            Width           =   435
         End
      End
      Begin VB.CommandButton cmdLegend 
         Caption         =   "&Legend / Help"
         Height          =   315
         Left            =   60
         TabIndex        =   260
         ToolTipText     =   "Show Legend Window"
         Top             =   2280
         Width           =   1875
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Next New RM #:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   279
         Top             =   5640
         Width           =   1425
      End
   End
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   60
      ScaleHeight     =   439
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   439
      TabIndex        =   1
      Top             =   60
      Width           =   6615
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   121
         Left            =   6075
         TabIndex        =   246
         Top             =   6270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   120
         Left            =   5475
         TabIndex        =   245
         Top             =   6270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   119
         Left            =   4875
         TabIndex        =   244
         Top             =   6270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   118
         Left            =   4275
         TabIndex        =   243
         Top             =   6270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   117
         Left            =   3675
         TabIndex        =   242
         Top             =   6270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   116
         Left            =   3075
         TabIndex        =   241
         Top             =   6270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   115
         Left            =   2475
         TabIndex        =   240
         Top             =   6270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   114
         Left            =   1875
         TabIndex        =   239
         Top             =   6270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   113
         Left            =   1275
         TabIndex        =   238
         Top             =   6270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   112
         Left            =   675
         TabIndex        =   237
         Top             =   6270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   111
         Left            =   75
         TabIndex        =   236
         Top             =   6270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   110
         Left            =   6075
         TabIndex        =   235
         Top             =   5670
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   109
         Left            =   5475
         TabIndex        =   234
         Top             =   5670
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   108
         Left            =   4875
         TabIndex        =   233
         Top             =   5670
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   107
         Left            =   4275
         TabIndex        =   232
         Top             =   5670
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   106
         Left            =   3675
         TabIndex        =   231
         Top             =   5670
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   105
         Left            =   3075
         TabIndex        =   230
         Top             =   5670
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   104
         Left            =   2475
         TabIndex        =   229
         Top             =   5670
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   103
         Left            =   1875
         TabIndex        =   228
         Top             =   5670
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   102
         Left            =   1275
         TabIndex        =   227
         Top             =   5670
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   101
         Left            =   675
         TabIndex        =   226
         Top             =   5670
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   100
         Left            =   75
         TabIndex        =   225
         Top             =   5670
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   99
         Left            =   6075
         TabIndex        =   224
         Top             =   5070
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   98
         Left            =   5475
         TabIndex        =   223
         Top             =   5070
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   97
         Left            =   4875
         TabIndex        =   222
         Top             =   5070
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   96
         Left            =   4275
         TabIndex        =   221
         Top             =   5070
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   95
         Left            =   3675
         TabIndex        =   220
         Top             =   5070
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   94
         Left            =   3075
         TabIndex        =   219
         Top             =   5070
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   93
         Left            =   2475
         TabIndex        =   218
         Top             =   5070
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   92
         Left            =   1875
         TabIndex        =   217
         Top             =   5070
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   91
         Left            =   1275
         TabIndex        =   216
         Top             =   5070
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   90
         Left            =   675
         TabIndex        =   215
         Top             =   5070
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   89
         Left            =   75
         TabIndex        =   214
         Top             =   5070
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   88
         Left            =   6075
         TabIndex        =   213
         Top             =   4470
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   87
         Left            =   5475
         TabIndex        =   212
         Top             =   4470
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   86
         Left            =   4875
         TabIndex        =   211
         Top             =   4470
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   85
         Left            =   4275
         TabIndex        =   210
         Top             =   4470
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   84
         Left            =   3675
         TabIndex        =   209
         Top             =   4470
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   83
         Left            =   3075
         TabIndex        =   208
         Top             =   4470
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   82
         Left            =   2475
         TabIndex        =   207
         Top             =   4470
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   81
         Left            =   1875
         TabIndex        =   206
         Top             =   4470
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   80
         Left            =   1275
         TabIndex        =   205
         Top             =   4470
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   79
         Left            =   675
         TabIndex        =   204
         Top             =   4470
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   78
         Left            =   75
         TabIndex        =   203
         Top             =   4470
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   77
         Left            =   6075
         TabIndex        =   202
         Top             =   3870
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   76
         Left            =   5475
         TabIndex        =   201
         Top             =   3870
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   75
         Left            =   4875
         TabIndex        =   200
         Top             =   3870
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   74
         Left            =   4275
         TabIndex        =   199
         Top             =   3870
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   73
         Left            =   3675
         TabIndex        =   198
         Top             =   3870
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   72
         Left            =   3075
         TabIndex        =   197
         Top             =   3870
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   71
         Left            =   2475
         TabIndex        =   196
         Top             =   3870
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   70
         Left            =   1875
         TabIndex        =   195
         Top             =   3870
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   69
         Left            =   1275
         TabIndex        =   194
         Top             =   3870
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   68
         Left            =   675
         TabIndex        =   193
         Top             =   3870
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   67
         Left            =   75
         TabIndex        =   192
         Top             =   3870
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   66
         Left            =   6075
         TabIndex        =   191
         Top             =   3270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   65
         Left            =   5475
         TabIndex        =   190
         Top             =   3270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   64
         Left            =   4875
         TabIndex        =   189
         Top             =   3270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   63
         Left            =   4275
         TabIndex        =   188
         Top             =   3270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   62
         Left            =   3675
         TabIndex        =   187
         Top             =   3270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   61
         Left            =   3075
         TabIndex        =   186
         Top             =   3270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   60
         Left            =   2475
         TabIndex        =   185
         Top             =   3270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   59
         Left            =   1875
         TabIndex        =   184
         Top             =   3270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   58
         Left            =   1275
         TabIndex        =   183
         Top             =   3270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   57
         Left            =   675
         TabIndex        =   182
         Top             =   3270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   56
         Left            =   75
         TabIndex        =   181
         Top             =   3270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   55
         Left            =   6075
         TabIndex        =   180
         Top             =   2670
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   54
         Left            =   5475
         TabIndex        =   179
         Top             =   2670
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   53
         Left            =   4875
         TabIndex        =   178
         Top             =   2670
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   52
         Left            =   4275
         TabIndex        =   177
         Top             =   2670
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   51
         Left            =   3675
         TabIndex        =   176
         Top             =   2670
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   50
         Left            =   3075
         TabIndex        =   175
         Top             =   2670
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   49
         Left            =   2475
         TabIndex        =   174
         Top             =   2670
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   48
         Left            =   1875
         TabIndex        =   173
         Top             =   2670
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   47
         Left            =   1275
         TabIndex        =   172
         Top             =   2670
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   46
         Left            =   675
         TabIndex        =   171
         Top             =   2670
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   45
         Left            =   75
         TabIndex        =   170
         Top             =   2670
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   44
         Left            =   6075
         TabIndex        =   169
         Top             =   2070
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   43
         Left            =   5475
         TabIndex        =   168
         Top             =   2070
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   42
         Left            =   4875
         TabIndex        =   167
         Top             =   2070
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   41
         Left            =   4275
         TabIndex        =   166
         Top             =   2070
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   40
         Left            =   3675
         TabIndex        =   165
         Top             =   2070
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   39
         Left            =   3075
         TabIndex        =   164
         Top             =   2070
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   38
         Left            =   2475
         TabIndex        =   163
         Top             =   2070
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   37
         Left            =   1875
         TabIndex        =   162
         Top             =   2070
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   36
         Left            =   1275
         TabIndex        =   161
         Top             =   2070
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   35
         Left            =   675
         TabIndex        =   160
         Top             =   2070
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   34
         Left            =   75
         TabIndex        =   159
         Top             =   2070
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   33
         Left            =   6075
         TabIndex        =   158
         Top             =   1470
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   32
         Left            =   5475
         TabIndex        =   157
         Top             =   1470
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   31
         Left            =   4875
         TabIndex        =   156
         Top             =   1470
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   30
         Left            =   4275
         TabIndex        =   155
         Top             =   1470
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   29
         Left            =   3675
         TabIndex        =   154
         Top             =   1470
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   28
         Left            =   3075
         TabIndex        =   153
         Top             =   1470
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   27
         Left            =   2475
         TabIndex        =   152
         Top             =   1470
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   26
         Left            =   1875
         TabIndex        =   151
         Top             =   1470
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   25
         Left            =   1275
         TabIndex        =   150
         Top             =   1470
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   24
         Left            =   675
         TabIndex        =   149
         Top             =   1470
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   23
         Left            =   75
         TabIndex        =   148
         Top             =   1470
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   22
         Left            =   6075
         TabIndex        =   147
         Top             =   870
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   21
         Left            =   5475
         TabIndex        =   146
         Top             =   870
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   20
         Left            =   4875
         TabIndex        =   145
         Top             =   870
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   19
         Left            =   4275
         TabIndex        =   144
         Top             =   870
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   18
         Left            =   3675
         TabIndex        =   143
         Top             =   870
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   17
         Left            =   3075
         TabIndex        =   142
         Top             =   870
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   16
         Left            =   2475
         TabIndex        =   141
         Top             =   870
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   15
         Left            =   1875
         TabIndex        =   140
         Top             =   870
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   14
         Left            =   1275
         TabIndex        =   139
         Top             =   870
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   13
         Left            =   675
         TabIndex        =   138
         Top             =   870
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   12
         Left            =   75
         TabIndex        =   137
         Top             =   870
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   11
         Left            =   6075
         TabIndex        =   136
         Top             =   270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   10
         Left            =   5475
         TabIndex        =   135
         Top             =   270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   9
         Left            =   4875
         TabIndex        =   134
         Top             =   270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   8
         Left            =   4275
         TabIndex        =   133
         Top             =   270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   7
         Left            =   3675
         TabIndex        =   132
         Top             =   270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   6
         Left            =   3075
         TabIndex        =   131
         Top             =   270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   5
         Left            =   2475
         TabIndex        =   130
         Top             =   270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   4
         Left            =   1875
         TabIndex        =   129
         Top             =   270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   3
         Left            =   1275
         TabIndex        =   128
         Top             =   270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   2
         Left            =   675
         TabIndex        =   127
         Top             =   270
         Width           =   465
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   125
         Top             =   270
         Width           =   465
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   10
         Left            =   5520
         TabIndex        =   122
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   9
         Left            =   4920
         TabIndex        =   121
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   8
         Left            =   4320
         TabIndex        =   120
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   7
         Left            =   3720
         TabIndex        =   119
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   6
         Left            =   3120
         TabIndex        =   118
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   5
         Left            =   2520
         TabIndex        =   117
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   4
         Left            =   1920
         TabIndex        =   116
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   1320
         TabIndex        =   115
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   114
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   113
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   11
         Left            =   6120
         TabIndex        =   112
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   12
         Left            =   120
         TabIndex        =   111
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   13
         Left            =   720
         TabIndex        =   110
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   14
         Left            =   1320
         TabIndex        =   109
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   15
         Left            =   1920
         TabIndex        =   108
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   16
         Left            =   2520
         TabIndex        =   107
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   17
         Left            =   3120
         TabIndex        =   106
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   18
         Left            =   3720
         TabIndex        =   105
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   19
         Left            =   4320
         TabIndex        =   104
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   20
         Left            =   4920
         TabIndex        =   103
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   21
         Left            =   5520
         TabIndex        =   102
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   22
         Left            =   6120
         TabIndex        =   101
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   23
         Left            =   120
         TabIndex        =   100
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   24
         Left            =   720
         TabIndex        =   99
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   25
         Left            =   1320
         TabIndex        =   98
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   26
         Left            =   1920
         TabIndex        =   97
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   27
         Left            =   2520
         TabIndex        =   96
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   28
         Left            =   3120
         TabIndex        =   95
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   29
         Left            =   3720
         TabIndex        =   94
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   30
         Left            =   4320
         TabIndex        =   93
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   31
         Left            =   4920
         TabIndex        =   92
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   32
         Left            =   5520
         TabIndex        =   91
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   33
         Left            =   6120
         TabIndex        =   90
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   34
         Left            =   120
         TabIndex        =   89
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   35
         Left            =   720
         TabIndex        =   88
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   36
         Left            =   1320
         TabIndex        =   87
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   37
         Left            =   1920
         TabIndex        =   86
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   38
         Left            =   2520
         TabIndex        =   85
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   39
         Left            =   3120
         TabIndex        =   84
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   40
         Left            =   3720
         TabIndex        =   83
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   41
         Left            =   4320
         TabIndex        =   82
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   42
         Left            =   4920
         TabIndex        =   81
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   43
         Left            =   5520
         TabIndex        =   80
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   44
         Left            =   6120
         TabIndex        =   79
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   45
         Left            =   120
         TabIndex        =   78
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   46
         Left            =   720
         TabIndex        =   77
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   47
         Left            =   1320
         TabIndex        =   76
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   48
         Left            =   1920
         TabIndex        =   75
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   49
         Left            =   2520
         TabIndex        =   74
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   50
         Left            =   3120
         TabIndex        =   73
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   51
         Left            =   3720
         TabIndex        =   72
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   52
         Left            =   4320
         TabIndex        =   71
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   53
         Left            =   4920
         TabIndex        =   70
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   54
         Left            =   5520
         TabIndex        =   69
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   55
         Left            =   6120
         TabIndex        =   68
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   56
         Left            =   120
         TabIndex        =   67
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   57
         Left            =   720
         TabIndex        =   66
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   58
         Left            =   1320
         TabIndex        =   65
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   59
         Left            =   1920
         TabIndex        =   64
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   60
         Left            =   2520
         TabIndex        =   63
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   61
         Left            =   3120
         TabIndex        =   62
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   62
         Left            =   3720
         TabIndex        =   61
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   63
         Left            =   4320
         TabIndex        =   60
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   64
         Left            =   4920
         TabIndex        =   59
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   65
         Left            =   5520
         TabIndex        =   58
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   66
         Left            =   6120
         TabIndex        =   57
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   67
         Left            =   120
         TabIndex        =   56
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   68
         Left            =   720
         TabIndex        =   55
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   69
         Left            =   1320
         TabIndex        =   54
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   70
         Left            =   1920
         TabIndex        =   53
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   71
         Left            =   2520
         TabIndex        =   52
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   72
         Left            =   3120
         TabIndex        =   51
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   73
         Left            =   3720
         TabIndex        =   50
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   74
         Left            =   4320
         TabIndex        =   49
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   75
         Left            =   4920
         TabIndex        =   48
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   76
         Left            =   5520
         TabIndex        =   47
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   77
         Left            =   6120
         TabIndex        =   46
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   78
         Left            =   120
         TabIndex        =   45
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   79
         Left            =   720
         TabIndex        =   44
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   80
         Left            =   1320
         TabIndex        =   43
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   81
         Left            =   1920
         TabIndex        =   42
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   82
         Left            =   2520
         TabIndex        =   41
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   83
         Left            =   3120
         TabIndex        =   40
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   84
         Left            =   3720
         TabIndex        =   39
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   85
         Left            =   4320
         TabIndex        =   38
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   86
         Left            =   4920
         TabIndex        =   37
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   87
         Left            =   5520
         TabIndex        =   36
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   88
         Left            =   6120
         TabIndex        =   35
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   89
         Left            =   120
         TabIndex        =   34
         Top             =   4920
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   90
         Left            =   720
         TabIndex        =   33
         Top             =   4920
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   91
         Left            =   1320
         TabIndex        =   32
         Top             =   4920
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   92
         Left            =   1920
         TabIndex        =   31
         Top             =   4920
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   93
         Left            =   2520
         TabIndex        =   30
         Top             =   4920
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   94
         Left            =   3120
         TabIndex        =   29
         Top             =   4920
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   95
         Left            =   3720
         TabIndex        =   28
         Top             =   4920
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   96
         Left            =   4320
         TabIndex        =   27
         Top             =   4920
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   97
         Left            =   4920
         TabIndex        =   26
         Top             =   4920
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   98
         Left            =   5520
         TabIndex        =   25
         Top             =   4920
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   99
         Left            =   6120
         TabIndex        =   24
         Top             =   4920
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   100
         Left            =   120
         TabIndex        =   23
         Top             =   5520
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   101
         Left            =   720
         TabIndex        =   22
         Top             =   5520
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   102
         Left            =   1320
         TabIndex        =   21
         Top             =   5520
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   103
         Left            =   1920
         TabIndex        =   20
         Top             =   5520
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   104
         Left            =   2520
         TabIndex        =   19
         Top             =   5520
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   105
         Left            =   3120
         TabIndex        =   18
         Top             =   5520
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   106
         Left            =   3720
         TabIndex        =   17
         Top             =   5520
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   107
         Left            =   4320
         TabIndex        =   16
         Top             =   5520
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   108
         Left            =   4920
         TabIndex        =   15
         Top             =   5520
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   109
         Left            =   5520
         TabIndex        =   14
         Top             =   5520
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   110
         Left            =   6120
         TabIndex        =   13
         Top             =   5520
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   111
         Left            =   120
         TabIndex        =   12
         Top             =   6120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   112
         Left            =   720
         TabIndex        =   11
         Top             =   6120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   113
         Left            =   1320
         TabIndex        =   10
         Top             =   6120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   114
         Left            =   1920
         TabIndex        =   9
         Top             =   6120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   115
         Left            =   2520
         TabIndex        =   8
         Top             =   6120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   116
         Left            =   3120
         TabIndex        =   7
         Top             =   6120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   117
         Left            =   3720
         TabIndex        =   6
         Top             =   6120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   118
         Left            =   4320
         TabIndex        =   5
         Top             =   6120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   119
         Left            =   4920
         TabIndex        =   4
         Top             =   6120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   120
         Left            =   5520
         TabIndex        =   3
         Top             =   6120
         Width           =   375
      End
      Begin VB.Label lblCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   121
         Left            =   6120
         TabIndex        =   2
         Top             =   6120
         Width           =   375
      End
      Begin VB.Label lblCellBG 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   410
         Left            =   3100
         TabIndex        =   123
         Top             =   3100
         Width           =   410
      End
   End
   Begin VB.PictureBox picBG 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   60
      ScaleHeight     =   439
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   439
      TabIndex        =   124
      Top             =   60
      Width           =   6615
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   1140
      TabIndex        =   126
      Top             =   180
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lblCell 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmMapEditor"
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

'Dim DB As Database
'Dim tabRooms As Recordset
'Dim tabInfo As Recordset
'Dim UpdateExistingADB As Boolean
'Dim DataSource As String
'Dim ExportPath As String

Dim nBackcolor As Long
Dim StartRoom As Long
Dim StartMap As Long
Dim CenterCell As Integer
Dim objTooltip As clsToolTip

Dim CellRoom(1 To 121, 1 To 2) As Long
'Dim CellLabel(120) As String
Dim UnchartedCells(1 To 121) As Integer
Dim StopBuild As Boolean

'Private Const STATE_SYSTEM_FOCUSABLE = &H100000
'Private Const STATE_SYSTEM_INVISIBLE = &H8000
'Private Const STATE_SYSTEM_OFFSCREEN = &H10000
'Private Const STATE_SYSTEM_UNAVAILABLE = &H1
'Private Const STATE_SYSTEM_PRESSED = &H8
'Private Const CCHILDREN_TITLEBAR = 5

Dim EnableKeypad As Boolean
'Dim UpdateFailed As Boolean

Private Sub Form_Load()
On Error Resume Next
Dim x As Integer

Set objTooltip = New clsToolTip
With objTooltip
    .DelayTime = 20
    .VisibleTime = 20000
    .BkColor = &HC0FFFF
    .txtColor = &H0
    .Style = ttStyleBalloon
    '.Style = ttStyleStandard
End With

objTooltip.Style = 1

picMap.ScaleMode = vbPixels
picMap.ScaleHeight = 439
picMap.ScaleWidth = 439

frmRoom.cmdPrevious.Enabled = False
frmRoom.cmdNext.Enabled = False
frmRoom.cmdLast.Enabled = False
frmRoom.cmdFirst.Enabled = False

Me.Top = Val(ReadINI("Windows", "MapEditorTop"))
Me.Left = Val(ReadINI("Windows", "MapEditorLeft"))

chkDontAsk.Value = ReadINI("Options", "MapEditorDontAskAnything")
chkAutoUseSameMap.Value = ReadINI("Options", "MapEditorAutoSame")
chkAutoUseRoom.Value = ReadINI("Options", "MapEditorAutoNext")
chkAutoCreate.Value = ReadINI("Options", "MapEditorAutoCreate")
chkAutoCreateNew.Value = ReadINI("Options", "MapEditorAutoNew")
chkAutoSelectExisting.Value = ReadINI("Options", "MapEditorAutoExisting")
txtIncrement.Text = ReadINI("Options", "MapEditorNextNewRoom")

chkDisplayNumbers.Value = ReadINI("Options", "MapEditorDisplayNumbers")
chkFollowMapChanges.Value = ReadINI("Options", "MapEditorFollowMapChanges")
chkDontFollowHidden.Value = ReadINI("Options", "MapEditorDontFollowHidden")
chkMarkLair.Value = ReadINI("Options", "MapEditorMarkLair")
chkMarkCMD.Value = ReadINI("Options", "MapEditorMarkCMD")
chkMarkNPC.Value = ReadINI("Options", "MapEditorMarkNPC")
chkNoTooltips.Value = ReadINI("Options", "MapEditorNoTooltips")
chkSmallForm.Value = ReadINI("Options", "MapEditorSmallForm")

CenterCell = 61
Call StartMapping

Me.Show
Me.SetFocus
cmdClose.SetFocus
EnableKeypad = True

End Sub

Private Sub chkSmallForm_Click()

If chkSmallForm.Value = 1 Then
    framTools.Left = 60
    Me.Width = framTools.Left + framTools.Width + 120
    picMap.Visible = False
    picBG.Visible = False
Else
    Call StartMapping
    framTools.Left = picMap.Left + picMap.Width + 105
    Me.Width = picMap.Width + framTools.Width + 270
End If

End Sub

Private Sub cmdLegend_Click()

If FormIsLoaded("frmMapLegend") Then
    Unload frmMapLegend
Else
    frmMapLegend.Show
    frmMapLegend.SetFocus
End If
End Sub

Private Sub cmdMove_Click(Index As Integer)
Dim nStatus As Integer
On Error GoTo error:

    framOptions.Visible = False
    Call frmRoom.saverecord
    Me.Enabled = False
    
    If Not frmRoom.txtRoomExit(Index) = 0 Then
        
        If frmRoom.cmbRoomType(Index).ListIndex = 8 Then
            nStatus = GotoRoom(Val(frmRoom.txtRoomPara(Index).Text), Val(frmRoom.txtRoomExit(Index).Text))
        Else
            nStatus = GotoRoom(Val(frmRoom.txtMap.Text), Val(frmRoom.txtRoomExit(Index).Text))
        End If
        
        If Not nStatus = 0 Then
            If nStatus = 4 Then GoTo DoesntExist:
            MsgBox "BGETEQUAL, Room, Error: " & BtrieveErrorCode(nStatus)
            GoTo quit:
        End If
            
        Call StartMapping
        
    Else
        If chkDontAsk.Value = 1 Then GoTo quit:
        CheckCreateRoom Index, Val(frmRoom.txtMap.Text), Val(frmRoom.txtRoom.Text)
    End If

GoTo quit:

DoesntExist:
If chkDontAsk.Value = 1 Then GoTo quit:
Call CheckCreateRoom(Index, Val(frmRoom.txtMap.Text), Val(frmRoom.txtRoom.Text))

quit:
Me.Enabled = True
Me.SetFocus
Exit Sub

error:
Call HandleError
Resume quit:
End Sub

Private Sub cmdOptions_Click()
If framOptions.Visible = True Then framOptions.Visible = False Else framOptions.Visible = True
End Sub

Private Sub cmdReload_Click()
Call StartMapping
End Sub

Private Sub chkAutoCreate_Click()
If chkAutoCreate.Value = 1 Then
    chkAutoCreateNew.Enabled = True
    chkAutoSelectExisting.Enabled = True
Else
    chkAutoCreateNew.Value = 0
    chkAutoSelectExisting.Value = 0
    chkAutoCreateNew.Enabled = False
    chkAutoSelectExisting.Enabled = False
End If
End Sub

Private Sub chkAutoCreateNew_Click()
If chkAutoCreateNew.Value = 1 And chkAutoSelectExisting.Value = 1 Then chkAutoSelectExisting.Value = 0
End Sub

Private Sub chkAutoSelectExisting_Click()
If chkAutoCreateNew.Value = 1 And chkAutoSelectExisting.Value = 1 Then chkAutoCreateNew.Value = 0
End Sub

Private Sub chkDontAsk_Click()
If chkDontAsk.Value = 1 Then
    'chkAutoUseRoom.Value = 0
    chkAutoUseRoom.Enabled = False
    'chkAutoCreate.Value = 0
    chkAutoCreate.Enabled = False
    'chkAutoUseSameMap.Value = 0
    chkAutoUseSameMap.Enabled = False
Else
    chkAutoUseRoom.Enabled = True
    chkAutoCreate.Enabled = True
    chkAutoUseSameMap.Enabled = True
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdKeypad_Click()
EnableKeypad = True
cmdKeypad.Caption = "(Keypad Enabled)"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim nStatus As Integer
On Error Resume Next

    frmRoom.cmdGoto.Enabled = True
    frmRoom.cmdInsert.Enabled = True
    frmRoom.cmdDelete.Enabled = True
    frmRoom.cmdPlusOneRoom.Enabled = True
    frmRoom.cmdMinusOneRoom.Enabled = True
    frmRoom.cmdCopy2Exist.Enabled = True
    frmRoom.cmdCopy2New.Enabled = True
    frmRoom.cmdPrevious.Enabled = True
    frmRoom.cmdNext.Enabled = True
    frmRoom.cmdLast.Enabled = True
    frmRoom.cmdFirst.Enabled = True

    RoomKeyStruct.MapNum = Val(frmRoom.txtMap.Text)
    RoomKeyStruct.RoomNum = Val(frmRoom.txtRoom.Text)

    nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "BGETEQUAL, Room, Error: " & BtrieveErrorCode(nStatus)
    Else
        Call frmRoom.DispRoomInfo(Roomdatabuf.buf)
    End If

    Call WriteINI("Options", "MapEditorDontAskAnything", chkDontAsk.Value)
    Call WriteINI("Options", "MapEditorAutoSame", chkAutoUseSameMap.Value)
    Call WriteINI("Options", "MapEditorAutoNext", chkAutoUseRoom.Value)
    Call WriteINI("Options", "MapEditorAutoCreate", chkAutoCreate.Value)
    Call WriteINI("Options", "MapEditorAutoNew", chkAutoCreateNew.Value)
    Call WriteINI("Options", "MapEditorAutoExisting", chkAutoSelectExisting.Value)
    Call WriteINI("Options", "MapEditorNextNewRoom", txtIncrement.Text)
    
    If Not Me.WindowState = vbMinimized Then
        Call WriteINI("Windows", "MapEditorTop", Me.Top)
        Call WriteINI("Windows", "MapEditorLeft", Me.Left)
    End If
    
    Call WriteINI("Options", "MapEditorFollowMapChanges", chkFollowMapChanges.Value)
    Call WriteINI("Options", "MapEditorDontFollowHidden", chkDontFollowHidden.Value)
    Call WriteINI("Options", "MapEditorMarkLair", chkMarkLair.Value)
    Call WriteINI("Options", "MapEditorMarkCMD", chkMarkCMD.Value)
    Call WriteINI("Options", "MapEditorMarkNPC", chkMarkNPC.Value)
    Call WriteINI("Options", "MapEditorNoTooltips", chkNoTooltips.Value)
    Call WriteINI("Options", "MapEditorDisplayNumbers", chkDisplayNumbers.Value)
    Call WriteINI("Options", "MapEditorSmallForm", chkSmallForm.Value)
End Sub


Public Sub StartMapping()
On Error GoTo error:
Dim x As Integer, nMapSize As Integer, bCheckAgain As Boolean

StartRoom = frmRoom.txtRoom
StartMap = frmRoom.txtMap

nLastMapClick = StartMap
nLastRoomClick = StartRoom

SECorner = 121
RowLength = 11

nBackcolor = &H404040

If CenterCell > SECorner Then CenterCell = 60

If chkSmallForm.Value = 1 Then
    Call MapExits(CenterCell, StartRoom, StartMap)
    Exit Sub
End If

picMap.Cls
picBG.Visible = True
picMap.Visible = False
frmMain.Enabled = False
'framOptions.Visible = False

For x = 1 To 121
    objTooltip.DelToolTip picMap.hwnd, x
    lblCell(x).BackColor = nBackcolor
    lblCell(x).Visible = True
    lblCell(x).Tag = 0
    UnchartedCells(x) = 0
    CellRoom(x, 1) = 0
    CellRoom(x, 2) = 0
    lblNumber(x).Caption = ""
Next x

StopBuild = False

'CellLabel(CenterCell) = StartMap & "," & StartRoom
CellRoom(CenterCell, 1) = StartMap
CellRoom(CenterCell, 2) = StartRoom

Call MapExits(CenterCell, StartRoom, StartMap)

'frmMain.Enabled = False
'frmProgressBar.sCaption = "Map Editor"
'frmProgressBar.lblCaption = "Please wait, " & vbCrLf & "Creating Map..."
'frmProgressBar.cmdCancel.Enabled = False
'Call frmProgressBar.SetRange(SECorner)
'frmProgressBar.lblPanel(0).Caption = "Current Map,Room:"
'frmProgressBar.Show

DoEvents
again:
bCheckAgain = False
For x = 1 To SECorner
    If StopBuild = True Then GoTo Cancel:
    If UnchartedCells(x) = 1 Then
        Call MapExits(x, CellRoom(x, 2), CellRoom(x, 1))
        bCheckAgain = True
    End If
    If Not bUseCPU Then DoEvents
Next x
If bCheckAgain Then GoTo again:

'If chkNoColors.Value = 0 And chkDontMarkStart.Value = 0 Then
'   Call DrawOnRoom(lblCell(CenterCell), drSquare, 4, BrightBlue)
'End If

frmMain.Enabled = True
'Unload frmProgressBar

DoEvents
'If cmdHideUnused.Caption = "Show &Blocks" Then
    For x = 1 To 121
        If CellRoom(x, 1) = 0 Then lblCell(x).Visible = False
    Next x
    DoEvents
'End If

Cancel:
Me.Show
frmMain.Enabled = True
picMap.Visible = True
'Unload frmProgressBar
'Unload Me

Exit Sub
error:
Call HandleError
'Unload frmProgressBar
Me.Show
DoEvents

End Sub

Private Sub ToggleAuto()

If chkAutoCreateNew.Value = 1 Then
    chkAutoSelectExisting.Value = 1
ElseIf chkAutoSelectExisting.Value = 1 Then
    chkAutoCreateNew.Value = 1
Else
    If chkAutoCreate.Enabled Then
        chkAutoCreate.Value = 1
        chkAutoCreateNew.Value = 1
    End If
End If

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If EnableKeypad = False Then Exit Sub

Select Case KeyCode
    Case 107:
        If chkDontAsk.Value = 1 Then
            chkDontAsk.Value = 0
        Else
            chkDontAsk.Value = 1
        End If
    Case 38, 104: Call cmdMove_Click(0)    'cmdN_Click
    Case 40, 98: Call cmdMove_Click(1)  'cmdS_Click
    Case 37, 100: Call cmdMove_Click(3)  'cmdW_Click
    Case 39, 102: Call cmdMove_Click(2)  'cmdE_Click
    Case 36, 103: Call cmdMove_Click(5)  'cmdNW_Click
    Case 35, 97: Call cmdMove_Click(7)  'cmdSW_Click
    Case 33, 105: Call cmdMove_Click(4)  'cmdNE_click
    Case 34, 99: Call cmdMove_Click(6)  'cmdSE_Click
    Case 45, 96: Call cmdMove_Click(8)  'cmdU_Click
    Case 46, 110: Call cmdMove_Click(9)  'cmdD_Click
    Case 38, 101: Call ToggleAuto
End Select

'38 = arrow up (north)
'37 = arrow left (west)
'40 = arrow down (south)
'39 = arrow right (east)
'36 = home (nw)
'35 = end (sw)
'34 = pg down (se)
'33 = pg up (ne)
'45 = ins (up)
'46 = del (down)

'104 = 8 (north)
'100 = 4 (west)
'98  = 2 (south)
'102 = 6 (east)
'103 = 7 (nw)
'97  = 1 (sw)
'99 = 3 (se)
'105 = 9 (ne)
'96 = 0 (up)
'110 = . (down)

End Sub
Private Sub CheckCreateRoom(ByVal nDir As Integer, ByVal MapNum As Long, ByVal RoomNum As Long)
On Error GoTo error:
Dim nStatus As Integer, x As Integer, nYesNo As Integer
Dim nNewMapNumber As Variant, nNewRoomNumber As Variant, NewExit As Integer

If chkAutoCreateNew.Value = 1 Then: Call CreateNewRoom(nDir, MapNum, RoomNum): Exit Sub
If chkAutoSelectExisting.Value = 1 Then: Call SelectExistingRoom(nDir, MapNum, RoomNum): Exit Sub
If chkAutoCreate.Value = 1 Then:    Call CreateRoom(nDir, MapNum, RoomNum): Exit Sub

'If rname2(nDir).Caption = "ROOM DOESN'T EXIST" Then
'    nYesNo = MsgBox("The room to the " & GetRoomExits(nDir, True) & " doesn't exist." & vbCrLf & vbCrLf & "Would you like to create or select one?", vbYesNo + vbQuestion + vbDefaultButton2)
'    If nYesNo <> vbYes Then Exit Sub
'Else
    nYesNo = MsgBox("There is no exit to the " & GetRoomExits(nDir, True) & vbCrLf & vbCrLf & "Would you like to create or select one?", vbYesNo + vbQuestion + vbDefaultButton2)
    If nYesNo <> vbYes Then Exit Sub
'End If

Call CreateRoom(nDir, MapNum, RoomNum)
Exit Sub
error:
Call HandleError
End Sub

Private Sub CreateRoom(direction As Integer, MapNum As Long, RoomNum As Long)
On Error GoTo error:
Dim nStatus As Integer, x As Integer, nYesNo As Integer

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
nYesNo = MsgBox("Will this be an existing room?", vbYesNoCancel + vbQuestion + vbDefaultButton2)
If nYesNo <> vbYes And nYesNo <> vbNo Then Exit Sub

If nYesNo = vbNo Then: Call CreateNewRoom(direction, MapNum, RoomNum): Exit Sub
If nYesNo = vbYes Then: Call SelectExistingRoom(direction, MapNum, RoomNum): Exit Sub

Exit Sub
error:
Call HandleError
End Sub

Private Sub CreateNewRoom(ByVal nDir As Integer, ByVal MapNum As Long, ByVal RoomNum As Long)
On Error GoTo error:
Dim nStatus As Integer, x As Integer, nYesNo As Integer, sInput As String
Dim nNewMapNumber As Long, nNewRoomNumber As Long, NewExit As Integer

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
If chkAutoUseSameMap.Value = 1 Then 'auto use same map?
    nNewMapNumber = Val(frmRoom.txtMap.Text)
Else
    sInput = InputBox("Map Number:", "Creating New Room", MapNum)
    If sInput = "" Then Exit Sub
    nNewMapNumber = Val(sInput)
End If

If chkAutoUseRoom.Value = 1 Then 'auto use next room?
    nNewRoomNumber = Val(txtIncrement.Text)
Else
    sInput = InputBox("Room Number:", "Creating New Room", Val(txtIncrement.Text))
    If sInput = "" Then Exit Sub
    nNewRoomNumber = Val(sInput)
End If

'get current room
RoomKeyStruct.MapNum = Val(frmRoom.txtMap.Text)
RoomKeyStruct.RoomNum = Val(frmRoom.txtRoom.Text)

nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error getting current room: " & BtrieveErrorCode(nStatus)
    frmRoom.bLoaded = False: Call frmRoom.GotoRoom(Val(frmRoom.txtMap.Text), Val(frmRoom.txtRoom.Text), True)
    Exit Sub
Else
    Call frmRoom.DispRoomInfo(Roomdatabuf.buf)
End If

'set the properties for the new room
Roomrec.MapNumber = nNewMapNumber
Roomrec.RoomNumber = nNewRoomNumber

Roomrec.PermNPC = 0
Roomrec.Runic = 0
Roomrec.Platinum = 0
Roomrec.Gold = 0
Roomrec.Silver = 0
Roomrec.Copper = 0
Roomrec.InvisRunic = 0
Roomrec.InvisPlatinum = 0
Roomrec.InvisGold = 0
Roomrec.InvisSilver = 0
Roomrec.InvisCopper = 0

For x = 0 To 16
    Roomrec.RoomItems(x) = 0
    Roomrec.RoomItemQty(x) = 0
    Roomrec.RoomItemUses(x) = 0
Next
For x = 0 To 14
    Roomrec.InvisItems(x) = 0
    Roomrec.InvisItemQty(x) = 0
    Roomrec.InvisItemUses(x) = 0
    Roomrec.CurrentRoomMon(x) = 0
Next
For x = 0 To 9
    Roomrec.RoomExit(x) = 0
    Roomrec.RoomType(x) = 0
    Roomrec.Para1(x) = 0
    Roomrec.Para2(x) = 0
    Roomrec.Para3(x) = 0
    Roomrec.Para4(x) = 0
    Roomrec.PlacedItems(x) = 0
Next

'create exit back in the new room
Select Case nDir
    Case 0: NewExit = 1
    Case 1: NewExit = 0
    Case 2: NewExit = 3
    Case 3: NewExit = 2
    Case 4: NewExit = 7
    Case 5: NewExit = 6
    Case 6: NewExit = 5
    Case 7: NewExit = 4
    Case 8: NewExit = 9
    Case 9: NewExit = 8
End Select

Roomrec.RoomExit(NewExit) = RoomNum

'check for map change
If nNewMapNumber <> MapNum Then
    Roomrec.RoomType(NewExit) = 8
    Roomrec.Para1(NewExit) = MapNum
Else
    Roomrec.RoomType(NewExit) = 0
End If

'insert the room
RoomStructToRow Roomdatabuf.buf

nStatus = BTRCALL(BINSERT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    If nStatus = 5 Then
        nYesNo = MsgBox("Room " & nNewMapNumber & "/" & nNewRoomNumber & " already exists.  Connect to this room?" _
            , vbQuestion + vbYesNo + vbDefaultButton2, "Room Exists ...")
        If nYesNo = vbYes Then Call SelectExistingRoom(nDir, MapNum, RoomNum, nNewMapNumber, nNewRoomNumber)
        If nNewRoomNumber = Val(txtIncrement.Text) Then txtIncrement.Text = Val(txtIncrement.Text) + 1
        Exit Sub
    Else
        MsgBox "Error inserting new room!, BINSERT Error: " & BtrieveErrorCode(nStatus)
    End If
    
    frmRoom.bLoaded = False: Call frmRoom.GotoRoom(Val(frmRoom.txtMap.Text), Val(frmRoom.txtRoom.Text), True)
    Exit Sub
Else
    If nNewRoomNumber = Val(txtIncrement.Text) Then txtIncrement.Text = Val(txtIncrement.Text) + 1
End If

'reget current room
RoomKeyStruct.MapNum = Val(frmRoom.txtMap.Text)
RoomKeyStruct.RoomNum = Val(frmRoom.txtRoom.Text)

nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "New room inserted, but error re-getting current room: " & BtrieveErrorCode(nStatus)
    frmRoom.bLoaded = False: Call frmRoom.GotoRoom(Val(frmRoom.txtMap.Text), Val(frmRoom.txtRoom.Text), True)
    Exit Sub
End If
Call RoomRowToStruct(Roomdatabuf.buf)

'set exit to the new room on this current room
Roomrec.RoomExit(nDir) = nNewRoomNumber

'check if its the same map, if not make it a map change
If nNewMapNumber <> MapNum Then
    Roomrec.RoomType(nDir) = 8
    Roomrec.Para1(nDir) = nNewMapNumber
Else
    Roomrec.RoomType(nDir) = 0
End If

'update current room
nStatus = UpdateRoom
If Not nStatus = 0 Then
    MsgBox "New room inserted, but error updating current room: " & BtrieveErrorCode(nStatus)
    frmRoom.bLoaded = False: Call frmRoom.GotoRoom(Val(frmRoom.txtMap.Text), Val(frmRoom.txtRoom.Text), True)
    Exit Sub
End If

'reget new room
RoomKeyStruct.MapNum = nNewMapNumber
RoomKeyStruct.RoomNum = nNewRoomNumber

nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Rooms inserted and updated, but error re-getting new room: " & BtrieveErrorCode(nStatus)
    frmRoom.bLoaded = False: Call frmRoom.GotoRoom(Val(frmRoom.txtMap.Text), Val(frmRoom.txtRoom.Text), True)
    Exit Sub
Else
    Call frmRoom.DispRoomInfo(Roomdatabuf.buf)
End If

Call StartMapping
Exit Sub
error:
Call HandleError
End Sub

Private Function GetDirectionalRoomNumber(ByVal nDirection As Integer)
Dim nCell As Integer, nTries As Integer

nCell = CenterCell

again:
nTries = nTries + 1
Select Case nDirection
    Case 0: nCell = nCell - 11 'n
    Case 1: nCell = nCell + 11 's
    Case 2: nCell = nCell + 1 'e
    Case 3: nCell = nCell - 1 'w
    Case 4: nCell = nCell - 10 'ne
    Case 5: nCell = nCell - 12 'nw
    Case 6: nCell = nCell + 12 'se
    Case 7: nCell = nCell + 10 'sw
End Select

If nCell < 1 Or nCell > 121 Then
    GetDirectionalRoomNumber = -1
Else
    If CellRoom(nCell, 2) > 0 Then
        GetDirectionalRoomNumber = CellRoom(nCell, 2)
    Else
        'If nTries < 3 Then GoTo again:
        GetDirectionalRoomNumber = -1
    End If
End If

End Function
Private Sub SelectExistingRoom(ByVal nDir As Integer, ByVal MapNum As Long, ByVal RoomNum As Long, _
    Optional ByVal ExistMap As Long, Optional ByVal ExistRoom As Long)
On Error GoTo error:
Dim nStatus As Integer, x As Integer, sInput As String, nYesNo As Integer
Dim nNewMapNumber As Long, nNewRoomNumber As Long, NewExit As Integer

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
'check for auto use same map

If ExistMap < 1 Then
    If chkAutoUseSameMap.Value = 1 Then
        nNewMapNumber = Val(frmRoom.txtMap.Text)
    Else
        sInput = InputBox("Existing Room, Map Number:", "Enter Existing Map Number", MapNum)
        If sInput = "" Then Exit Sub
        nNewMapNumber = Val(sInput)
    End If
Else
    nNewMapNumber = ExistMap
End If

If ExistRoom < 1 Then
    nNewRoomNumber = GetDirectionalRoomNumber(nDir)
    sInput = InputBox("Existing Room, Room Number:", "Enter Existing Room Number", nNewRoomNumber)
    If sInput = "" Then Exit Sub
    nNewRoomNumber = Val(sInput)
    If nNewRoomNumber < 1 Then Exit Sub
Else
    nNewRoomNumber = ExistRoom
End If

'get the existing room
RoomKeyStruct.MapNum = nNewMapNumber
RoomKeyStruct.RoomNum = nNewRoomNumber

nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "That room doesn't exist!, BGETEQUAL Error: " & BtrieveErrorCode(nStatus)
    frmRoom.bLoaded = False: Call frmRoom.GotoRoom(Val(frmRoom.txtMap.Text), Val(frmRoom.txtRoom.Text), True)
    Exit Sub
End If
Call RoomRowToStruct(Roomdatabuf.buf)

'create exit back
Select Case nDir
    Case 0: NewExit = 1
    Case 1: NewExit = 0
    Case 2: NewExit = 3
    Case 3: NewExit = 2
    Case 4: NewExit = 7
    Case 5: NewExit = 6
    Case 6: NewExit = 5
    Case 7: NewExit = 4
    Case 8: NewExit = 9
    Case 9: NewExit = 8
End Select

'check to make sure the exit back on the selected room doesn't already exist
If Not Roomrec.RoomExit(NewExit) = 0 Then
    nYesNo = MsgBox("The exit back to the starting room from the chosen existing room is already in use.  Override?" _
        , vbQuestion + vbYesNo + vbDefaultButton1, "Exit in use ...")
    If Not nYesNo = vbYes Then GoTo skipupdate:
End If

Roomrec.RoomExit(NewExit) = RoomNum
If nNewMapNumber <> MapNum Then
    Roomrec.RoomType(NewExit) = 8
    Roomrec.Para1(NewExit) = MapNum
Else
    Roomrec.RoomType(NewExit) = 0
End If
    
'update the exisitng room
nStatus = UpdateRoom
If Not nStatus = 0 Then
    MsgBox "Error Saving room, BUPDATE: " & BtrieveErrorCode(nStatus)
    frmRoom.bLoaded = False: Call frmRoom.GotoRoom(Val(frmRoom.txtMap.Text), Val(frmRoom.txtRoom.Text), True)
    Exit Sub
End If

skipupdate:
'reget current room
RoomKeyStruct.MapNum = MapNum
RoomKeyStruct.RoomNum = RoomNum

nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error getting current room, BGETEQUAL Error: " & BtrieveErrorCode(nStatus)
    frmRoom.bLoaded = False: Call frmRoom.GotoRoom(Val(frmRoom.txtMap.Text), Val(frmRoom.txtRoom.Text), True)
    Exit Sub
End If
Call RoomRowToStruct(Roomdatabuf.buf)

'create the exit and check for map change
Roomrec.RoomExit(nDir) = nNewRoomNumber
If nNewMapNumber <> MapNum Then
    Roomrec.RoomType(nDir) = 8
    Roomrec.Para1(nDir) = nNewMapNumber
Else
    Roomrec.RoomType(nDir) = 0
End If

'update current room
nStatus = UpdateRoom
If Not nStatus = 0 Then
    MsgBox "Error Saving room, BUPDATE: " & BtrieveErrorCode(nStatus)
    frmRoom.bLoaded = False: Call frmRoom.GotoRoom(Val(frmRoom.txtMap.Text), Val(frmRoom.txtRoom.Text), True)
    Exit Sub
End If

'reget new room
RoomKeyStruct.MapNum = nNewMapNumber
RoomKeyStruct.RoomNum = nNewRoomNumber

nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Rooms updated, but error re-getting new room: " & BtrieveErrorCode(nStatus)
    frmRoom.bLoaded = False: Call frmRoom.GotoRoom(Val(frmRoom.txtMap.Text), Val(frmRoom.txtRoom.Text), True)
    Exit Sub
Else
    Call frmRoom.DispRoomInfo(Roomdatabuf.buf)
End If

Call StartMapping

Exit Sub
error:
Call HandleError
End Sub


Private Sub lblNumber_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Call lblCell_MouseDown(Index, Button, Shift, x, y)
End Sub


Private Sub txtIncrement_GotFocus()
EnableKeypad = False
cmdKeypad.Caption = "Enable Keypad"
Call SelectAll(txtIncrement)
End Sub

Private Sub txtIncrement_LostFocus()
EnableKeypad = True
cmdKeypad.Caption = "(Keypad Enabled)"
End Sub

Private Sub MapExits(Cell As Integer, Room As Long, Map As Long)
Dim ActivatedCell As Integer, nStatus As Integer, x As Integer, y As Integer
Dim rc As RECT, ToolTipString As String, sText As String, sMonsters As String
Dim sRemote As String, sCommand As String, nRemote As Integer, sRemoteEffect As String
Dim sExits As String

RoomKeyStruct.MapNum = Map
RoomKeyStruct.RoomNum = Room
nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    
    If Cell = CenterCell Then
        For x = 0 To 9
            cmdMove(x).BackColor = &HC0C0C0
        Next x
        If chkSmallForm.Value = 1 Then Exit Sub
    End If
    UnchartedCells(Cell) = 2
    lblCell(Cell).BackColor = &HFF&      'Call DrawOnRoom(lblCell(Cell), drSquare, 8, BrightRed)
    ToolTipString = "Map " & Map & " Room " & Room
    rc.Left = lblCell(Cell).Left
    rc.Top = lblCell(Cell).Top
    rc.Bottom = (lblCell(Cell).Top + lblCell(Cell).Height)
    rc.Right = (lblCell(Cell).Left + lblCell(Cell).Width)
    objTooltip.SetToolTipItem picMap.hwnd, 0, rc.Left, rc.Top, rc.Right, rc.Bottom, ToolTipString, False
    Exit Sub
End If
RoomRowToStruct Roomdatabuf.buf

If Cell = CenterCell Then
    For x = 0 To 9
        If Roomrec.RoomExit(x) > 0 Then
            cmdMove(x).BackColor = &HFF8080
        Else
            cmdMove(x).BackColor = &HC0C0C0
        End If
    Next x
    If chkSmallForm.Value = 1 Then Exit Sub
End If

CellRoom(Cell, 1) = Map
CellRoom(Cell, 2) = Room

If chkDisplayNumbers.Value = 1 Then lblNumber(Cell).Caption = Room

If chkMarkCMD.Value = 1 And Roomrec.CmdText > 0 Then
    sCommand = vbCrLf & vbCrLf & "Room commands: " & GetTextblockCMDS(Roomrec.CmdText)
    Call DrawOnRoom(lblCell(Cell), drSquare, 10, BrightGreen)
End If

If Roomrec.ControlRoom > 0 Then Call AddControlRoom(Roomrec.MapNumber, Roomrec.ControlRoom, Roomrec.RoomNumber)

'map exits
For x = 0 To 9
    If Not Roomrec.RoomExit(x) = 0 Then
        Select Case Roomrec.RoomType(x)
            Case 8: 'map change
                sExits = sExits & IIf(sExits = "", "", vbCrLf) _
                    & GetRoomExits(x, False) & ": " & "Map Change to map " & Roomrec.Para1(x)
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
                    
                    sRemote = sRemote & "[on room " & Roomrec.RoomExit(x) & ", " & sRemoteEffect & ": "
                    sRemote = sRemote & GetMessages(Roomrec.Para1(x), -1) & "]"
                    
                    Call DrawOnRoom(lblCell(Cell), drSquare, 10, BrightGreen)
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
    Call DrawOnRoom(lblCell(Cell), drCircle, 6, BrightMagenta)
End If
    
If chkMarkNPC.Value = 1 And Roomrec.PermNPC > 0 Then
    Call DrawOnRoom(lblCell(Cell), drOpenCircle, 3, BrightRed)
End If

lblCell(Cell).Caption = ""
'If chkNoColors.Value = 1 Then
'    lblCell(Cell).BackColor = nBackcolor '-- nothing
'Else
    'set color of this room
        If Roomrec.RoomExit(8) = 0 And Roomrec.RoomExit(9) = 0 Then
            lblCell(Cell).BackColor = nBackcolor '&H0& '-- nothing
        ElseIf Not Roomrec.RoomExit(8) = 0 And Roomrec.RoomExit(9) = 0 Then
            If chkDisplayNumbers.Value = 1 Then
                lblCell(Cell).BackColor = &H800000
            Else
                lblCell(Cell).BackColor = &HFF00& '-- up
            End If
            
            lblCell(Cell).Caption = "U"
        ElseIf Roomrec.RoomExit(8) = 0 And Not Roomrec.RoomExit(9) = 0 Then
            If chkDisplayNumbers.Value = 1 Then
                lblCell(Cell).BackColor = &H800000
            Else
                lblCell(Cell).BackColor = &HFFFF& '-- down
            End If
            lblCell(Cell).Caption = "D"
        Else
            If chkDisplayNumbers.Value = 1 Then
                lblCell(Cell).BackColor = &H800000
            Else
                lblCell(Cell).BackColor = &HFFFF00 '-- both
            End If
            lblCell(Cell).Caption = "B"
        End If
'End If

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
    
    'If chkMonsterRegen.Value = 1 Then
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
    'End If
    
    rc.Left = lblCell(Cell).Left
    rc.Top = lblCell(Cell).Top
    rc.Bottom = (lblCell(Cell).Top + lblCell(Cell).Height)
    rc.Right = (lblCell(Cell).Left + lblCell(Cell).Width)
    objTooltip.SetToolTipItem picMap.hwnd, Cell, rc.Left, rc.Top, rc.Right, rc.Bottom, ToolTipString, False
End If

'Call frmProgressBar.IncreaseProgress
'frmProgressBar.lblPanel(1).Caption = CellLabel(x)
'frmProgressBar.lblPanel(1).Caption = map & " / " & Room
        
UnchartedCells(Cell) = 2
End Sub
Private Function ActivateCell(FromCell As Integer, direction As Integer, ExitType As Integer) As Integer
Dim temp As Integer, LineColor As Long
'Dim x1 As Long, x2 As Long, y1 As Long, y2 As Long

'0 = N = -11
'1 = S = +11
'2 = E = +1
'3 = W = -1
'4 = NE = -40
'5 = NW = -42
'6 = SE = +42
'7 = SW = +40

'figure out which cell is to be activated
Select Case direction
    Case 0: 'north
        ActivateCell = (FromCell - 11)
        'checking to see if it's on the north edge
        If ActivateCell < 1 Then
            Call DrawOnRoom(lblCell(FromCell), drLineN, 8, Grey)
            GoTo DontActivate
        End If

    Case 1: 'south
        ActivateCell = (FromCell + 11)
        'checking to see if it's on the south edge
        If ActivateCell > SECorner Then
            Call DrawOnRoom(lblCell(FromCell), drLineS, 8, Grey)
            GoTo DontActivate
        End If

    Case 2: 'east
        ActivateCell = (FromCell + 1)
        'checking to see if it's on the east edge
        For temp = RowLength To SECorner Step 11
            If FromCell = temp Then
                Call DrawOnRoom(lblCell(FromCell), drLineE, 8, Grey)
                GoTo DontActivate
            End If
        Next
        
    Case 3: 'west
        ActivateCell = (FromCell - 1)
        'checking to see if it's on the west edge
        For temp = 1 To SECorner Step 11
            If FromCell = temp Then
                Call DrawOnRoom(lblCell(FromCell), drLineW, 8, Grey)
                GoTo DontActivate
            End If
        Next

    Case 4: 'northeast
        ActivateCell = (FromCell - 10)
        'checking to see if it's on the north edge
        If ActivateCell < 1 Then
            Call DrawOnRoom(lblCell(FromCell), drLineNE, 8, Grey)
            GoTo DontActivate
        End If
        'checking to see if it's on the east edge
        For temp = RowLength To SECorner Step 11
            If FromCell = temp Then
                Call DrawOnRoom(lblCell(FromCell), drLineNE, 8, Grey)
                GoTo DontActivate
            End If
        Next

    Case 5: 'northwest
        ActivateCell = (FromCell - 12)
        'checking to see if it's on the north edge
        If ActivateCell < 1 Then
            Call DrawOnRoom(lblCell(FromCell), drLineNW, 8, Grey)
            GoTo DontActivate:
        End If
        'checking to see if it's on the west edge
        For temp = 1 To SECorner Step 11
            If FromCell = temp Then
                Call DrawOnRoom(lblCell(FromCell), drLineNW, 8, Grey)
                GoTo DontActivate
            End If
        Next

    Case 6: 'southeast
        ActivateCell = (FromCell + 12)
        'checking to see if it's on the south edge
        If ActivateCell > SECorner Then
            Call DrawOnRoom(lblCell(FromCell), drLineSE, 8, Grey)
            GoTo DontActivate
        End If
        'checking to see if it's on the east edge
        For temp = RowLength To SECorner Step 11
            If FromCell = temp Then
                Call DrawOnRoom(lblCell(FromCell), drLineSE, 8, Grey)
                GoTo DontActivate
            End If
        Next

    Case 7: 'southwest
        ActivateCell = (FromCell + 10)
        'checking to see if it's on the south edge
        If ActivateCell > SECorner Then
            Call DrawOnRoom(lblCell(FromCell), drLineSW, 8, Grey)
            GoTo DontActivate:
        End If
        'checking to see if it's on the west edge
        For temp = 1 To SECorner Step 11
            If FromCell = temp Then
                Call DrawOnRoom(lblCell(FromCell), drLineSW, 8, Grey)
                GoTo DontActivate
            End If
        Next

    Case 8:
        'If chkNoColors.value = 1 Then GoTo DontActivate:
'        Select Case lblCell(FromCell).Tag
'            Case 1: 'up
'            Case 2: 'down
'                Call DrawOnRoom(lblCell(FromCell), drCircle, BrightCyan)
'                lblCell(FromCell).Tag = 3
'            Case 3: 'up and down
'            Case Else:
'                Call DrawOnRoom(lblCell(FromCell), drCircle, BrightGreen)
'                lblCell(FromCell).Tag = 1
'        End Select
'        If lblCell(FromCell).BackColor = &H0 Then lblCell(FromCell).BackColor = &HFF00&
'        If lblCell(FromCell).BackColor = &HFFFFFF Then lblCell(FromCell).BackColor = &HFF00&
'        If lblCell(FromCell).BackColor = &HFFFF& Then lblCell(FromCell).BackColor = &HFFFF00
        GoTo DontActivate:

    Case 9:
        'If chkNoColors.value = 1 Then GoTo DontActivate:
'        Select Case lblCell(FromCell).Tag
'            Case 1: 'up
'                Call DrawOnRoom(lblCell(FromCell), drCircle, BrightCyan)
'                lblCell(FromCell).Tag = 3
'            Case 2: 'down
'            Case 3: 'up and down
'            Case Else:
'                Call DrawOnRoom(lblCell(FromCell), drCircle, Yellow)
'                lblCell(FromCell).Tag = 2
'        End Select
        
'        If lblCell(FromCell).BackColor = &H0 Then lblCell(FromCell).BackColor = &HFFFF&
'        If lblCell(FromCell).BackColor = &HFFFFFF Then lblCell(FromCell).BackColor = &HFFFF&
'        If lblCell(FromCell).BackColor = &HFF00& Then lblCell(FromCell).BackColor = &HFFFF00
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
    Case Else: LineColor = 8 '0  'grey (was black) - anything else
End Select

'If chkNoColors.Value = 1 Or chkNoLineColors.Value = 1 Then
    'If chkUseWhiteBG.Value = 1 Then
    '    LineColor = black
    'Else
        'LineColor = Grey
    'End If
'End If

'draw the line
Select Case direction
    Case 0:
'        x1 = lblCell(FromCell).Left + 150
'        y1 = lblCell(FromCell).Top + 150
'        x2 = lblCell(ActivateCell).Left + 150
'        y2 = Abs((lblCell(ActivateCell).Top + 150 + lblCell(FromCell).Top + 150) / 2)
'        Line (x1, y1)-(x2, y2), QBColor(LineColor), BF
        Call DrawOnRoom(lblCell(FromCell), drLineN, 8, LineColor)
    Case 1:
'        x1 = lblCell(FromCell).Left + 150
'        y1 = lblCell(FromCell).Top + 150
'        x2 = lblCell(ActivateCell).Left + 150
'        y2 = Abs(((lblCell(ActivateCell).Top + 150 + lblCell(FromCell).Top + 150) / 2)) + 1
'        Line (x1, y1)-(x2, y2), QBColor(LineColor), BF
        Call DrawOnRoom(lblCell(FromCell), drLineS, 8, LineColor)
    Case 2:
'        x1 = lblCell(FromCell).Left + 150
'        y1 = lblCell(FromCell).Top + 150
'        x2 = Abs(((lblCell(ActivateCell).Left + 150 + lblCell(FromCell).Left + 150) / 2))
'        y2 = lblCell(ActivateCell).Top + 150
'        Line (x1, y1)-(x2, y2), QBColor(LineColor), BF
        Call DrawOnRoom(lblCell(FromCell), drLineE, 8, LineColor)
    Case 3:
'        x1 = lblCell(FromCell).Left + 150
'        y1 = lblCell(FromCell).Top + 150
'        x2 = Abs(((lblCell(ActivateCell).Left + 150 + lblCell(FromCell).Left + 150) / 2))
'        y2 = lblCell(ActivateCell).Top + 150
'        Line (x1, y1)-(x2, y2), QBColor(LineColor), BF
        Call DrawOnRoom(lblCell(FromCell), drLineW, 8, LineColor)
    Case 4:
'        x1 = lblCell(FromCell).Left + 150
'        y1 = lblCell(FromCell).Top + 150
'        x2 = Abs(((lblCell(ActivateCell).Left + 150 + lblCell(FromCell).Left + 150) / 2))
'        y2 = Abs(((lblCell(ActivateCell).Top + 150 + lblCell(FromCell).Top + 150) / 2))
'        Line (x1, y1)-(x2, y2), QBColor(LineColor)
        Call DrawOnRoom(lblCell(FromCell), drLineNE, 8, LineColor)
    Case 5:
'        x1 = lblCell(FromCell).Left + 150
'        y1 = lblCell(FromCell).Top + 150
'        x2 = Abs(((lblCell(ActivateCell).Left + 150 + lblCell(FromCell).Left + 150) / 2))
'        y2 = Abs(((lblCell(ActivateCell).Top + 150 + lblCell(FromCell).Top + 150) / 2))
'        Line (x1, y1)-(x2, y2), QBColor(LineColor)
        Call DrawOnRoom(lblCell(FromCell), drLineNW, 8, LineColor)
    Case 6:
'        x1 = lblCell(FromCell).Left + 150
'        y1 = lblCell(FromCell).Top + 150
'        x2 = Abs(((lblCell(ActivateCell).Left + 150 + lblCell(FromCell).Left + 150) / 2))
'        y2 = Abs(((lblCell(ActivateCell).Top + 150 + lblCell(FromCell).Top + 150) / 2))
'        Line (x1, y1)-(x2, y2), QBColor(LineColor)
        Call DrawOnRoom(lblCell(FromCell), drLineSE, 8, LineColor)
    Case 7:
'        x1 = lblCell(FromCell).Left + 150
'        y1 = lblCell(FromCell).Top + 150
'        x2 = Abs(((lblCell(ActivateCell).Left + 150 + lblCell(FromCell).Left + 150) / 2))
'        y2 = Abs(((lblCell(ActivateCell).Top + 150 + lblCell(FromCell).Top + 150) / 2))
'        Line (x1, y1)-(x2, y2), QBColor(LineColor)
        Call DrawOnRoom(lblCell(FromCell), drLineSW, 8, LineColor)
        
End Select

'if the cell to be activated has already been mapped, dont map it again
If UnchartedCells(ActivateCell) = 2 Then GoTo DontActivate:

Select Case ExitType
    Case 12: ActivateCell = -1 'if it's a remote action, dont map it
    Case 8: 'if it's a map change, check to see if it should be mapped
        'If chkFollowMapChanges.Value = 1 Then
            lblCell(ActivateCell).BackColor = &H0
        'Else
        '    ActivateCell = -1
        'End If
    Case Else: lblCell(ActivateCell).BackColor = &H0
End Select

Exit Function
DontActivate:
ActivateCell = -1

End Function

Private Sub lblCell_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nStatus As Integer

On Error GoTo error:

EnableKeypad = True
cmdKeypad.Caption = "(Keypad Enabled)"

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
    If frmRoom.Visible = False Then frmRoom.Show
    Call frmRoom.GotoRoom(CellRoom(Index, 1), CellRoom(Index, 2))
    Exit Sub
    
ElseIf Button = 2 Then
    If Shift = 1 Then CenterCell = Index
    
    If lblCell(Index).Caption = "U" Then '-- up
        PopupMenu frmMain.mnuMapEditorUp
    ElseIf lblCell(Index).Caption = "D" Then '-- down
        PopupMenu frmMain.mnuMapEditorDown
    ElseIf lblCell(Index).Caption = "B" Then '-- both
        PopupMenu frmMain.mnuMapEditorUpDown
    Else
        nStatus = GotoRoom(CellRoom(Index, 1), CellRoom(Index, 2))
        If Not nStatus = 0 Then
            MsgBox "BGETEQUAL, Room, Error: " & BtrieveErrorCode(nStatus)
            Exit Sub
        End If
            
        Call StartMapping
        Me.SetFocus
    End If
End If

out:
Exit Sub
error:
Call HandleError("lblCell_MouseDown")
Resume out:

End Sub

Private Sub DrawOnRoom(ByRef oLabel As label, ByVal drDrawType As DrawRoomEnum, ByVal nSize As Integer, ByVal nColor As QBColorCode)
Dim x1 As Integer, x2 As Integer, y1 As Integer, y2 As Integer
Dim nTemp As Integer, nLineColor As QBColorCode

nTemp = picMap.DrawWidth

'If chkNoColors.Value = 1 Then
'    If chkUseWhiteBG.Value = 1 Then
'        nColor = black
'    Else
'        nColor = Grey
'    End If
'End If

nLineColor = nColor
'If chkNoLineColors.Value = 1 Then
'    If chkUseWhiteBG.Value = 1 Then
'        nLineColor = black
'    Else
'        nLineColor = Grey
'    End If
'End If

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
        x1 = oLabel.Left - 12
        y1 = oLabel.Top + oLabel.Height + 12
        x2 = oLabel.Left + 12
        y2 = oLabel.Top - 12
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
        
        '\
        x1 = x2
        y1 = y2
        x2 = oLabel.Left + oLabel.Width + 12
        y2 = oLabel.Top + oLabel.Height + 12
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
        
        '\
        x1 = x2
        y1 = y2
        x2 = oLabel.Left - 12
        y2 = oLabel.Top
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
        
        '-
        x1 = x2
        y1 = y2
        x2 = oLabel.Left + oLabel.Width + 12
        y2 = y1
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
        
        '/
        x1 = x2
        y1 = y2
        x2 = oLabel.Left - 12
        y2 = oLabel.Top + oLabel.Height + 12
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
        
    Case 2: 'open circle
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 12
        y1 = oLabel.Top + 12
        picMap.Circle (x1, y1), 18, QBColor(nColor)
      
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
        x1 = oLabel.Left + 12
        y1 = oLabel.Top + 12
        picMap.Circle (x1, y1), 15, QBColor(nColor)
    
    Case 6: 'LineN
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 12
        y1 = oLabel.Top + 12
        x2 = x1
        y2 = y1 - 20
        picMap.Line (x1, y1)-(x2, y2), QBColor(nLineColor), BF
        
    Case 7: 'LineS
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 12
        y1 = oLabel.Top + 12
        x2 = x1
        y2 = y1 + 20
        picMap.Line (x1, y1)-(x2, y2), QBColor(nLineColor), BF
        
    Case 8: 'LineE
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 12
        y1 = oLabel.Top + 12
        x2 = x1 + 20
        y2 = y1
        picMap.Line (x1, y1)-(x2, y2), QBColor(nLineColor), BF
        
    Case 9: 'LineW
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 12
        y1 = oLabel.Top + 12
        x2 = x1 - 20
        y2 = y1
        picMap.Line (x1, y1)-(x2, y2), QBColor(nLineColor), BF
        
    Case 10: 'LineNE
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 12
        y1 = oLabel.Top + 12
        x2 = x1 + 19
        y2 = y1 - 19
        picMap.Line (x1, y1)-(x2, y2), QBColor(nLineColor)
        
    Case 11: 'LineNW
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 12
        y1 = oLabel.Top + 12
        x2 = x1 - 19
        y2 = y1 - 19
        picMap.Line (x1, y1)-(x2, y2), QBColor(nLineColor)
        
    Case 12: 'LineSE
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 12
        y1 = oLabel.Top + 12
        x2 = x1 + 20
        y2 = y1 + 20
        picMap.Line (x1, y1)-(x2, y2), QBColor(nLineColor)
    
    Case 13: 'LineSW
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 12
        y1 = oLabel.Top + 12
        x2 = x1 - 20
        y2 = y1 + 20
        picMap.Line (x1, y1)-(x2, y2), QBColor(nLineColor)
        
End Select

picMap.DrawWidth = nTemp
End Sub

Private Function GotoRoom(ByVal nMap As Long, ByVal nRoom As Long) As Integer

RoomKeyStruct.MapNum = nMap
RoomKeyStruct.RoomNum = nRoom
        
GotoRoom = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If GotoRoom = 0 Then
    Call frmRoom.DispRoomInfo(Roomdatabuf.buf)
End If

End Function
