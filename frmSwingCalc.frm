VERSION 5.00
Begin VB.Form frmSwingCalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Swing Calculator"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8745
   Icon            =   "frmSwingCalc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   8745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCopytoClip 
      Caption         =   "Copy Only True AVG"
      Height          =   375
      Index           =   2
      Left            =   4380
      TabIndex        =   36
      Top             =   4020
      Width           =   2055
   End
   Begin VB.CommandButton cmdCopytoClip 
      Caption         =   "Copy Only Swings"
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   35
      Top             =   4020
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      Caption         =   "True Average"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   6720
      TabIndex        =   73
      Top             =   60
      Width           =   1935
      Begin VB.TextBox txtTrueAVG 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   7
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   89
         Top             =   2700
         Width           =   1695
      End
      Begin VB.TextBox txtTrueAVG 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   6
         Left            =   120
         MaxLength       =   10
         TabIndex        =   87
         Top             =   2100
         Width           =   1695
      End
      Begin VB.TextBox txtTrueAVG 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   5
         Left            =   1020
         MaxLength       =   4
         TabIndex        =   85
         Top             =   1560
         Width           =   795
      End
      Begin VB.TextBox txtTrueAVG 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   120
         MaxLength       =   4
         TabIndex        =   83
         Top             =   1560
         Width           =   795
      End
      Begin VB.TextBox txtTrueAVG 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   1020
         MaxLength       =   4
         TabIndex        =   81
         Top             =   1020
         Width           =   795
      End
      Begin VB.TextBox txtTrueAVG 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   120
         MaxLength       =   4
         TabIndex        =   79
         Top             =   1020
         Width           =   795
      End
      Begin VB.TextBox txtTrueAVG 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   1020
         MaxLength       =   4
         TabIndex        =   77
         Top             =   480
         Width           =   795
      End
      Begin VB.TextBox txtTrueAVG 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   120
         MaxLength       =   4
         TabIndex        =   75
         Top             =   480
         Width           =   795
      End
      Begin VB.CommandButton cmdPasteMega 
         Caption         =   "Paste MegaMUD Stats"
         Height          =   555
         Left            =   120
         TabIndex        =   74
         Top             =   3180
         Width           =   1695
      End
      Begin VB.Label lblTrueAVG 
         Alignment       =   2  'Center
         Caption         =   "Average Round"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   90
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label lblTrueAVG 
         Alignment       =   2  'Center
         Caption         =   "Swings"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   88
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblTrueAVG 
         Alignment       =   2  'Center
         Caption         =   "Crit AVG"
         Height          =   195
         Index           =   5
         Left            =   1020
         TabIndex        =   86
         Top             =   1380
         Width           =   795
      End
      Begin VB.Label lblTrueAVG 
         Alignment       =   2  'Center
         Caption         =   "Crit %"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   84
         Top             =   1380
         Width           =   795
      End
      Begin VB.Label lblTrueAVG 
         Alignment       =   2  'Center
         Caption         =   "Extra AVG"
         Height          =   195
         Index           =   3
         Left            =   1020
         TabIndex        =   82
         Top             =   840
         Width           =   795
      End
      Begin VB.Label lblTrueAVG 
         Alignment       =   2  'Center
         Caption         =   "Extra %"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   80
         Top             =   840
         Width           =   795
      End
      Begin VB.Label lblTrueAVG 
         Alignment       =   2  'Center
         Caption         =   "Hit AVG"
         Height          =   195
         Index           =   1
         Left            =   1020
         TabIndex        =   78
         Top             =   300
         Width           =   795
      End
      Begin VB.Label lblTrueAVG 
         Alignment       =   2  'Center
         Caption         =   "Hit %"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   76
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   6720
      TabIndex        =   37
      Top             =   4020
      Width           =   1935
   End
   Begin VB.CommandButton cmdCopytoClip 
      Caption         =   "Cop&y to Clipboard"
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   34
      Top             =   4020
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Swings / EU Remaining"
      Height          =   1695
      Left            =   60
      TabIndex        =   38
      Top             =   2220
      Width           =   6615
      Begin VB.TextBox txtSwing 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   480
         Width           =   435
      End
      Begin VB.TextBox txtSwing 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   480
         Width           =   435
      End
      Begin VB.TextBox txtSwing 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   480
         Width           =   435
      End
      Begin VB.TextBox txtSwing 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   480
         Width           =   435
      End
      Begin VB.TextBox txtSwing 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   480
         Width           =   435
      End
      Begin VB.TextBox txtSwing 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   480
         Width           =   435
      End
      Begin VB.TextBox txtSwing 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   480
         Width           =   435
      End
      Begin VB.TextBox txtSwing 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   4740
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   480
         Width           =   435
      End
      Begin VB.TextBox txtSwing 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   480
         Width           =   435
      End
      Begin VB.TextBox txtSwing 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   6060
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   480
         Width           =   435
      End
      Begin VB.Label lblEncum 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Encum %"
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
         Left            =   5685
         TabIndex        =   72
         Top             =   1380
         Width           =   780
      End
      Begin VB.Label lblEnergy 
         AutoSize        =   -1  'True
         Caption         =   "Energy per swing"
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
         Left            =   180
         TabIndex        =   71
         Top             =   1080
         Width           =   1470
      End
      Begin VB.Label lblRawSwing 
         AutoSize        =   -1  'True
         Caption         =   "RawSwing"
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
         Left            =   180
         TabIndex        =   70
         Top             =   1380
         Width           =   900
      End
      Begin VB.Label lblQND 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "QnD Crits"
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
         Left            =   5640
         TabIndex        =   69
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label lblEU 
         Alignment       =   2  'Center
         Caption         =   "1000"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   9
         Left            =   6060
         TabIndex        =   68
         Top             =   780
         Width           =   435
      End
      Begin VB.Label lblEU 
         Alignment       =   2  'Center
         Caption         =   "1000"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   8
         Left            =   5400
         TabIndex        =   67
         Top             =   780
         Width           =   435
      End
      Begin VB.Label lblEU 
         Alignment       =   2  'Center
         Caption         =   "1000"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   7
         Left            =   4740
         TabIndex        =   66
         Top             =   780
         Width           =   435
      End
      Begin VB.Label lblEU 
         Alignment       =   2  'Center
         Caption         =   "1000"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   6
         Left            =   4080
         TabIndex        =   65
         Top             =   780
         Width           =   435
      End
      Begin VB.Label lblEU 
         Alignment       =   2  'Center
         Caption         =   "1000"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   5
         Left            =   3420
         TabIndex        =   64
         Top             =   780
         Width           =   435
      End
      Begin VB.Label lblEU 
         Alignment       =   2  'Center
         Caption         =   "1000"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   4
         Left            =   2760
         TabIndex        =   63
         Top             =   780
         Width           =   435
      End
      Begin VB.Label lblEU 
         Alignment       =   2  'Center
         Caption         =   "1000"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   3
         Left            =   2100
         TabIndex        =   62
         Top             =   780
         Width           =   435
      End
      Begin VB.Label lblEU 
         Alignment       =   2  'Center
         Caption         =   "1000"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   2
         Left            =   1440
         TabIndex        =   61
         Top             =   780
         Width           =   435
      End
      Begin VB.Label lblEU 
         Alignment       =   2  'Center
         Caption         =   "1000"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   1
         Left            =   780
         TabIndex        =   60
         Top             =   780
         Width           =   435
      End
      Begin VB.Label lblEU 
         Alignment       =   2  'Center
         Caption         =   "1000"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   59
         Top             =   780
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   1
         Left            =   780
         TabIndex        =   57
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   2
         Left            =   1440
         TabIndex        =   56
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   3
         Left            =   2100
         TabIndex        =   55
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   4
         Left            =   2760
         TabIndex        =   54
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   5
         Left            =   3420
         TabIndex        =   53
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   6
         Left            =   4080
         TabIndex        =   52
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   7
         Left            =   4740
         TabIndex        =   51
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   8
         Left            =   5400
         TabIndex        =   50
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   9
         Left            =   6060
         TabIndex        =   49
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   2115
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4575
      Begin VB.CommandButton cmdAlterLevel 
         Caption         =   "+"
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
         Index           =   3
         Left            =   1860
         TabIndex        =   8
         Top             =   780
         Width           =   315
      End
      Begin VB.CommandButton cmdAlterLevel 
         Caption         =   "+"
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
         Index           =   7
         Left            =   4140
         TabIndex        =   16
         Top             =   780
         Width           =   315
      End
      Begin VB.CommandButton cmdAlterLevel 
         Caption         =   "-"
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
         Index           =   6
         Left            =   3840
         TabIndex        =   15
         Top             =   780
         Width           =   315
      End
      Begin VB.CommandButton cmdAlterLevel 
         Caption         =   "+"
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
         Index           =   9
         Left            =   4140
         TabIndex        =   20
         Top             =   1200
         Width           =   315
      End
      Begin VB.CommandButton cmdAlterLevel 
         Caption         =   "+"
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
         Index           =   5
         Left            =   4140
         TabIndex        =   12
         Top             =   360
         Width           =   315
      End
      Begin VB.CommandButton cmdAlterLevel 
         Caption         =   "-"
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
         Index           =   8
         Left            =   3840
         TabIndex        =   19
         Top             =   1200
         Width           =   315
      End
      Begin VB.CommandButton cmdAlterLevel 
         Caption         =   "-"
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
         Index           =   4
         Left            =   3840
         TabIndex        =   11
         Top             =   360
         Width           =   315
      End
      Begin VB.CommandButton cmdAlterLevel 
         Caption         =   "-"
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
         Index           =   2
         Left            =   1560
         TabIndex        =   7
         Top             =   780
         Width           =   315
      End
      Begin VB.CommandButton cmdAlterLevel 
         Caption         =   "+"
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
         Left            =   1860
         TabIndex        =   4
         Top             =   360
         Width           =   315
      End
      Begin VB.CommandButton cmdAlterLevel 
         Caption         =   "-"
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
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   315
      End
      Begin VB.ComboBox cmbCombat 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdJump 
         Caption         =   ">"
         Height          =   315
         Left            =   4260
         TabIndex        =   25
         Top             =   1680
         Width           =   195
      End
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   2
         Text            =   "60"
         Top             =   360
         Width           =   675
      End
      Begin VB.TextBox txtAgility 
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   6
         Text            =   "100"
         Top             =   780
         Width           =   675
      End
      Begin VB.TextBox txtEncum 
         Height          =   285
         Left            =   3120
         MaxLength       =   5
         TabIndex        =   14
         Text            =   "1000"
         Top             =   780
         Width           =   675
      End
      Begin VB.TextBox txtMaxEncum 
         Height          =   285
         Left            =   3120
         MaxLength       =   6
         TabIndex        =   18
         Text            =   "4800"
         Top             =   1200
         Width           =   675
      End
      Begin VB.ComboBox cmbWeapon 
         Height          =   315
         ItemData        =   "frmSwingCalc.frx":08CA
         Left            =   840
         List            =   "frmSwingCalc.frx":08CC
         Sorted          =   -1  'True
         TabIndex        =   24
         Text            =   "cmbWeapon"
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox txtStrength 
         Height          =   285
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   10
         Text            =   "100"
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Level"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Combat"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Agility"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   780
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Encum."
         Height          =   195
         Index           =   3
         Left            =   2280
         TabIndex        =   13
         Top             =   780
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Max Enc."
         Height          =   195
         Index           =   4
         Left            =   2280
         TabIndex        =   17
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Weapon"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Strength"
         Height          =   195
         Index           =   6
         Left            =   2280
         TabIndex        =   9
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Speed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   4740
      TabIndex        =   26
      Top             =   60
      Width           =   1935
      Begin VB.CheckBox chkBashing 
         Caption         =   "Bashing"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Top             =   1800
         Width           =   1515
      End
      Begin VB.OptionButton optSpeed 
         Caption         =   "Custom"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   30
         Top             =   1140
         Width           =   915
      End
      Begin VB.TextBox txtSpeed 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   31
         Text            =   "100"
         Top             =   1110
         Width           =   615
      End
      Begin VB.CheckBox chkSlowness 
         Caption         =   "Slowness Ability"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   1500
         Width           =   1515
      End
      Begin VB.OptionButton optSpeed 
         Caption         =   "Fast (85)"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   27
         Top             =   240
         Width           =   1395
      End
      Begin VB.OptionButton optSpeed 
         Caption         =   "Normal (100)"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   28
         Top             =   540
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.OptionButton optSpeed 
         Caption         =   "Slow (125)"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   29
         Top             =   840
         Width           =   1395
      End
   End
   Begin VB.Timer timMouseDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   60
   End
End
Attribute VB_Name = "frmSwingCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'Private Const WM_SETREDRAW = &HB
'Private msOldString As String ' module level global
'Private miStart As Integer    ' module level global
'Private miLength As Integer   ' module level global
Dim bMouseDown As Boolean





Private Sub chkBashing_Click()
Call CalcSwings
End Sub

Private Sub chkSlowness_Click()
Call CalcSwings
End Sub

Private Sub cmbCombat_Click()
Call CalcSwings
End Sub

Private Sub cmbWeapon_Click()
Call CalcSwings
End Sub

Private Sub cmdAlterLevel_Click(Index As Integer)
    
If Not bMouseDown Then Call AlterLevel(Index)

End Sub

Private Sub AlterLevel(ByVal Index As Integer)

On Error GoTo error:

If Index = 0 Then 'minus LEVEL
        If Val(txtLevel.Text) <= 0 Then
            txtLevel.Text = 0
        Else
            txtLevel.Text = Val(txtLevel.Text) - 1
        End If
    ElseIf Index = 1 Then 'plus
        If Val(txtLevel.Text) >= 9999 Then
            txtLevel.Text = 9999
        Else
            txtLevel.Text = Val(txtLevel.Text) + 1
        End If
    ElseIf Index = 2 Then 'minus AGL
        If Val(txtAgility.Text) <= 0 Then
            txtAgility.Text = 0
        Else
            txtAgility.Text = Val(txtAgility.Text) - 1
        End If
    ElseIf Index = 3 Then 'plus
        If Val(txtAgility.Text) >= 9999 Then
            txtAgility.Text = 9999
        Else
            txtAgility.Text = Val(txtAgility.Text) + 1
        End If
    ElseIf Index = 4 Then 'minus STR
        If Val(txtStrength.Text) <= 0 Then
            txtStrength.Text = 0
        Else
            txtStrength.Text = Val(txtStrength.Text) - 1
        End If
    ElseIf Index = 5 Then 'plus
        If Val(txtStrength.Text) >= 9999 Then
            txtStrength.Text = 9999
        Else
            txtStrength.Text = Val(txtStrength.Text) + 1
        End If
    ElseIf Index = 6 Then 'minus ENC
        If Val(txtEncum.Text) <= 0 Then
            txtEncum.Text = 0
        Else
            txtEncum.Text = Val(txtEncum.Text) - 25
        End If
    ElseIf Index = 7 Then 'plus
        If Val(txtEncum.Text) >= 99999 Then
            txtEncum.Text = 99999
        Else
            txtEncum.Text = Val(txtEncum.Text) + 25
        End If
    ElseIf Index = 8 Then 'minus MAX ENC
        If Val(txtMaxEncum.Text) <= 0 Then
            txtMaxEncum.Text = 0
        Else
            txtMaxEncum.Text = Val(txtMaxEncum.Text) - 1
        End If
    ElseIf Index = 9 Then 'plus
        If Val(txtMaxEncum.Text) >= 99999 Then
            txtMaxEncum.Text = 99999
        Else
            txtMaxEncum.Text = Val(txtMaxEncum.Text) + 1
        End If
    End If
    Call CalcSwings

Exit Sub

error:
Call HandleError("AlterLevel")
    
End Sub
Private Sub cmdAlterLevel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

bMouseDown = True

Do While bMouseDown
    timMouseDown.Enabled = True
    Call AlterLevel(Index)
    Do While timMouseDown.Enabled
        DoEvents
    Loop
Loop

'bMouseDown = True
'
'Do While bMouseDown
'    timMouseDown.Enabled = True
'    If Index = 0 Then 'minus LEVEL
'        If Val(txtLevel.Text) <= 0 Then
'            txtLevel.Text = 0
'        Else
'            txtLevel.Text = Val(txtLevel.Text) - 1
'        End If
'    ElseIf Index = 1 Then 'plus
'        If Val(txtLevel.Text) >= 9999 Then
'            txtLevel.Text = 9999
'        Else
'            txtLevel.Text = Val(txtLevel.Text) + 1
'        End If
'    ElseIf Index = 2 Then 'minus AGL
'        If Val(txtAgility.Text) <= 0 Then
'            txtAgility.Text = 0
'        Else
'            txtAgility.Text = Val(txtAgility.Text) - 1
'        End If
'    ElseIf Index = 3 Then 'plus
'        If Val(txtAgility.Text) >= 9999 Then
'            txtAgility.Text = 9999
'        Else
'            txtAgility.Text = Val(txtAgility.Text) + 1
'        End If
'    ElseIf Index = 4 Then 'minus STR
'        If Val(txtStrength.Text) <= 0 Then
'            txtStrength.Text = 0
'        Else
'            txtStrength.Text = Val(txtStrength.Text) - 1
'        End If
'    ElseIf Index = 5 Then 'plus
'        If Val(txtStrength.Text) >= 9999 Then
'            txtStrength.Text = 9999
'        Else
'            txtStrength.Text = Val(txtStrength.Text) + 1
'        End If
'    ElseIf Index = 6 Then 'minus ENC
'        If Val(txtEncum.Text) <= 0 Then
'            txtEncum.Text = 0
'        Else
'            txtEncum.Text = Val(txtEncum.Text) - 25
'        End If
'    ElseIf Index = 7 Then 'plus
'        If Val(txtEncum.Text) >= 99999 Then
'            txtEncum.Text = 99999
'        Else
'            txtEncum.Text = Val(txtEncum.Text) + 25
'        End If
'    ElseIf Index = 8 Then 'minus MAX ENC
'        If Val(txtMaxEncum.Text) <= 0 Then
'            txtMaxEncum.Text = 0
'        Else
'            txtMaxEncum.Text = Val(txtMaxEncum.Text) - 1
'        End If
'    ElseIf Index = 9 Then 'plus
'        If Val(txtMaxEncum.Text) >= 99999 Then
'            txtMaxEncum.Text = 99999
'        Else
'            txtMaxEncum.Text = Val(txtMaxEncum.Text) + 1
'        End If
'    End If
'    Call CalcSwings
'    Do While timMouseDown.Enabled
'        DoEvents
'    Loop
'Loop

End Sub

Private Sub cmdAlterLevel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
bMouseDown = False
DoEvents
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdCopytoClip_Click(Index As Integer)
Dim str As String, x As Integer, sTrue As String

On Error GoTo error:

str = "Swings: "
For x = 0 To 9
    str = str & txtSwing(x).Text
    If x < 9 Then str = str & "/"
Next x

If Index = 1 Then 'swings only
    str = str & " (Raw: " & lblRawSwing.Tag & ")"
    GoTo copytoclip:
End If

If Val(txtTrueAVG(7).Text) > 0 Then
    For x = 0 To 6
        If Not sTrue = "" Then sTrue = sTrue & ", "
        sTrue = sTrue & lblTrueAVG(x).Caption & "-" & Val(txtTrueAVG(x).Text)
    Next x
    sTrue = "True Average: " & Val(txtTrueAVG(7).Text) & " (" & sTrue & ")"
End If

If Index = 2 Then 'true only
    If sTrue = "" Then Exit Sub
    str = sTrue
    GoTo copytoclip:
End If

str = "Weapon: " & cmbWeapon.Text & vbCrLf & str

str = str & vbCrLf & "Energy Remaining: "
For x = 0 To 9
    str = str & lblEU(x).Caption
    If x < 9 Then str = str & "/"
Next x

str = str & vbCrLf & lblEnergy.Caption & ", "
str = str & lblRawSwing.Caption & ", "
str = str & lblQND.Caption

str = str & vbCrLf & "Combat: " & cmbCombat.Text & ", "
str = str & "Level: " & txtLevel.Text & ", "
str = str & "Agility: " & txtAgility.Text & ", "
str = str & "Strength: " & txtStrength & ", "
str = str & "Encumbrance: " & txtEncum & "/" & txtMaxEncum & " (" & lblEncum.Tag & "%)" & ", "

For x = 0 To 3
    If optSpeed(x) = True Then
        Select Case x
            Case 0: 'fast
                str = str & "Speed: Fast (85)"
            Case 1: 'normal
                str = str & "Speed: Normal (100)"
            Case 2: 'slow
                str = str & "Speed: Slow (125)"
            Case 3: 'custom
                str = str & "Speed: Custom (" & txtSpeed.Text & ")"
        End Select
    End If
Next x
If chkSlowness.Value = 1 Then str = str & " +Slowness Ability"
If chkBashing.Value = 1 Then str = str & " (Bashing)"
If Not sTrue = "" Then str = str & vbCrLf & sTrue

copytoclip:

If Not str = "" Then
    Clipboard.clear
    Clipboard.SetText str
End If

Exit Sub

error:
Call HandleError("cmdCopytoClip_Click")
End Sub

Private Sub cmdJump_Click()
If cmbWeapon.ListIndex < 0 Then Exit Sub
Call frmItem.GotoItem(cmbWeapon.ItemData(cmbWeapon.ListIndex))
End Sub

Private Sub cmdPasteMega_Click()
Dim nHitP As Double, nHitA As Long, nCritP As Double, nCritA As Long
Dim nExtraP As Double, nExtraA As Long
Dim x As Long, sClipText As String

On Error GoTo error:

sClipText = Clipboard.GetText
If sClipText = "" Then GoTo notext:

'HITS
x = InStr(1, sClipText, "Hit:")
If x = 0 Then GoTo notext:
x = x + 7 '7=len("Hit:   ")

If InStr(x, sClipText, "%") = 0 Then GoTo notext:

nHitP = Val(Mid(sClipText, x, InStr(x, sClipText, "%") - x))
If nHitP = 0 Then GoTo notext:

x = InStr(x, sClipText, "Avg:")
If x = 0 Then GoTo notext:
x = x + 4 '4=len("Avg:")

nHitA = Val(Mid(sClipText, x, InStr(x, sClipText, "Extra") - x))
If nHitA = 0 Then GoTo notext:

'EXTRA
x = InStr(1, sClipText, "Extra:")
If x = 0 Then GoTo Crit:
x = x + 7 '7=len("Extra: ")

nExtraP = Val(Mid(sClipText, x, InStr(x, sClipText, "%") - x))
If nExtraP = 0 Then GoTo Crit:

x = InStr(x, sClipText, "Avg:")
If x = 0 Then GoTo Crit:
x = x + 4 '4=len("Avg:")

nExtraA = Val(Mid(sClipText, x, InStr(x, sClipText, "Crit") - x))
If nExtraA = 0 Then GoTo Crit:

Crit:
'CRIT
x = InStr(1, sClipText, "Crit:")
If x = 0 Then GoTo notext:
x = x + 7 '7=len("Crit:  ")

nCritP = Val(Mid(sClipText, x, InStr(x, sClipText, "%") - x))
If nCritP = 0 Then GoTo calc:

x = InStr(x, sClipText, "Avg:")
If x = 0 Then GoTo calc:
x = x + 4 '4=len("Avg:")

nCritA = Val(Mid(sClipText, x, InStr(x, sClipText, "BS:") - x))
If nCritA = 0 Then GoTo calc:

calc:
txtTrueAVG(0).Text = nHitP
txtTrueAVG(1).Text = nHitA
txtTrueAVG(2).Text = nExtraP
txtTrueAVG(3).Text = nExtraA
txtTrueAVG(4).Text = nCritP
txtTrueAVG(5).Text = nCritA

Exit Sub

notext:
MsgBox "Incomplete/Missing MegaMUD Statistics in Clipboard", vbInformation

Exit Sub
error:
Call HandleError("cmdPasteMega_Click")

End Sub

Private Sub Form_Load()
On Error GoTo error:
Dim x As Integer, nCombat As Integer

cmbCombat.clear
cmbCombat.AddItem "1 (Poor)"
cmbCombat.ItemData(cmbCombat.NewIndex) = 1
cmbCombat.AddItem "2 (Fair)"
cmbCombat.ItemData(cmbCombat.NewIndex) = 2
cmbCombat.AddItem "3 (Average)"
cmbCombat.ItemData(cmbCombat.NewIndex) = 3
cmbCombat.AddItem "4 (Good)"
cmbCombat.ItemData(cmbCombat.NewIndex) = 4
cmbCombat.AddItem "5 (Excellent)"
cmbCombat.ItemData(cmbCombat.NewIndex) = 5

'If frmMain.cmbGlobalClass(0).ListIndex > 0 Then
'    If frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex) > 0 Then
'        nCombat = GetClassCombat(frmMain.cmbGlobalClass(0).ItemData(frmMain.cmbGlobalClass(0).ListIndex))
'        For x = 0 To 4
'            If cmbCombat.ItemData(x) = nCombat Then
'                cmbCombat.ListIndex = x
'                Exit For
'            End If
'        Next x
'    End If
'End If
If cmbCombat.ListIndex < 0 Then cmbCombat.ListIndex = 2

Call LoadWeapons

'If Val(frmMain.txtCharStats(3).Text) > 0 Then
'    txtAgility.Text = Val(frmMain.txtCharStats(3).Text)
'End If
'
'If Val(frmMain.txtGlobalLevel(0).Text) > 0 Then
'    txtLevel.Text = Val(frmMain.txtGlobalLevel(0).Text)
'End If
'
'If Val(frmMain.txtCharStats(0).Text) > 0 Then
'    txtStrength.Text = Val(frmMain.txtCharStats(0).Text)
'End If
'
'If Val(frmMain.txtStat(0).Text) > 0 Then
'    txtEncum.Text = Val(frmMain.txtStat(0).Text)
'End If
'
'If Val(frmMain.txtStat(1).Text) > 0 Then
'    txtMaxEncum.Text = Val(frmMain.txtStat(1).Text)
'End If
'
'If frmMain.cmbEquip(16).ListIndex >= 0 Then
'    Call GotoWeapon(frmMain.cmbEquip(16).ItemData(frmMain.cmbEquip(16).ListIndex))
'End If

Exit Sub
error:
Call HandleError
Resume Next
End Sub

Public Sub GotoWeapon(ByVal nItem As Long)
Dim x As Integer

For x = 0 To cmbWeapon.ListCount - 1
    If cmbWeapon.ItemData(x) = nItem Then
        cmbWeapon.ListIndex = x
        Exit For
    End If
Next x

End Sub
Private Sub cmbWeapon_KeyUp(KeyCode As Integer, Shift As Integer)

Dim x As Integer, sText As String

If KeyCode = Asc(vbTab) Or KeyCode = vbKeyShift Or _
     KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or _
     KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Then Exit Sub
     
If cmbWeapon.ListCount = 0 Then Exit Sub

sText = cmbWeapon.Text

For x = 0 To cmbWeapon.ListCount - 1
    If LCase(Left(cmbWeapon.List(x), Len(sText))) = LCase(sText) Then
        cmbWeapon.ListIndex = x
        cmbWeapon.SelStart = Len(sText)
        cmbWeapon.SelLength = Len(cmbWeapon.Text) - Len(sText)
        Exit Sub
    End If
    'Debug.Print LCase(Left(cmbWeapon.List(x), Len(sText)))
Next x

'Dim sComboText As String
'Dim iLoop As Integer
'Dim sTempString As String
'Dim lReturn As Long
'Dim bInList As Boolean
'Dim sItem
'
'If Not KeyCode = Asc(vbTab) And Not KeyCode = vbKeyShift And _
'    Not KeyCode = vbKeyLeft And Not KeyCode = vbKeyRight And _
'    Not KeyCode = vbKeyHome And Not KeyCode = vbKeyEnd Then
'
'    bInList = False
'
'    With cmbWeapon
'        sTempString = .Text
'        If Len(sTempString) = 1 Then sComboText = sTempString
'        lReturn = SendMessage(.hwnd, WM_SETREDRAW, False, 0&)
'        For iLoop = 0 To (.ListCount - 1)
'            sItem = .List(iLoop)
'            If UCase((sTempString & Mid$(sItem, _
'                Len(sTempString) + 1))) = UCase(sItem) Then
'
'                .ListIndex = iLoop
'                .Text = sItem
'                msOldString = sItem
'                miStart = Len(sTempString)
'                .SelStart = miStart
'                miLength = Len(sItem) - (Len(sTempString))
'                .SelLength = miLength
'                sComboText = sComboText & Mid$(sTempString, _
'                    Len(sComboText) + 1)
'                bInList = True
'                Exit For
'            End If
'        Next iLoop
'
'        If Not bInList Then
'            .Text = msOldString
'            .SelStart = miStart
'            .SelLength = miLength
'        End If
'
'        lReturn = SendMessage(.hwnd, WM_SETREDRAW, True, 0&)
'    End With
'End If

End Sub

Private Sub CalcSwings()
Dim nWeaponSpeed As Currency, nEnergy As Currency, nQnDBonus As Currency
Dim nTemp As Integer, nSpeed As Integer, x As Integer, i As Integer ', k As Integer, J As Integer
Dim nEncum As Currency, nSwings As Double
Dim nItem As Long, nStatus As Integer
'If tabItems.RecordCount = 0 Then Exit Sub
If cmbWeapon.ListIndex < 0 Then Exit Sub

'tabItems.Index = "pkItems"
'tabItems.Seek "=", cmbWeapon.ItemData(cmbWeapon.ListIndex)
'If tabItems.NoMatch Then Exit Sub
nItem = cmbWeapon.ItemData(cmbWeapon.ListIndex)
nStatus = BTRCALL(BGETEQUAL, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), nItem, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error getting first item: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

Call ItemRowToStruct(Itemdatabuf.buf)

'If the player has the "Slowness" flag on them (which is different, I believe,
'from the "Slowness" ability), then AdjustSpeedForSlowness is applied to the weapons Speed.
nWeaponSpeed = Itemrec.Speed 'tabItems.Fields("Speed")
If chkSlowness.Value = 1 Then nWeaponSpeed = AdjustSpeedForSlowness(nWeaponSpeed)

'The Speed value of the weapon (either adjusted as above, or as read raw from the item) is
'passed to CalcEnergyUsedWithEncum (using a previously calculated encumbrance percentage
'using CalcEncumbrancePercent).
'
'CalcEnergyUsedWithEncum handles both the heavy in hands scenario and the issue of
'encumbrance percentage affecting the actual EU per swing (and in the correct order).
nEncum = CalcEncumbrancePercent(Val(txtEncum.Text), Val(txtMaxEncum.Text))

nEnergy = CalcEnergyUsedWithEncum(cmbCombat.ItemData(cmbCombat.ListIndex), Val(txtLevel.Text), nWeaponSpeed, _
    Val(txtAgility.Text), Val(txtStrength.Text), nEncum, Itemrec.ReqStr) 'tabItems.Fields("Req Str"))

'After this, the calculated EU is passed to AdjustEnergyUsedWithSpeed.
'So on sped, you'd pass in 85, for slow you'd pass in 125
If optSpeed(0).Value = True Then 'speed
    nSpeed = 85
ElseIf optSpeed(1).Value = True Then 'normal
    nSpeed = 100
ElseIf optSpeed(2).Value = True Then 'slow
    nSpeed = 125
ElseIf optSpeed(3).Value = True Then 'custom
    nSpeed = Val(txtSpeed.Text)
    If nSpeed <= 0 Then
        txtSpeed.Text = 1
        nSpeed = 1
    End If
Else
    nSpeed = 100
End If

nEnergy = AdjustEnergyUsedWithSpeed(nEnergy, nSpeed)

If chkBashing.Value = 1 Then nEnergy = nEnergy * 2
'Finally, if the weapon is not "heavy in hands", the final EU, AGL, and encumbrance
'percent can be passed to CalcQuickAndDeadlyBonus to get the Q&D bonus.  And that should do it.

If Not Val(txtStrength.Text) < Itemrec.ReqStr Then 'tabItems.Fields("Req Str") Then
    nQnDBonus = CalcQuickAndDeadlyBonus(Val(txtAgility.Text), nEnergy, nEncum)
End If
If nQnDBonus > 0 Then
    lblQND.Caption = "QND Crits: " & nQnDBonus
Else
    lblQND.Caption = "QND Crits: None"
End If

'i've also included some helper functions that can be used outside of the usual flow of code
'to calculate weapon EU, for example AdjustEnergyUsedWithEncum, IsQuickAndDeadly, and CalcEncumbrance.

If nEnergy = 0 Then
    lblEnergy.Caption = "Energy per swing: 0"
    lblEnergy.Caption = "Energy per swing: " & nEnergy
    lblEncum.Caption = "Encumbrance: " & nEncum & "%"
    lblEncum.Tag = nEncum
    lblRawSwing.Caption = "Raw swing: " & 0
    lblRawSwing.Tag = 0
    txtTrueAVG(6).Text = 0

    For x = 0 To 9
        lblEU(x).Caption = 0
        txtSwing(x).Text = 0
    Next x
    
    Exit Sub
End If

'SWING CALC
nSwings = Round((1000 / nEnergy), 4)

lblEnergy.Caption = "Energy per swing: " & nEnergy
lblEncum.Caption = "Encumbrance: " & nEncum & "%"
lblEncum.Tag = nEncum
lblRawSwing.Caption = "Raw swing: " & nSwings
lblRawSwing.Tag = nSwings

txtTrueAVG(6).Text = nSwings

'txtTrueAVG(7).Text = CalcTrueAverage(nSwings, Val(txtTrueAVG(0).Text), Val(txtTrueAVG(1).Text), _
    Val(txtTrueAVG(2).Text), Val(txtTrueAVG(3).Text), Val(txtTrueAVG(4).Text), Val(txtTrueAVG(5).Text))

'Temp := 1000;
nTemp = 1000
If nEnergy <= 0 Then nEnergy = 1

'   Cnt := 1;
'   repeat
For x = 0 To 9

'     I := Temp div EU;
    i = Fix(nTemp \ nEnergy)

'     Temp := (Temp mod EU) + 1000;
    nTemp = (nTemp Mod nEnergy) + 1000
    
'     If (i > 5) Then
'       I := 5;
    If (i > 5) Then i = 5
    
'     K := 0;
'     while (K < I) do begin
'       J := Random(99);
'       If (J < Crits) Then
'         { Critical hit }
'         LWriteLn('|12' + Format(R, [PickHitMsg(@Msg), Random((MaxHit * 4) - (MaxHit * 2)) + (MaxHit * 2)]))
'       Else
'         { Normal hit }
'         LWriteLn('|12' + Format(S, [PickHitMsg(@Msg), Random(MaxHit - MinHit) + MinHit]));
'       ProcessSpell(Weapon);
'       Inc(K);
'     end;

    lblEU(x).Caption = nTemp - 1000
'     LWrite(Format('|08[|03EU REMAINING|08=|11%-4.4s|08/|03SWINGS|08=|11%d|08/|03ROUND|08=|11%d|08]:|15', [IntToStr(Temp - 1000), K, Cnt]));
'     case UpCase(ReadKey) of
'     'Q', 'X', #27: Break;
'     end;
'     WriteLn;
'     Inc(Cnt);
    txtSwing(x).Text = i
'   until (False);
Next x

'If cmbWeapon.ListIndex < cmbWeapon.ListCount - 1 Then cmbWeapon.ListIndex = cmbWeapon.ListIndex + 1

End Sub

Public Sub LoadWeapons()
On Error GoTo error:
'If tabItems.RecordCount = 0 Then Exit Sub
Dim nStatus As Integer
nStatus = BTRCALL(BGETFIRST, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error getting first item: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

Me.MousePointer = vbHourglass
'tabItems.MoveFirst
DoEvents

cmbWeapon.clear

'Do Until tabItems.EOF
'    If bOnlyInGame And tabItems.Fields("In Game") = 0 Then GoTo Skip:
'    If tabItems.Fields("Type") = 1 Then
'        cmbWeapon.AddItem (tabItems.Fields("Name") & " (" & tabItems.Fields("Number") & ")")
'        cmbWeapon.ItemData(cmbWeapon.NewIndex) = tabItems.Fields("Number")
'    End If
'Skip:
'    tabItems.MoveNext
'Loop
Do While nStatus = 0
    Call ItemRowToStruct(Itemdatabuf.buf)
    
    If Itemrec.Type = 1 Then
        cmbWeapon.AddItem (ClipNull(Itemrec.Name) & " (" & Itemrec.Number & ")")
        cmbWeapon.ItemData(cmbWeapon.NewIndex) = Itemrec.Number
    End If

    nStatus = BTRCALL(BGETNEXT, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), ByVal ItemKeyBuffer, KEY_BUF_LEN, 0)
Loop

If cmbWeapon.ListCount > 0 Then cmbWeapon.ListIndex = 0

Me.MousePointer = vbDefault
Exit Sub
error:
Call HandleError("SwingCalc_LoadItems")
Me.MousePointer = vbDefault

End Sub
Private Function CalcEncum(ByVal nSTR As Currency) As Currency
'{ Calculates Encumbrance for a given Strength } function  CalcEncumbrance(STR: integer): integer; begin
'Result := STR * 48;
'If (STR > 100) Then
'    Result := Result + ((STR - 100) * 36); end;

CalcEncum = nSTR * 48

If (nSTR > 100) Then
    CalcEncum = CalcEncum + ((nSTR - 100) * 36)
End If

CalcEncum = Fix(CalcEncum)
End Function

Private Function CalcEncumbrancePercent(ByVal nCurrent As Currency, ByVal nMax As Currency) As Currency
'{ Calculates the encumbrance percentage used for calculating Q&D bonuses and
'  energy used }
'function  CalcEncumbrancePercent(Current, Maximum: integer): integer; begin
'  Result := (Current * 100) div Maximum; end;

If nMax <= 0 Then nMax = 1

CalcEncumbrancePercent = Fix((nCurrent * 100) / nMax)

End Function

Private Function AdjustSpeedForSlowness(ByVal nSpeed As Currency) As Currency
'{ Adjusts the Speed of a weapon for the case where a player has the Slowness
'  flag on them }
'function  AdjustSpeedForSlowness(Speed: integer): integer; begin
'  Result := (Speed * 3) div 2;
'end;

AdjustSpeedForSlowness = Fix((nSpeed * 3) / 2)

End Function

Private Function CalcEnergyUsed(ByVal nCombat As Currency, ByVal nLevel As Currency, _
    ByVal nSpeed As Currency, ByVal nAGL As Currency, Optional ByVal nSTR As Currency = 0, _
    Optional ByVal nItemSTR As Currency = 0) As Currency
'{ Calculates the energy used for a given Combat rating, Level, Speed, AGL, STR,
'  and ItemSTR }
'function  CalcEnergyUsed(Combat, Level, Speed, AGL: integer; STR: integer = 0; ItemSTR: integer = 0): longword; begin
'  Result := longword(Speed * 1000) div (longword((((Level * (Combat + 2)) + 45) * (AGL + 150)) * 1500) div 9000);
'  If (STR < ItemSTR) Then
'    Result := longword(longword((longword(ItemSTR - STR) * 3) + 200) *
'Result) div 200; end;

CalcEnergyUsed = Fix((nSpeed * 1000) / Fix(((((nLevel * (nCombat + 2)) + 45) * (nAGL + 150)) * 1500) / 9000))

If (nSTR < nItemSTR) Then
    CalcEnergyUsed = Fix(((((nItemSTR - nSTR) * 3) + 200) * CalcEnergyUsed) / 200)
End If

End Function

Private Function CalcEnergyUsedWithEncum(ByVal nCombat As Currency, ByVal nLevel As Currency, _
    ByVal nSpeed As Currency, ByVal nAGL As Currency, ByVal nSTR As Currency, ByVal nEncum As Currency, _
    Optional ByVal nItemSTR As Currency = 0) As Currency
'{ Calculates the energy used for a given Combat rating, Level, Speed, AGL, STR,
'  Encumbrance, and ItemSTR }
'function  CalcEnergyUsedWithEncum(Combat, Level, Speed, AGL, STR: integer;
'Encumbrance: integer; ItemSTR: integer = 0): integer; begin
'  Result := CalcEnergyUsed(Combat, Level, Speed, AGL, STR, ItemSTR);
'  Result := (Result * ((Encumbrance div 2) + 75)) div 100; end;
    
CalcEnergyUsedWithEncum = CalcEnergyUsed(nCombat, nLevel, nSpeed, nAGL, nSTR, nItemSTR)
CalcEnergyUsedWithEncum = Fix((CalcEnergyUsedWithEncum * Fix(Fix(nEncum / 2) + 75)) / 100)

End Function

Private Function IsQuickAndDeadly(ByVal nEU As Currency, ByVal nEncum As Currency) As Boolean
'{ Determines whether the quick and deadly bonus message is displayed when a
'  weapon is wielded }
'function  IsQuickAndDeadly(EU, Encumbrance: integer): boolean; begin
'  Result := (EU < 200) and (Encumbrance < 67); end;

If (nEU < 200) And (nEncum < 67) Then IsQuickAndDeadly = True

End Function

Private Function CalcQuickAndDeadlyBonus(ByVal nAGL As Currency, ByVal nEU As Currency, _
    ByVal nEncum As Currency) As Currency
' { Calculates the critical hit chance bonus for being quick and deadly with a
'   weapon for a previously calculated energy use }
' function  CalcQuickAndDeadlyBonus(AGL, EU, Encumbrance: integer): integer;
' begin
'   Result := 0;
'   If (EU >= 200) Or (Encumbrance > 66) Then
'     Exit;
'
'   Result := 200 - EU;
'   Result := Result + ((AGL - 50) div 10);
'
' //  Result := ((200 - EU) + ((AGL - 50) div 10));
'   If (Result > 20) Then
'     Result := 20;
'
'   If (Encumbrance >= 33) Then
'     Result := Result div 2;
' end;

CalcQuickAndDeadlyBonus = 0

If (nEU >= 200) Or (nEncum > 66) Then Exit Function

CalcQuickAndDeadlyBonus = (200 - nEU) + Fix((nAGL - 50) / 10)

If (CalcQuickAndDeadlyBonus > 20) Then CalcQuickAndDeadlyBonus = 20
If (nEncum >= 33) Then CalcQuickAndDeadlyBonus = Fix(CalcQuickAndDeadlyBonus / 2)

End Function

Private Function AdjustEnergyUsedWithSpeed(ByVal nEU As Currency, ByVal nSpeed As Currency) As Currency
'{ Adjusts a previously calculated energy use with a specified Speed amount }
'
'function  AdjustEnergyUsedWithSpeed(EU, Speed: integer): integer; begin
'  Result := (EU * Speed) div 100;
'end;

AdjustEnergyUsedWithSpeed = Fix((nEU * nSpeed) / 100)

End Function

Private Function AdjustEnergyUsedWithEncum(ByVal nEU As Currency, ByVal nEncum As Currency) As Currency
'{ Adjusts a previously calculated energy use with a specified Encumbrance
'  amount }
'function  AdjustEnergyUsedWithEncum(EU, Encumbrance: longword): longword; begin
'  Result := (EU * ((Encumbrance div 2) + 75)) div 100; end;

AdjustEnergyUsedWithEncum = Fix((nEU * (Fix(nEncum / 2) + 75)) / 100)

End Function

Private Sub optSpeed_Click(Index As Integer)
If optSpeed(3).Value = True Then txtSpeed.Enabled = True Else txtSpeed.Enabled = False
If optSpeed(2).Value = True Then
    chkSlowness.Value = 1
Else
    chkSlowness.Value = 0
End If
Call CalcSwings
End Sub

Private Sub timMouseDown_Timer()
timMouseDown.Enabled = False
End Sub

Private Sub txtAgility_GotFocus()
Call SelectAll(txtAgility)
End Sub

Private Sub txtAgility_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtAgility_KeyUp(KeyCode As Integer, Shift As Integer)
Call CalcSwings
End Sub

Private Sub txtEncum_GotFocus()
Call SelectAll(txtEncum)
End Sub

Private Sub txtEncum_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtEncum_KeyUp(KeyCode As Integer, Shift As Integer)
Call CalcSwings
End Sub

Private Sub txtLevel_GotFocus()
Call SelectAll(txtLevel)
End Sub

Private Sub txtLevel_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtLevel_KeyUp(KeyCode As Integer, Shift As Integer)
Call CalcSwings
End Sub

Private Sub txtMaxEncum_GotFocus()
Call SelectAll(txtMaxEncum)
End Sub

Private Sub txtMaxEncum_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtMaxEncum_KeyUp(KeyCode As Integer, Shift As Integer)
Call CalcSwings
End Sub

Private Sub txtSpeed_GotFocus()
Call SelectAll(txtSpeed)
End Sub

Private Sub txtSpeed_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtSpeed_KeyUp(KeyCode As Integer, Shift As Integer)
Call CalcSwings
End Sub

Private Sub txtStrength_Change()
txtMaxEncum.Text = CalcEncum(Val(txtStrength.Text))
End Sub

Private Sub txtStrength_GotFocus()
Call SelectAll(txtStrength)
End Sub

Private Sub txtStrength_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtStrength_KeyUp(KeyCode As Integer, Shift As Integer)

Call CalcSwings
End Sub

Private Sub txtTrueAVG_Change(Index As Integer)

On Error GoTo error:

txtTrueAVG(7).Text = CalcTrueAverage(Val(txtTrueAVG(6).Text), Val(txtTrueAVG(0).Text), Val(txtTrueAVG(1).Text), _
        Val(txtTrueAVG(4).Text), Val(txtTrueAVG(5).Text), Val(txtTrueAVG(2).Text), Val(txtTrueAVG(3).Text))

Exit Sub
error:
Call HandleError("txtTrueAVG_Change")
        
End Sub

Private Sub txtTrueAVG_GotFocus(Index As Integer)
Call SelectAll(txtTrueAVG(Index))
End Sub
