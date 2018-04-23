VERSION 5.00
Begin VB.Form frmMapLegend 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Legend"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "frmMapLegend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2775
   ScaleWidth      =   6240
   Begin VB.Label lblRoomCell 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C0C0C0&
      Height          =   135
      Index           =   0
      Left            =   2280
      TabIndex        =   30
      Top             =   990
      Width           =   135
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "NPC Assigned Room"
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
      Height          =   195
      Index           =   15
      Left            =   2520
      TabIndex        =   29
      Top             =   990
      Width           =   1515
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "- Right click on an up/down exit to have the option of following and redrawing."
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
      Height          =   495
      Left            =   4260
      TabIndex        =   28
      Top             =   2160
      Width           =   1875
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Lair Room"
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
      Height          =   195
      Index           =   17
      Left            =   2580
      TabIndex        =   27
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label lblRoomCell 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C0C0C0&
      Height          =   135
      Index           =   2
      Left            =   2280
      TabIndex        =   26
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Command In Room (Remote or Command)"
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
      Height          =   375
      Index           =   16
      Left            =   2580
      TabIndex        =   25
      Top             =   1620
      Width           =   1455
   End
   Begin VB.Label lblRoomCell 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C0C0C0&
      Height          =   135
      Index           =   1
      Left            =   2280
      TabIndex        =   24
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label lblRoomCell 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C0C0C0&
      Height          =   135
      Index           =   8
      Left            =   2280
      TabIndex        =   23
      Top             =   2100
      Width           =   135
   End
   Begin VB.Label lblRoomCell 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   135
      Index           =   7
      Left            =   2280
      TabIndex        =   22
      Top             =   360
      Width           =   135
   End
   Begin VB.Label lblRoomCell 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   135
      Index           =   6
      Left            =   2280
      TabIndex        =   21
      Top             =   60
      Width           =   135
   End
   Begin VB.Label lblRoomCell 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   5
      Left            =   2280
      TabIndex        =   20
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Room doesn't exist!"
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
      Height          =   195
      Index           =   14
      Left            =   2580
      TabIndex        =   19
      Top             =   2460
      Width           =   1275
   End
   Begin VB.Label lblRoomCell 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C0C0C0&
      Height          =   135
      Index           =   4
      Left            =   2280
      TabIndex        =   18
      Top             =   2460
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   255
      Left            =   2220
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Starting Point"
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
      Height          =   195
      Index           =   13
      Left            =   2580
      TabIndex        =   17
      Top             =   2100
      Width           =   1275
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Class / Race / Level / Alignment / Ability"
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
      Height          =   375
      Index           =   9
      Left            =   780
      TabIndex        =   16
      Top             =   1860
      Width           =   1275
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Timed"
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
      Height          =   195
      Index           =   8
      Left            =   780
      TabIndex        =   15
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Line linLegend 
      BorderColor     =   &H00400000&
      BorderWidth     =   5
      Index           =   9
      X1              =   120
      X2              =   600
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Line linLegend 
      BorderColor     =   &H00004000&
      BorderWidth     =   5
      Index           =   8
      X1              =   120
      X2              =   600
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Door / Gate"
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
      Height          =   195
      Index           =   7
      Left            =   780
      TabIndex        =   14
      Top             =   1020
      Width           =   1095
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Remote Action / Action"
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
      Height          =   375
      Index           =   6
      Left            =   780
      TabIndex        =   13
      Top             =   1500
      Width           =   1095
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Hidden"
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
      Height          =   195
      Index           =   5
      Left            =   780
      TabIndex        =   12
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Key / Item / Toll"
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
      Height          =   195
      Index           =   4
      Left            =   780
      TabIndex        =   11
      Top             =   1260
      Width           =   1095
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Text"
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
      Height          =   195
      Index           =   3
      Left            =   780
      TabIndex        =   10
      Top             =   780
      Width           =   1095
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Trap / Spell Trap"
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
      Height          =   195
      Index           =   2
      Left            =   780
      TabIndex        =   9
      Top             =   540
      Width           =   1095
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Map Change"
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
      Height          =   195
      Index           =   1
      Left            =   780
      TabIndex        =   8
      Top             =   300
      Width           =   1095
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "General"
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
      Height          =   195
      Index           =   0
      Left            =   780
      TabIndex        =   7
      Top             =   60
      Width           =   1095
   End
   Begin VB.Line linLegend 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      Index           =   7
      X1              =   120
      X2              =   600
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line linLegend 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   5
      Index           =   6
      X1              =   120
      X2              =   600
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Line linLegend 
      BorderColor     =   &H00800080&
      BorderWidth     =   5
      Index           =   5
      X1              =   120
      X2              =   600
      Y1              =   2580
      Y2              =   2580
   End
   Begin VB.Line linLegend 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   5
      Index           =   4
      X1              =   120
      X2              =   600
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line linLegend 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   5
      Index           =   3
      X1              =   120
      X2              =   600
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line linLegend 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Index           =   2
      X1              =   120
      X2              =   600
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line linLegend 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   5
      Index           =   1
      X1              =   120
      X2              =   600
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line linLegend 
      BorderColor     =   &H00808080&
      BorderWidth     =   5
      Index           =   0
      X1              =   120
      X2              =   600
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "-Right Click: Redraw map from selected room"
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
      Height          =   315
      Left            =   4260
      TabIndex        =   6
      Top             =   540
      Width           =   1875
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "-Hover mouse over room to see Information"
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
      Height          =   375
      Left            =   4260
      TabIndex        =   5
      Top             =   1740
      Width           =   1875
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "-Left Click: Jump to room in Room / Map Editor"
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
      Height          =   375
      Left            =   4260
      TabIndex        =   4
      Top             =   60
      Width           =   1875
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "-Shift+Right Click: Redraw map using selected cell as the ""center"".  This may also be done on unused cells."
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
      Height          =   675
      Left            =   4260
      TabIndex        =   3
      Top             =   960
      Width           =   1875
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   4140
      X2              =   4140
      Y1              =   60
      Y2              =   2700
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Exit Up"
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
      Height          =   195
      Index           =   10
      Left            =   2580
      TabIndex        =   2
      Top             =   60
      Width           =   1275
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Exit Down"
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
      Height          =   195
      Index           =   11
      Left            =   2580
      TabIndex        =   1
      Top             =   360
      Width           =   1275
   End
   Begin VB.Label lblLegendText 
      BackColor       =   &H00000000&
      Caption         =   "Exit Up and Down"
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
      Height          =   195
      Index           =   12
      Left            =   2580
      TabIndex        =   0
      Top             =   660
      Width           =   1275
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   225
      Left            =   2235
      Top             =   2055
      Width           =   225
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      Height          =   255
      Left            =   2220
      Shape           =   1  'Square
      Top             =   1620
      Width           =   255
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FF00FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   4
      Height          =   195
      Left            =   2250
      Shape           =   3  'Circle
      Top             =   1290
      Width           =   195
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   255
      Left            =   2220
      Shape           =   3  'Circle
      Top             =   930
      Width           =   255
   End
End
Attribute VB_Name = "frmMapLegend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Private Sub Form_Load()
On Error Resume Next
Me.Top = ReadINI("Windows", "MapLegendTop")
Me.Left = ReadINI("Windows", "MapLegendLeft")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Me.WindowState = vbMinimized Then Exit Sub
Call WriteINI("Windows", "MapLegendTop", Me.Top)
Call WriteINI("Windows", "MapLegendLeft", Me.Left)
End Sub

