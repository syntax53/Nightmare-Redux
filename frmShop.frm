VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20D5284F-7B23-4F0A-B8B1-6C9D18B64F1C}#1.0#0"; "exlimiter.ocx"
Begin VB.Form frmShop 
   Caption         =   "Shop Editor"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9810
   Icon            =   "frmShop.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   9810
   Begin VB.Frame framNav 
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   3120
      TabIndex        =   5
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton cmdDiscard 
         Caption         =   "Dis&card"
         Height          =   285
         Left            =   5580
         TabIndex        =   10
         Top             =   0
         Width           =   1035
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   0
         Width           =   1035
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "&Insert"
         Height          =   285
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdOther 
         Caption         =   "&General Info."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         TabIndex        =   8
         Top             =   0
         Width           =   1815
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   285
         Left            =   4620
         TabIndex        =   9
         Top             =   0
         Width           =   975
      End
      Begin VB.Frame frameGeneral 
         Caption         =   "General"
         Height          =   3375
         Left            =   540
         TabIndex        =   200
         Top             =   1680
         Visible         =   0   'False
         Width           =   4635
         Begin VB.CheckBox chkAutoSave 
            Caption         =   "Auto-Save"
            Height          =   195
            Left            =   3360
            TabIndex        =   214
            Top             =   180
            Value           =   1  'Checked
            Width           =   1155
         End
         Begin VB.TextBox txtBankAcct 
            Height          =   315
            Left            =   1440
            MaxLength       =   30
            TabIndex        =   213
            Top             =   2820
            Width           =   2775
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1440
            MaxLength       =   39
            TabIndex        =   202
            Top             =   720
            Width           =   2775
         End
         Begin VB.TextBox txtShopMaxLvl 
            Height          =   315
            Left            =   2400
            TabIndex        =   209
            Top             =   1980
            Width           =   855
         End
         Begin VB.TextBox txtShopMinLvl 
            Height          =   315
            Left            =   1440
            TabIndex        =   208
            Top             =   1980
            Width           =   855
         End
         Begin VB.TextBox txtShopMarkup 
            Height          =   315
            Left            =   1440
            MaxLength       =   5
            TabIndex        =   211
            Top             =   2400
            Width           =   855
         End
         Begin VB.ComboBox cmbType 
            Height          =   315
            ItemData        =   "frmShop.frx":08CA
            Left            =   1440
            List            =   "frmShop.frx":08F5
            Style           =   2  'Dropdown List
            TabIndex        =   204
            Top             =   1140
            Width           =   1815
         End
         Begin VB.ComboBox cmbClasses 
            Height          =   315
            ItemData        =   "frmShop.frx":0969
            Left            =   1440
            List            =   "frmShop.frx":096B
            Style           =   2  'Dropdown List
            TabIndex        =   206
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bank Account"
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   212
            Top             =   2880
            Width           =   1020
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Name"
            Height          =   195
            Index           =   1
            Left            =   840
            TabIndex        =   201
            Top             =   780
            Width           =   420
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Min/Max LVL"
            Height          =   195
            Index           =   9
            Left            =   300
            TabIndex        =   207
            Top             =   2040
            Width           =   960
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Type"
            Height          =   195
            Index           =   8
            Left            =   900
            TabIndex        =   203
            Top             =   1200
            Width           =   360
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Class"
            Height          =   195
            Index           =   3
            Left            =   885
            TabIndex        =   205
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Markup %"
            Height          =   195
            Index           =   2
            Left            =   555
            TabIndex        =   210
            Top             =   2460
            Width           =   705
         End
      End
      Begin exlimiter.EL EL1 
         Left            =   5820
         Top             =   -180
         _ExtentX        =   1270
         _ExtentY        =   1270
      End
      Begin VB.Frame frameShopItems 
         Caption         =   "Items Sold"
         Height          =   6435
         Left            =   0
         TabIndex        =   11
         Top             =   300
         Width           =   6615
         Begin VB.CommandButton cmdItemLookup 
            Height          =   195
            Index           =   9
            Left            =   180
            TabIndex        =   29
            Top             =   3120
            Width           =   195
         End
         Begin VB.CommandButton cmdItemLookup 
            Height          =   195
            Index           =   8
            Left            =   180
            TabIndex        =   28
            Top             =   2820
            Width           =   195
         End
         Begin VB.CommandButton cmdItemLookup 
            Height          =   195
            Index           =   7
            Left            =   180
            TabIndex        =   27
            Top             =   2520
            Width           =   195
         End
         Begin VB.CommandButton cmdItemLookup 
            Height          =   195
            Index           =   6
            Left            =   180
            TabIndex        =   26
            Top             =   2220
            Width           =   195
         End
         Begin VB.CommandButton cmdItemLookup 
            Height          =   195
            Index           =   5
            Left            =   180
            TabIndex        =   25
            Top             =   1920
            Width           =   195
         End
         Begin VB.CommandButton cmdItemLookup 
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   24
            Top             =   1620
            Width           =   195
         End
         Begin VB.CommandButton cmdItemLookup 
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   23
            Top             =   1320
            Width           =   195
         End
         Begin VB.CommandButton cmdItemLookup 
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   22
            Top             =   1020
            Width           =   195
         End
         Begin VB.CommandButton cmdItemLookup 
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   21
            Top             =   720
            Width           =   195
         End
         Begin VB.CommandButton cmdItemLookup 
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   20
            Top             =   420
            Width           =   195
         End
         Begin VB.TextBox txtShopItemNumber 
            Height          =   315
            Index           =   0
            Left            =   420
            TabIndex        =   40
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtShopItemNumber 
            Height          =   315
            Index           =   1
            Left            =   420
            TabIndex        =   41
            Top             =   660
            Width           =   615
         End
         Begin VB.TextBox txtShopItemNumber 
            Height          =   315
            Index           =   2
            Left            =   420
            TabIndex        =   42
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtShopItemNumber 
            Height          =   315
            Index           =   3
            Left            =   420
            TabIndex        =   43
            Top             =   1260
            Width           =   615
         End
         Begin VB.TextBox txtShopItemNumber 
            Height          =   315
            Index           =   4
            Left            =   420
            TabIndex        =   44
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox txtShopItemNumber 
            Height          =   315
            Index           =   5
            Left            =   420
            TabIndex        =   45
            Top             =   1860
            Width           =   615
         End
         Begin VB.TextBox txtShopItemNumber 
            Height          =   315
            Index           =   6
            Left            =   420
            TabIndex        =   46
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtShopItemNumber 
            Height          =   315
            Index           =   7
            Left            =   420
            TabIndex        =   47
            Top             =   2460
            Width           =   615
         End
         Begin VB.TextBox txtShopItemNumber 
            Height          =   315
            Index           =   8
            Left            =   420
            TabIndex        =   48
            Top             =   2760
            Width           =   615
         End
         Begin VB.TextBox txtShopItemNumber 
            Height          =   315
            Index           =   9
            Left            =   420
            TabIndex        =   49
            Top             =   3060
            Width           =   615
         End
         Begin VB.TextBox txtShopItemName 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   0
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   29
            TabIndex        =   160
            TabStop         =   0   'False
            Top             =   360
            Width           =   1875
         End
         Begin VB.TextBox txtShopItemName 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   1
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   29
            TabIndex        =   162
            TabStop         =   0   'False
            Top             =   660
            Width           =   1875
         End
         Begin VB.TextBox txtShopItemName 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   2
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   29
            TabIndex        =   164
            TabStop         =   0   'False
            Top             =   960
            Width           =   1875
         End
         Begin VB.TextBox txtShopItemName 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   3
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   29
            TabIndex        =   166
            TabStop         =   0   'False
            Top             =   1260
            Width           =   1875
         End
         Begin VB.TextBox txtShopItemName 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   4
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   29
            TabIndex        =   168
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1875
         End
         Begin VB.TextBox txtShopItemName 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   5
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   29
            TabIndex        =   170
            TabStop         =   0   'False
            Top             =   1860
            Width           =   1875
         End
         Begin VB.TextBox txtShopItemName 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   6
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   29
            TabIndex        =   172
            TabStop         =   0   'False
            Top             =   2160
            Width           =   1875
         End
         Begin VB.TextBox txtShopItemName 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   7
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   29
            TabIndex        =   174
            TabStop         =   0   'False
            Top             =   2460
            Width           =   1875
         End
         Begin VB.TextBox txtShopItemName 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   8
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   29
            TabIndex        =   176
            TabStop         =   0   'False
            Top             =   2760
            Width           =   1875
         End
         Begin VB.TextBox txtShopItemName 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   9
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   29
            TabIndex        =   178
            TabStop         =   0   'False
            Top             =   3060
            Width           =   1875
         End
         Begin VB.TextBox txtShopNormal 
            Height          =   315
            Index           =   0
            Left            =   2880
            TabIndex        =   60
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtShopNormal 
            Height          =   315
            Index           =   1
            Left            =   2880
            TabIndex        =   61
            Top             =   660
            Width           =   495
         End
         Begin VB.TextBox txtShopNormal 
            Height          =   315
            Index           =   2
            Left            =   2880
            TabIndex        =   62
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtShopNormal 
            Height          =   315
            Index           =   3
            Left            =   2880
            TabIndex        =   63
            Top             =   1260
            Width           =   495
         End
         Begin VB.TextBox txtShopNormal 
            Height          =   315
            Index           =   4
            Left            =   2880
            TabIndex        =   64
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtShopNormal 
            Height          =   315
            Index           =   5
            Left            =   2880
            TabIndex        =   65
            Top             =   1860
            Width           =   495
         End
         Begin VB.TextBox txtShopNormal 
            Height          =   315
            Index           =   6
            Left            =   2880
            TabIndex        =   66
            Top             =   2160
            Width           =   495
         End
         Begin VB.TextBox txtShopNormal 
            Height          =   315
            Index           =   7
            Left            =   2880
            TabIndex        =   67
            Top             =   2460
            Width           =   495
         End
         Begin VB.TextBox txtShopNormal 
            Height          =   315
            Index           =   8
            Left            =   2880
            TabIndex        =   68
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox txtShopNormal 
            Height          =   315
            Index           =   9
            Left            =   2880
            TabIndex        =   69
            Top             =   3060
            Width           =   495
         End
         Begin VB.TextBox txtShopMax 
            Height          =   315
            Index           =   0
            Left            =   3360
            TabIndex        =   80
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtShopMax 
            Height          =   315
            Index           =   1
            Left            =   3360
            TabIndex        =   81
            Top             =   660
            Width           =   495
         End
         Begin VB.TextBox txtShopMax 
            Height          =   315
            Index           =   2
            Left            =   3360
            TabIndex        =   82
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtShopMax 
            Height          =   315
            Index           =   3
            Left            =   3360
            TabIndex        =   83
            Top             =   1260
            Width           =   495
         End
         Begin VB.TextBox txtShopMax 
            Height          =   315
            Index           =   4
            Left            =   3360
            TabIndex        =   84
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtShopMax 
            Height          =   315
            Index           =   5
            Left            =   3360
            TabIndex        =   85
            Top             =   1860
            Width           =   495
         End
         Begin VB.TextBox txtShopMax 
            Height          =   315
            Index           =   6
            Left            =   3360
            TabIndex        =   86
            Top             =   2160
            Width           =   495
         End
         Begin VB.TextBox txtShopMax 
            Height          =   315
            Index           =   7
            Left            =   3360
            TabIndex        =   87
            Top             =   2460
            Width           =   495
         End
         Begin VB.TextBox txtShopMax 
            Height          =   315
            Index           =   8
            Left            =   3360
            TabIndex        =   88
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox txtShopMax 
            Height          =   315
            Index           =   9
            Left            =   3360
            TabIndex        =   89
            Top             =   3060
            Width           =   495
         End
         Begin VB.TextBox txtShopRgnTime 
            Height          =   315
            Index           =   0
            Left            =   3840
            TabIndex        =   100
            ToolTipText     =   "Time in minutes"
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtShopRgnTime 
            Height          =   315
            Index           =   1
            Left            =   3840
            TabIndex        =   101
            ToolTipText     =   "Time in minutes"
            Top             =   660
            Width           =   495
         End
         Begin VB.TextBox txtShopRgnTime 
            Height          =   315
            Index           =   2
            Left            =   3840
            TabIndex        =   102
            ToolTipText     =   "Time in minutes"
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtShopRgnTime 
            Height          =   315
            Index           =   3
            Left            =   3840
            TabIndex        =   103
            ToolTipText     =   "Time in minutes"
            Top             =   1260
            Width           =   495
         End
         Begin VB.TextBox txtShopRgnTime 
            Height          =   315
            Index           =   4
            Left            =   3840
            TabIndex        =   104
            ToolTipText     =   "Time in minutes"
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtShopRgnTime 
            Height          =   315
            Index           =   5
            Left            =   3840
            TabIndex        =   105
            ToolTipText     =   "Time in minutes"
            Top             =   1860
            Width           =   495
         End
         Begin VB.TextBox txtShopRgnTime 
            Height          =   315
            Index           =   6
            Left            =   3840
            TabIndex        =   106
            ToolTipText     =   "Time in minutes"
            Top             =   2160
            Width           =   495
         End
         Begin VB.TextBox txtShopRgnTime 
            Height          =   315
            Index           =   7
            Left            =   3840
            TabIndex        =   107
            ToolTipText     =   "Time in minutes"
            Top             =   2460
            Width           =   495
         End
         Begin VB.TextBox txtShopRgnTime 
            Height          =   315
            Index           =   8
            Left            =   3840
            TabIndex        =   108
            ToolTipText     =   "Time in minutes"
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox txtShopRgnTime 
            Height          =   315
            Index           =   9
            Left            =   3840
            TabIndex        =   109
            ToolTipText     =   "Time in minutes"
            Top             =   3060
            Width           =   495
         End
         Begin VB.TextBox txtRgnPercentage 
            Height          =   315
            Index           =   0
            Left            =   4320
            TabIndex        =   120
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtRgnPercentage 
            Height          =   315
            Index           =   1
            Left            =   4320
            TabIndex        =   121
            Top             =   660
            Width           =   495
         End
         Begin VB.TextBox txtRgnPercentage 
            Height          =   315
            Index           =   2
            Left            =   4320
            TabIndex        =   122
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtRgnPercentage 
            Height          =   315
            Index           =   3
            Left            =   4320
            TabIndex        =   123
            Top             =   1260
            Width           =   495
         End
         Begin VB.TextBox txtRgnPercentage 
            Height          =   315
            Index           =   4
            Left            =   4320
            TabIndex        =   124
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtRgnPercentage 
            Height          =   315
            Index           =   5
            Left            =   4320
            TabIndex        =   125
            Top             =   1860
            Width           =   495
         End
         Begin VB.TextBox txtRgnPercentage 
            Height          =   315
            Index           =   6
            Left            =   4320
            TabIndex        =   126
            Top             =   2160
            Width           =   495
         End
         Begin VB.TextBox txtRgnPercentage 
            Height          =   315
            Index           =   7
            Left            =   4320
            TabIndex        =   127
            Top             =   2460
            Width           =   495
         End
         Begin VB.TextBox txtRgnPercentage 
            Height          =   315
            Index           =   8
            Left            =   4320
            TabIndex        =   128
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox txtRgnPercentage 
            Height          =   315
            Index           =   9
            Left            =   4320
            TabIndex        =   129
            Top             =   3060
            Width           =   495
         End
         Begin VB.TextBox txtRgnNumber 
            Height          =   315
            Index           =   0
            Left            =   4800
            TabIndex        =   140
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtRgnNumber 
            Height          =   315
            Index           =   1
            Left            =   4800
            TabIndex        =   141
            Top             =   660
            Width           =   495
         End
         Begin VB.TextBox txtRgnNumber 
            Height          =   315
            Index           =   2
            Left            =   4800
            TabIndex        =   142
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtRgnNumber 
            Height          =   315
            Index           =   3
            Left            =   4800
            TabIndex        =   143
            Top             =   1260
            Width           =   495
         End
         Begin VB.TextBox txtRgnNumber 
            Height          =   315
            Index           =   4
            Left            =   4800
            TabIndex        =   144
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtRgnNumber 
            Height          =   315
            Index           =   5
            Left            =   4800
            TabIndex        =   145
            Top             =   1860
            Width           =   495
         End
         Begin VB.TextBox txtRgnNumber 
            Height          =   315
            Index           =   6
            Left            =   4800
            TabIndex        =   146
            Top             =   2160
            Width           =   495
         End
         Begin VB.TextBox txtRgnNumber 
            Height          =   315
            Index           =   7
            Left            =   4800
            TabIndex        =   147
            Top             =   2460
            Width           =   495
         End
         Begin VB.TextBox txtRgnNumber 
            Height          =   315
            Index           =   8
            Left            =   4800
            TabIndex        =   148
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox txtRgnNumber 
            Height          =   315
            Index           =   9
            Left            =   4800
            TabIndex        =   149
            Top             =   3060
            Width           =   495
         End
         Begin VB.TextBox txtShopItemNumber 
            Height          =   315
            Index           =   10
            Left            =   420
            TabIndex        =   50
            Top             =   3360
            Width           =   615
         End
         Begin VB.CommandButton cmdItemLookup 
            Height          =   195
            Index           =   19
            Left            =   180
            TabIndex        =   39
            Top             =   6120
            Width           =   195
         End
         Begin VB.CommandButton cmdItemLookup 
            Height          =   195
            Index           =   18
            Left            =   180
            TabIndex        =   38
            Top             =   5820
            Width           =   195
         End
         Begin VB.CommandButton cmdItemLookup 
            Height          =   195
            Index           =   17
            Left            =   180
            TabIndex        =   37
            Top             =   5520
            Width           =   195
         End
         Begin VB.CommandButton cmdItemLookup 
            Height          =   195
            Index           =   16
            Left            =   180
            TabIndex        =   36
            Top             =   5220
            Width           =   195
         End
         Begin VB.CommandButton cmdItemLookup 
            Height          =   195
            Index           =   15
            Left            =   180
            TabIndex        =   35
            Top             =   4920
            Width           =   195
         End
         Begin VB.CommandButton cmdItemLookup 
            Height          =   195
            Index           =   14
            Left            =   180
            TabIndex        =   34
            Top             =   4620
            Width           =   195
         End
         Begin VB.CommandButton cmdItemLookup 
            Height          =   195
            Index           =   13
            Left            =   180
            TabIndex        =   33
            Top             =   4320
            Width           =   195
         End
         Begin VB.CommandButton cmdItemLookup 
            Height          =   195
            Index           =   12
            Left            =   180
            TabIndex        =   32
            Top             =   4020
            Width           =   195
         End
         Begin VB.CommandButton cmdItemLookup 
            Height          =   195
            Index           =   11
            Left            =   180
            TabIndex        =   31
            Top             =   3720
            Width           =   195
         End
         Begin VB.CommandButton cmdItemLookup 
            Height          =   195
            Index           =   10
            Left            =   180
            TabIndex        =   30
            Top             =   3420
            Width           =   195
         End
         Begin VB.TextBox txtShopItemNumber 
            Height          =   315
            Index           =   11
            Left            =   420
            TabIndex        =   51
            Top             =   3660
            Width           =   615
         End
         Begin VB.TextBox txtShopItemNumber 
            Height          =   315
            Index           =   12
            Left            =   420
            TabIndex        =   52
            Top             =   3960
            Width           =   615
         End
         Begin VB.TextBox txtShopItemNumber 
            Height          =   315
            Index           =   13
            Left            =   420
            TabIndex        =   53
            Top             =   4260
            Width           =   615
         End
         Begin VB.TextBox txtShopItemNumber 
            Height          =   315
            Index           =   14
            Left            =   420
            TabIndex        =   54
            Top             =   4560
            Width           =   615
         End
         Begin VB.TextBox txtShopItemNumber 
            Height          =   315
            Index           =   15
            Left            =   420
            TabIndex        =   55
            Top             =   4860
            Width           =   615
         End
         Begin VB.TextBox txtShopItemNumber 
            Height          =   315
            Index           =   16
            Left            =   420
            TabIndex        =   56
            Top             =   5160
            Width           =   615
         End
         Begin VB.TextBox txtShopItemNumber 
            Height          =   315
            Index           =   17
            Left            =   420
            TabIndex        =   57
            Top             =   5460
            Width           =   615
         End
         Begin VB.TextBox txtShopItemName 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   10
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   29
            TabIndex        =   180
            TabStop         =   0   'False
            Top             =   3360
            Width           =   1875
         End
         Begin VB.TextBox txtShopItemName 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   11
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   29
            TabIndex        =   182
            TabStop         =   0   'False
            Top             =   3660
            Width           =   1875
         End
         Begin VB.TextBox txtShopItemName 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   12
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   29
            TabIndex        =   184
            TabStop         =   0   'False
            Top             =   3960
            Width           =   1875
         End
         Begin VB.TextBox txtShopItemName 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   13
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   29
            TabIndex        =   186
            TabStop         =   0   'False
            Top             =   4260
            Width           =   1875
         End
         Begin VB.TextBox txtShopItemName 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   14
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   29
            TabIndex        =   188
            TabStop         =   0   'False
            Top             =   4560
            Width           =   1875
         End
         Begin VB.TextBox txtShopItemName 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   15
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   29
            TabIndex        =   190
            TabStop         =   0   'False
            Top             =   4860
            Width           =   1875
         End
         Begin VB.TextBox txtShopItemName 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   16
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   29
            TabIndex        =   192
            TabStop         =   0   'False
            Top             =   5160
            Width           =   1875
         End
         Begin VB.TextBox txtShopItemName 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   17
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   29
            TabIndex        =   194
            TabStop         =   0   'False
            Top             =   5460
            Width           =   1875
         End
         Begin VB.TextBox txtShopItemNumber 
            Height          =   315
            Index           =   18
            Left            =   420
            TabIndex        =   58
            Top             =   5760
            Width           =   615
         End
         Begin VB.TextBox txtShopItemNumber 
            Height          =   315
            Index           =   19
            Left            =   420
            TabIndex        =   59
            Top             =   6060
            Width           =   615
         End
         Begin VB.TextBox txtShopItemName 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   18
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   29
            TabIndex        =   196
            TabStop         =   0   'False
            Top             =   5760
            Width           =   1875
         End
         Begin VB.TextBox txtShopItemName 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   19
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   29
            TabIndex        =   198
            TabStop         =   0   'False
            Top             =   6060
            Width           =   1875
         End
         Begin VB.TextBox txtShopNormal 
            Height          =   315
            Index           =   10
            Left            =   2880
            TabIndex        =   70
            Top             =   3360
            Width           =   495
         End
         Begin VB.TextBox txtShopNormal 
            Height          =   315
            Index           =   11
            Left            =   2880
            TabIndex        =   71
            Top             =   3660
            Width           =   495
         End
         Begin VB.TextBox txtShopNormal 
            Height          =   315
            Index           =   12
            Left            =   2880
            TabIndex        =   72
            Top             =   3960
            Width           =   495
         End
         Begin VB.TextBox txtShopNormal 
            Height          =   315
            Index           =   13
            Left            =   2880
            TabIndex        =   73
            Top             =   4260
            Width           =   495
         End
         Begin VB.TextBox txtShopNormal 
            Height          =   315
            Index           =   14
            Left            =   2880
            TabIndex        =   74
            Top             =   4560
            Width           =   495
         End
         Begin VB.TextBox txtShopNormal 
            Height          =   315
            Index           =   15
            Left            =   2880
            TabIndex        =   75
            Top             =   4860
            Width           =   495
         End
         Begin VB.TextBox txtShopNormal 
            Height          =   315
            Index           =   16
            Left            =   2880
            TabIndex        =   76
            Top             =   5160
            Width           =   495
         End
         Begin VB.TextBox txtShopNormal 
            Height          =   315
            Index           =   17
            Left            =   2880
            TabIndex        =   77
            Top             =   5460
            Width           =   495
         End
         Begin VB.TextBox txtShopNormal 
            Height          =   315
            Index           =   18
            Left            =   2880
            TabIndex        =   78
            Top             =   5760
            Width           =   495
         End
         Begin VB.TextBox txtShopNormal 
            Height          =   315
            Index           =   19
            Left            =   2880
            TabIndex        =   79
            Top             =   6060
            Width           =   495
         End
         Begin VB.TextBox txtShopMax 
            Height          =   315
            Index           =   10
            Left            =   3360
            TabIndex        =   90
            Top             =   3360
            Width           =   495
         End
         Begin VB.TextBox txtShopMax 
            Height          =   315
            Index           =   11
            Left            =   3360
            TabIndex        =   91
            Top             =   3660
            Width           =   495
         End
         Begin VB.TextBox txtShopMax 
            Height          =   315
            Index           =   12
            Left            =   3360
            TabIndex        =   92
            Top             =   3960
            Width           =   495
         End
         Begin VB.TextBox txtShopMax 
            Height          =   315
            Index           =   13
            Left            =   3360
            TabIndex        =   93
            Top             =   4260
            Width           =   495
         End
         Begin VB.TextBox txtShopMax 
            Height          =   315
            Index           =   14
            Left            =   3360
            TabIndex        =   94
            Top             =   4560
            Width           =   495
         End
         Begin VB.TextBox txtShopMax 
            Height          =   315
            Index           =   15
            Left            =   3360
            TabIndex        =   95
            Top             =   4860
            Width           =   495
         End
         Begin VB.TextBox txtShopMax 
            Height          =   315
            Index           =   16
            Left            =   3360
            TabIndex        =   96
            Top             =   5160
            Width           =   495
         End
         Begin VB.TextBox txtShopMax 
            Height          =   315
            Index           =   17
            Left            =   3360
            TabIndex        =   97
            Top             =   5460
            Width           =   495
         End
         Begin VB.TextBox txtShopMax 
            Height          =   315
            Index           =   18
            Left            =   3360
            TabIndex        =   98
            Top             =   5760
            Width           =   495
         End
         Begin VB.TextBox txtShopMax 
            Height          =   315
            Index           =   19
            Left            =   3360
            TabIndex        =   99
            Top             =   6060
            Width           =   495
         End
         Begin VB.TextBox txtShopRgnTime 
            Height          =   315
            Index           =   10
            Left            =   3840
            TabIndex        =   110
            ToolTipText     =   "Time in minutes"
            Top             =   3360
            Width           =   495
         End
         Begin VB.TextBox txtShopRgnTime 
            Height          =   315
            Index           =   11
            Left            =   3840
            TabIndex        =   111
            ToolTipText     =   "Time in minutes"
            Top             =   3660
            Width           =   495
         End
         Begin VB.TextBox txtShopRgnTime 
            Height          =   315
            Index           =   12
            Left            =   3840
            TabIndex        =   112
            ToolTipText     =   "Time in minutes"
            Top             =   3960
            Width           =   495
         End
         Begin VB.TextBox txtShopRgnTime 
            Height          =   315
            Index           =   13
            Left            =   3840
            TabIndex        =   113
            ToolTipText     =   "Time in minutes"
            Top             =   4260
            Width           =   495
         End
         Begin VB.TextBox txtShopRgnTime 
            Height          =   315
            Index           =   14
            Left            =   3840
            TabIndex        =   114
            ToolTipText     =   "Time in minutes"
            Top             =   4560
            Width           =   495
         End
         Begin VB.TextBox txtShopRgnTime 
            Height          =   315
            Index           =   15
            Left            =   3840
            TabIndex        =   115
            ToolTipText     =   "Time in minutes"
            Top             =   4860
            Width           =   495
         End
         Begin VB.TextBox txtShopRgnTime 
            Height          =   315
            Index           =   16
            Left            =   3840
            TabIndex        =   116
            ToolTipText     =   "Time in minutes"
            Top             =   5160
            Width           =   495
         End
         Begin VB.TextBox txtShopRgnTime 
            Height          =   315
            Index           =   17
            Left            =   3840
            TabIndex        =   117
            ToolTipText     =   "Time in minutes"
            Top             =   5460
            Width           =   495
         End
         Begin VB.TextBox txtShopRgnTime 
            Height          =   315
            Index           =   18
            Left            =   3840
            TabIndex        =   118
            ToolTipText     =   "Time in minutes"
            Top             =   5760
            Width           =   495
         End
         Begin VB.TextBox txtShopRgnTime 
            Height          =   315
            Index           =   19
            Left            =   3840
            TabIndex        =   119
            ToolTipText     =   "Time in minutes"
            Top             =   6060
            Width           =   495
         End
         Begin VB.TextBox txtRgnPercentage 
            Height          =   315
            Index           =   10
            Left            =   4320
            TabIndex        =   130
            Top             =   3360
            Width           =   495
         End
         Begin VB.TextBox txtRgnPercentage 
            Height          =   315
            Index           =   11
            Left            =   4320
            TabIndex        =   131
            Top             =   3660
            Width           =   495
         End
         Begin VB.TextBox txtRgnPercentage 
            Height          =   315
            Index           =   12
            Left            =   4320
            TabIndex        =   132
            Top             =   3960
            Width           =   495
         End
         Begin VB.TextBox txtRgnPercentage 
            Height          =   315
            Index           =   13
            Left            =   4320
            TabIndex        =   133
            Top             =   4260
            Width           =   495
         End
         Begin VB.TextBox txtRgnPercentage 
            Height          =   315
            Index           =   14
            Left            =   4320
            TabIndex        =   134
            Top             =   4560
            Width           =   495
         End
         Begin VB.TextBox txtRgnPercentage 
            Height          =   315
            Index           =   15
            Left            =   4320
            TabIndex        =   135
            Top             =   4860
            Width           =   495
         End
         Begin VB.TextBox txtRgnPercentage 
            Height          =   315
            Index           =   16
            Left            =   4320
            TabIndex        =   136
            Top             =   5160
            Width           =   495
         End
         Begin VB.TextBox txtRgnPercentage 
            Height          =   315
            Index           =   17
            Left            =   4320
            TabIndex        =   137
            Top             =   5460
            Width           =   495
         End
         Begin VB.TextBox txtRgnPercentage 
            Height          =   315
            Index           =   18
            Left            =   4320
            TabIndex        =   138
            Top             =   5760
            Width           =   495
         End
         Begin VB.TextBox txtRgnPercentage 
            Height          =   315
            Index           =   19
            Left            =   4320
            TabIndex        =   139
            Top             =   6060
            Width           =   495
         End
         Begin VB.TextBox txtRgnNumber 
            Height          =   315
            Index           =   10
            Left            =   4800
            TabIndex        =   150
            Top             =   3360
            Width           =   495
         End
         Begin VB.TextBox txtRgnNumber 
            Height          =   315
            Index           =   11
            Left            =   4800
            TabIndex        =   151
            Top             =   3660
            Width           =   495
         End
         Begin VB.TextBox txtRgnNumber 
            Height          =   315
            Index           =   12
            Left            =   4800
            TabIndex        =   152
            Top             =   3960
            Width           =   495
         End
         Begin VB.TextBox txtRgnNumber 
            Height          =   315
            Index           =   13
            Left            =   4800
            TabIndex        =   153
            Top             =   4260
            Width           =   495
         End
         Begin VB.TextBox txtRgnNumber 
            Height          =   315
            Index           =   14
            Left            =   4800
            TabIndex        =   154
            Top             =   4560
            Width           =   495
         End
         Begin VB.TextBox txtRgnNumber 
            Height          =   315
            Index           =   15
            Left            =   4800
            TabIndex        =   155
            Top             =   4860
            Width           =   495
         End
         Begin VB.TextBox txtRgnNumber 
            Height          =   315
            Index           =   16
            Left            =   4800
            TabIndex        =   156
            Top             =   5160
            Width           =   495
         End
         Begin VB.TextBox txtRgnNumber 
            Height          =   315
            Index           =   17
            Left            =   4800
            TabIndex        =   157
            Top             =   5460
            Width           =   495
         End
         Begin VB.TextBox txtRgnNumber 
            Height          =   315
            Index           =   18
            Left            =   4800
            TabIndex        =   158
            Top             =   5760
            Width           =   495
         End
         Begin VB.TextBox txtRgnNumber 
            Height          =   315
            Index           =   19
            Left            =   4800
            TabIndex        =   159
            Top             =   6060
            Width           =   495
         End
         Begin VB.TextBox txtCost 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   0
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   161
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   1
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   163
            TabStop         =   0   'False
            Top             =   660
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   2
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   165
            TabStop         =   0   'False
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   3
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   167
            TabStop         =   0   'False
            Top             =   1260
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   4
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   169
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   5
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   171
            TabStop         =   0   'False
            Top             =   1860
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   6
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   173
            TabStop         =   0   'False
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   7
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   175
            TabStop         =   0   'False
            Top             =   2460
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   8
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   177
            TabStop         =   0   'False
            Top             =   2760
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   9
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   179
            TabStop         =   0   'False
            Top             =   3060
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   10
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   181
            TabStop         =   0   'False
            Top             =   3360
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   11
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   183
            TabStop         =   0   'False
            Top             =   3660
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   12
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   185
            TabStop         =   0   'False
            Top             =   3960
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   13
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   187
            TabStop         =   0   'False
            Top             =   4260
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   14
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   189
            TabStop         =   0   'False
            Top             =   4560
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   15
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   191
            TabStop         =   0   'False
            Top             =   4860
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   16
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   193
            TabStop         =   0   'False
            Top             =   5160
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   17
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   195
            TabStop         =   0   'False
            Top             =   5460
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   18
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   197
            TabStop         =   0   'False
            Top             =   5760
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   19
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   199
            TabStop         =   0   'False
            Top             =   6060
            Width           =   1215
         End
         Begin VB.Label lblColumns 
            Alignment       =   2  'Center
            Caption         =   "Cost"
            Height          =   195
            Index           =   7
            Left            =   5340
            TabIndex        =   19
            Top             =   180
            Width           =   1095
         End
         Begin VB.Label lblColumns 
            Alignment       =   2  'Center
            Caption         =   "#"
            Height          =   195
            Index           =   0
            Left            =   420
            TabIndex        =   12
            Top             =   180
            Width           =   615
         End
         Begin VB.Label lblColumns 
            Alignment       =   2  'Center
            Caption         =   "Item Name"
            Height          =   195
            Index           =   1
            Left            =   1020
            TabIndex        =   13
            Top             =   180
            Width           =   1875
         End
         Begin VB.Label lblColumns 
            Alignment       =   2  'Center
            Caption         =   "Now"
            Height          =   255
            Index           =   2
            Left            =   2880
            TabIndex        =   14
            Top             =   180
            Width           =   495
         End
         Begin VB.Label lblColumns 
            Alignment       =   2  'Center
            Caption         =   "Max"
            Height          =   255
            Index           =   3
            Left            =   3360
            TabIndex        =   15
            Top             =   180
            Width           =   495
         End
         Begin VB.Label lblColumns 
            Alignment       =   2  'Center
            Caption         =   "Time"
            Height          =   255
            Index           =   4
            Left            =   3840
            TabIndex        =   16
            Top             =   180
            Width           =   495
         End
         Begin VB.Label lblColumns 
            Alignment       =   2  'Center
            Caption         =   "Rgn%"
            Height          =   255
            Index           =   5
            Left            =   4320
            TabIndex        =   17
            Top             =   180
            Width           =   495
         End
         Begin VB.Label lblColumns 
            Alignment       =   2  'Center
            Caption         =   "Rgn#"
            Height          =   255
            Index           =   6
            Left            =   4800
            TabIndex        =   18
            Top             =   180
            Width           =   495
         End
      End
   End
   Begin VB.TextBox txtNumberSearch 
      Height          =   285
      Left            =   60
      TabIndex        =   2
      Top             =   180
      Width           =   615
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   180
      Width           =   2295
   End
   Begin MSComctlLib.ListView lvDatabase 
      Height          =   6195
      Left            =   60
      TabIndex        =   4
      Top             =   480
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   10927
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
      Top             =   0
      Width           =   615
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
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "frmShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim bLoaded As Boolean
Dim nCurrentRecord As Long

Private Sub cmbType_Click()
Dim x As Integer
'0=General
'1=Weapons
'2=Armour
'3=Items
'4=Spells
'5=Hospital
'6=Tavern
'7=Bank
'8=Training
'9=Inn
'10=Specific
'11=Gang Shop
'12=Deed Shop

On Error GoTo Error:

cmbClasses.Enabled = False
txtShopMinLvl.Enabled = False
txtShopMaxLvl.Enabled = False
txtBankAcct.Enabled = False
txtBankAcct.Text = "<not a gangshop>"

lblColumns(3).Caption = "Max"
lblColumns(4).Caption = "Time"
lblColumns(5).Caption = "Rgn%"
lblColumns(6).Caption = "Rgn#"
    
If cmbType.ListIndex = 8 Then 'training
    cmbClasses.Enabled = True
    txtShopMinLvl.Enabled = True
    txtShopMaxLvl.Enabled = True
    For x = 0 To 19
        txtShopItemName(x).Enabled = True
        txtShopItemNumber(x).Enabled = True
        txtShopMax(x).Enabled = True
        txtShopNormal(x).Enabled = True
        txtShopRgnTime(x).Enabled = True
        txtRgnNumber(x).Enabled = True
        txtRgnPercentage(x).Enabled = True
        txtCost(x).Enabled = True
    Next
'ElseIf cmbType.ListIndex = 7 Then 'bank
'    For x = 0 To 19
'        txtShopItemName(x).Enabled = False
'        txtShopItemNumber(x).Enabled = False
'        txtShopMax(x).Enabled = False
'        txtShopNormal(x).Enabled = False
'        txtShopRgnTime(x).Enabled = False
'        txtRgnNumber(x).Enabled = False
'        txtRgnPercentage(x).Enabled = False
'        txtCost(x).Enabled = False
'    Next
ElseIf cmbType.ListIndex = 11 Then 'gangshop
    lblColumns(3).Caption = "Acct"
    lblColumns(4).Caption = "Cost"
    lblColumns(5).Caption = ""
    lblColumns(6).Caption = "Coin"
    txtBankAcct.Enabled = True
    txtBankAcct.Text = txtBankAcct.Tag
    
    For x = 0 To 9
        txtShopMax(x).Enabled = False
        txtRgnPercentage(x).Enabled = False
    Next
    For x = 10 To 19
        txtCost(x).Enabled = False
        txtShopItemName(x).Enabled = False
        txtShopItemNumber(x).Enabled = False
        txtShopMax(x).Enabled = False
        txtShopNormal(x).Enabled = False
        txtShopRgnTime(x).Enabled = False
        txtRgnNumber(x).Enabled = False
        txtRgnPercentage(x).Enabled = False
    Next
Else
    For x = 0 To 19
        txtCost(x).Enabled = True
        txtShopItemNumber(x).Enabled = True
        txtShopItemName(x).Enabled = True
        txtShopMax(x).Enabled = True
        txtShopNormal(x).Enabled = True
        txtShopRgnTime(x).Enabled = True
        txtRgnNumber(x).Enabled = True
        txtRgnPercentage(x).Enabled = True
    Next
End If

Call txtShopMarkup_Change

Exit Sub
Error:
Call HandleError("cmbType_Click")
End Sub

Private Sub cmdItemLookup_Click(Index As Integer)
Call frmItem.GotoItem(Val(txtShopItemNumber(Index).Text))
frmItem.Show
frmItem.SetFocus

End Sub

Private Sub Form_Load()
Dim sCaption As String, i As Integer
On Error Resume Next

bLoaded = False

sCaption = frmMain.Caption
frmMain.Caption = sCaption & " - Loading Shops ..."
DoEvents

With EL1
    .FormInQuestion = Me
    .MINHEIGHT = 480 + (TITLEBAR_OFFSET / 10)
    .MINWIDTH = 665
    .CenterOnLoad = False
    .EnableLimiter = True
End With

Me.Top = ReadINI("Windows", "ShopTop")
Me.Left = ReadINI("Windows", "ShopLeft")
Me.Width = ReadINI("Windows", "ShopWidth")
Me.Height = ReadINI("Windows", "ShopHeight")

lvDatabase.ListItems.clear

For i = 0 To UBound(Classes)
    cmbClasses.AddItem Classes(i).Name
Next i

Call LoadShops

Me.Show
Me.SetFocus
txtSearch.SetFocus
frmMain.Caption = sCaption
If ReadINI("Windows", "ShopMaxed") = "1" Then Me.WindowState = vbMaximized
End Sub
Private Sub cmdDiscard_Click()
Dim nStatus As Integer

On Error GoTo Error:

If lvDatabase.SelectedItem Is Nothing Or nCurrentRecord = 0 Then
    MsgBox "No current record."
    Exit Sub
End If

nStatus = BTRCALL(BGETEQUAL, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), nCurrentRecord, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error on BGETEQUAL: " & BtrieveErrorCode(nStatus)
Else
    DispShopInfo Shopdatabuf.buf
End If

Exit Sub
Error:
Call HandleError("cmdDiscard_Click")
End Sub
Public Sub GotoShop(ByVal nRecnum As Integer)
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
Private Sub cmdSave_Click()
On Error GoTo Error:

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
If lvDatabase.SelectedItem Is Nothing Then Exit Sub

Call saverecord(nCurrentRecord)
'Call lvDatabase_ItemClick(lvDatabase.SelectedItem)

Dim oLI As ListItem
Set oLI = lvDatabase.FindItem(Shoprec.Number, lvwText, , 0)
If Not oLI Is Nothing Then
    oLI.ListSubItems(1).Text = ClipNull(Shoprec.Name)
    If Not bOnlyNames Then
        oLI.ListSubItems(2).Text = GetShopType(Shoprec.ShopType)
        oLI.ListSubItems(3).Text = Shoprec.ShopMinLvL & "-" & Shoprec.ShopMaxLvl
        oLI.ListSubItems(4).Text = Shoprec.ShopMarkUp & "%"
    End If
End If
Set oLI = Nothing

Exit Sub
Error:
Call HandleError("cmdSave_Click")
End Sub

Private Sub LoadShops()
Dim nStatus As Integer

lvDatabase.ColumnHeaders.clear
lvDatabase.ColumnHeaders.add 1, "Number", "#", 600, lvwColumnLeft
lvDatabase.ColumnHeaders.add 2, "Name", "Name", 1900, lvwColumnCenter
If Not bOnlyNames Then
    lvDatabase.ColumnHeaders.add 3, "Type", "Type", 1000, lvwColumnCenter
    lvDatabase.ColumnHeaders.add 4, "Level", "Level", 800, lvwColumnCenter
    lvDatabase.ColumnHeaders.add 5, "Markup", "Markup", 800, lvwColumnCenter
End If

nStatus = BTRCALL(BGETFIRST, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "LoadShops, BGETFIRST, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

Do While nStatus = 0
    ShopRowToStruct Shopdatabuf.buf
    
    Call AddShopToLV(Shoprec.Number)
    
    nStatus = BTRCALL(BGETNEXT, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "LoadShops, Error: " & BtrieveErrorCode(nStatus)
End If

If lvDatabase.ListItems.Count >= 1 Then Call lvDatabase_ItemClick(lvDatabase.ListItems(1))

lvDatabase.refresh
SortListView lvDatabase, 1, ldtNumber, True
bLoaded = True

Exit Sub
Error:
Call HandleError
End Sub
Private Sub AddShopToLV(ByVal nNumber As Integer)
Dim nStatus As Integer, oLI As ListItem
On Error GoTo Error:

If Not nNumber = Shoprec.Number Then
    nStatus = BTRCALL(BGETEQUAL, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), nNumber, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then MsgBox "Error getting record " & nNumber & ": " & BtrieveErrorCode(nStatus)
    bLoaded = False
    Exit Sub
End If

Set oLI = lvDatabase.ListItems.add()
oLI.Text = Shoprec.Number

oLI.ListSubItems.add (1), "Name", ClipNull(Shoprec.Name)
If Not bOnlyNames Then
    oLI.ListSubItems.add (2), "Type", GetShopType(Shoprec.ShopType)
    oLI.ListSubItems.add (3), "Level", Shoprec.ShopMinLvL & "-" & Shoprec.ShopMaxLvl
    oLI.ListSubItems.add (4), "Markup", Shoprec.ShopMarkUp & "%"
End If

Set oLI = Nothing
Exit Sub
Error:
Call HandleError
Set oLI = Nothing
End Sub

Private Sub DispShopInfo(row() As Byte)
On Error GoTo Error:
Dim x As Integer, sBankAccount As String, sChars As String, i As Integer

Call ShopRowToStruct(row())

Me.Caption = "Shop Editor -- " & ClipNull(Shoprec.Name)

'needs to come before the Type is changed
For x = 0 To 19
    sChars = Hex(Shoprec.ShopMax(x))
    If Len(sChars) < 3 Then
        sChars = String(4 - Len(sChars), "0") & sChars
    ElseIf Len(sChars) < 4 Then
        sChars = sChars & String(4 - Len(sChars), "0")
    End If
    sBankAccount = sBankAccount & Chr(Val("&H" & Right(sChars, 2))) & Chr(Val("&H" & Left(sChars, 2)))
    If Shoprec.ShopType = 11 Then
        txtShopMax(x).Text = Chr(Val("&H" & Right(sChars, 2))) & Chr(Val("&H" & Left(sChars, 2)))
    End If
Next x
If Shoprec.ShopType = 11 Then
    txtBankAcct.Text = sBankAccount
    txtBankAcct.Tag = sBankAccount
Else
    txtBankAcct.Tag = sBankAccount
    txtBankAcct.Text = "<not a gangshop>"
End If

txtName.Text = Shoprec.Name
cmbType.ListIndex = Shoprec.ShopType
txtShopMinLvl.Text = Shoprec.ShopMinLvL
txtShopMaxLvl.Text = Shoprec.ShopMaxLvl

For x = 0 To 19
    txtShopItemNumber(x).Text = Shoprec.ShopItemNumber(x)
    'txtShopItemName(x).Text = GetItemName(Shoprec.ShopItemNumber(x))
    If Not Shoprec.ShopType = 11 Then
        txtShopMax(x).Text = Shoprec.ShopMax(x)
    End If
    txtShopNormal(x).Text = Shoprec.ShopNow(x)
    txtShopRgnTime(x).Text = Shoprec.ShopRgnTime(x)
    txtRgnNumber(x).Text = Shoprec.ShopRgnNumber(x)
    txtRgnPercentage(x).Text = Shoprec.ShopRgnPercentage(x)
Next

If cmbClasses.ListCount <= Shoprec.ShopClassLimit Then
    Call Add2ClassArray(CInt(Shoprec.ShopClassLimit))
    cmbClasses.clear
    For i = 0 To UBound(Classes)
        cmbClasses.AddItem Classes(i).Name
    Next i
End If

cmbClasses.ListIndex = Shoprec.ShopClassLimit
txtShopMarkup.Text = Shoprec.ShopMarkUp

Exit Sub
Error:
Call HandleError
MsgBox "Warning, record was not completely displayed." & vbCrLf _
    & "Previous records stats may still be in memory.  Select 'Disable DB Writing'" & vbCrLf _
    & "from the file menu and then reload the editor.", vbExclamation
End Sub

Private Sub cmdOther_Click()
If frameGeneral.Visible = False Then
    frameGeneral.Visible = True
    cmdOther.Caption = "Ite&ms Sold"
Else
    frameGeneral.Visible = False
    cmdOther.Caption = "&General Info."
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = vbMinimized Then Exit Sub
framNav.Left = Me.Width - framNav.Width - 200
lvDatabase.Width = framNav.Left - 175
lvDatabase.Height = Me.Height - 925 - TITLEBAR_OFFSET
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If bLoaded = True Then Call saverecord(nCurrentRecord)
If Me.WindowState = vbMinimized Then Exit Sub

If Me.WindowState = vbMaximized Then
    Call WriteINI("Windows", "ShopMaxed", 1)
Else
    Call WriteINI("Windows", "ShopMaxed", 0)
    Call WriteINI("Windows", "ShopTop", Me.Top)
    Call WriteINI("Windows", "ShopLeft", Me.Left)
    Call WriteINI("Windows", "ShopHeight", Me.Height)
    Call WriteINI("Windows", "ShopWidth", Me.Width)
End If
End Sub

Private Sub saverecord(ByVal nRecord As Long)
On Error GoTo Error:
Dim nStatus As Integer, x As Integer, sChar As String, sChar2 As String

If nRecord = 0 Then Exit Sub
nStatus = BTRCALL(BGETEQUAL, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), nRecord, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error on BGETEQUAL: " & BtrieveErrorCode(nStatus)
    Exit Sub
Else
    ShopRowToStruct Shopdatabuf.buf
End If

'DoEvents
Shoprec.Name = txtName.Text & Chr(0)

Shoprec.ShopType = cmbType.ListIndex
Shoprec.ShopMinLvL = Val(txtShopMinLvl.Text)
Shoprec.ShopMaxLvl = Val(txtShopMaxLvl.Text)
Shoprec.ShopMarkUp = Val(txtShopMarkup.Text)

For x = 0 To 19
    Shoprec.ShopItemNumber(x) = Val(txtShopItemNumber(x).Text)
    
    If Not cmbType.ListIndex = 11 Then
        Shoprec.ShopMax(x) = Val(txtShopMax(x).Text)
    End If
    
    Shoprec.ShopNow(x) = Val(txtShopNormal(x).Text)
    Shoprec.ShopRgnTime(x) = Val(txtShopRgnTime(x).Text)
    Shoprec.ShopRgnNumber(x) = Val(txtRgnNumber(x).Text)
    Shoprec.ShopRgnPercentage(x) = Val(txtRgnPercentage(x).Text)
Next x

If cmbType.ListIndex = 11 Then
    For x = 1 To Len(txtBankAcct.Text)
        sChar = Asc(Mid(txtBankAcct.Text, x, 1))
        If Not sChar2 = "" Then
            Shoprec.ShopMax((x / 2) - 1) = Val("&H" & Hex(sChar) & Hex(sChar2))
            sChar2 = ""
        Else
            sChar2 = sChar
        End If
    Next x
    If Not sChar2 = "" Then
        Shoprec.ShopMax((x / 2) - 1) = Val("&H" & Hex(sChar2))
        sChar2 = ""
        'x = x + 1
    End If
    x = Fix(x / 2)
    For x = x To 19
        Shoprec.ShopMax(x) = 0
    Next x
End If
    
Shoprec.ShopClassLimit = cmbClasses.ListIndex

nStatus = UpdateShop
If Not nStatus = 0 Then
    MsgBox "SaveRecord, Error: " & BtrieveErrorCode(nStatus)
Else
    DispShopInfo Shopdatabuf.buf
End If
Exit Sub
Error:
Call HandleError
End Sub

Private Sub lvDatabase_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim nSort As ListDataType
On Error GoTo Error:

Select Case ColumnHeader.Index
    Case 1, 4, 5: nSort = ldtNumber
    Case Else: nSort = ldtString
End Select
SortListView lvDatabase, ColumnHeader.Index, nSort, lvDatabase.SortOrder

Exit Sub
Error:
Call HandleError("lvDatabase_ColumnClick")
End Sub

Public Sub lvDatabase_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim temp As Long, nStatus As Integer
On Error GoTo Error:

If bLoaded = True And chkAutoSave.Value = 1 Then Call saverecord(nCurrentRecord)

temp = Val(Item.Text)
nStatus = BTRCALL(BGETEQUAL, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), temp, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error on BGETEQUAL: " & BtrieveErrorCode(nStatus)
    bLoaded = False
Else
    nCurrentRecord = temp
    DispShopInfo Shopdatabuf.buf
    bLoaded = True
End If

Exit Sub
Error:
Call HandleError("lvDatabase_ItemClick")
End Sub

Private Sub txtBankAcct_GotFocus()
Call SelectAll(txtBankAcct)
End Sub

Private Sub txtName_GotFocus()
Call SelectAll(txtName)
End Sub

Private Sub txtNumberSearch_GotFocus()
Call SelectAll(txtNumberSearch)

End Sub

Private Sub txtNumberSearch_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)

End Sub

Private Sub txtNumberSearch_KeyUp(KeyCode As Integer, Shift As Integer)
Dim x As Long, SearchStart As Long

On Error GoTo Error:

If txtNumberSearch.Text = "" Then Exit Sub
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
    If Val(lvDatabase.ListItems(x).Text) = Val(txtNumberSearch.Text) Then
        Set lvDatabase.SelectedItem = lvDatabase.ListItems(x)
        lvDatabase.SelectedItem.EnsureVisible
        Call lvDatabase_ItemClick(lvDatabase.SelectedItem)
        Exit For
    End If
Next x

Exit Sub
Error:
Call HandleError("txtNumberSearch_KeyUp")

End Sub



Private Sub txtRgnNumber_Change(Index As Integer)
If cmbType.ListIndex = 11 Then
    txtCost(Index).Text = CalcMarkup(Val(txtShopRgnTime(Index).Text), Val(txtShopMarkup.Text)) _
            & " " & GetCostType(Val(txtRgnNumber(Index).Text))
End If
End Sub

Private Sub txtSearch_GotFocus()
Call SelectAll(txtSearch)

End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
Dim x As Long, SearchStart As Long

On Error GoTo Error:

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

Exit Sub
Error:
Call HandleError("txtSearch_KeyUp")
    
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

nStatus = BTRCALL(BGETEQUAL, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), nCurrentRecord, KEY_BUF_LEN, 0)
If nStatus = 0 Then
    nStatus = BTRCALL(BDELETE, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
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
Error:
Call HandleError
End Sub

Private Sub cmdInsert_Click()
On Error GoTo Error:
Dim nStatus As Integer
Dim nNewShopNumber As String, oLI As ListItem

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

If bLoaded = True Then Call saverecord(nCurrentRecord)

nNewShopNumber = InputBox("New Shop Number:" & vbCrLf & vbCrLf & "Enter 0 for the next highest number.", "Insert", "0")
If nNewShopNumber = "" Then Exit Sub

Shoprec.Number = Val(nNewShopNumber)
'Shoprec.Name = "New Shop" & Chr(0)
Call ShopStructToRow(Shopdatabuf.buf)

nStatus = BTRCALL(BINSERT, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), ByVal ShopKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdInsert, BINSERT, Error: " & BtrieveErrorCode(nStatus)
    bLoaded = False
Else
    ShopRowToStruct Shopdatabuf.buf
    
    Call AddShopToLV(Shoprec.Number)
    
    nCurrentRecord = Shoprec.Number
    DispShopInfo Shopdatabuf.buf
    
    SortListView lvDatabase, 1, ldtNumber, True
    
    Set oLI = lvDatabase.FindItem(Shoprec.Number, lvwText, , 0)
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
Error:
Call HandleError
Set oLI = Nothing
End Sub


Private Sub txtShopItemNumber_Change(Index As Integer)
On Error GoTo Error:

txtShopItemName(Index).Text = GetItemName(Val(txtShopItemNumber(Index).Text))

If cmbType.ListIndex = 11 Then
    txtCost(Index).Text = CalcMarkup(Val(txtShopRgnTime(Index).Text), Val(txtShopMarkup.Text)) _
            & " " & GetCostType(Val(txtRgnNumber(Index).Text))
Else
    txtCost(Index).Text = GetItemCost(Val(txtShopItemNumber(Index).Text), Val(txtShopMarkup.Text))
End If

Exit Sub
Error:
Call HandleError("txtShopItemNumber_Change")
End Sub

Private Sub txtShopItemNumber_GotFocus(Index As Integer)
Call SelectAll(txtShopItemNumber(Index))
End Sub

Private Sub txtShopMarkup_Change()
Dim x As Integer

On Error GoTo Error:

If cmbType.ListIndex = 11 Then 'gangshop
    For x = 0 To 19
        txtCost(x).Text = CalcMarkup(Val(txtShopRgnTime(x).Text), Val(txtShopMarkup.Text)) _
            & " " & GetCostType(Val(txtRgnNumber(x).Text))
    Next
Else
    For x = 0 To 19
        txtCost(x).Text = GetItemCost(Val(txtShopItemNumber(x).Text), Val(txtShopMarkup.Text))
    Next
End If

out:
Exit Sub
Error:
Call HandleError("txtShopMarkup_Change")
Resume out:

End Sub

Private Sub txtShopMarkup_GotFocus()
Call SelectAll(txtShopMarkup)
End Sub

Private Sub txtShopMax_GotFocus(Index As Integer)
Call SelectAll(txtShopMax(Index))
End Sub

Private Sub txtShopMaxLvl_GotFocus()
Call SelectAll(txtShopMaxLvl)
End Sub

Private Sub txtShopMinLvl_GotFocus()
Call SelectAll(txtShopMinLvl)
End Sub

Private Sub txtShopNormal_GotFocus(Index As Integer)
Call SelectAll(txtShopNormal(Index))
End Sub

Private Sub txtShopRgnTime_Change(Index As Integer)
If cmbType.ListIndex = 11 Then
    txtCost(Index).Text = CalcMarkup(Val(txtShopRgnTime(Index).Text), Val(txtShopMarkup.Text)) _
            & " " & GetCostType(Val(txtRgnNumber(Index).Text))
End If
End Sub

Private Sub txtShopRgnTime_GotFocus(Index As Integer)
Call SelectAll(txtShopRgnTime(Index))
End Sub
Private Sub txtRgnNumber_GotFocus(Index As Integer)
Call SelectAll(txtRgnNumber(Index))
End Sub

Private Sub txtRgnPercentage_GotFocus(Index As Integer)
Call SelectAll(txtRgnPercentage(Index))
End Sub
