VERSION 5.00
Begin VB.Form frmMonsterItemDropPct 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monster Item Drop Percentage Modifier"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2160
   ScaleWidth      =   7935
   Begin VB.CheckBox chkExcludeDelMain 
      Caption         =   "Exclude items with abil-119 Del@maint"
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
      Left            =   3660
      TabIndex        =   19
      Top             =   1140
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.TextBox txtMonRegen 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2880
      TabIndex        =   18
      Text            =   "1"
      Top             =   1800
      Width           =   615
   End
   Begin VB.CheckBox chkMonRegen 
      Caption         =   "AND only if monster regen > "
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
      Left            =   120
      TabIndex        =   17
      Top             =   1800
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.ComboBox cmbOnlyIfModifier 
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
      ItemData        =   "frmMonItemDropPct.frx":0000
      Left            =   60
      List            =   "frmMonItemDropPct.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1080
      Width           =   1035
   End
   Begin VB.ComboBox cmbOption 
      Height          =   315
      ItemData        =   "frmMonItemDropPct.frx":002E
      Left            =   60
      List            =   "frmMonItemDropPct.frx":0038
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Start"
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
      Left            =   6660
      TabIndex        =   11
      Top             =   1140
      Width           =   1155
   End
   Begin VB.CheckBox chkLogOnly 
      Caption         =   "Log Only - No Changes"
      Height          =   255
      Left            =   5820
      TabIndex        =   10
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1995
   End
   Begin VB.CheckBox chkItemTypes 
      Caption         =   "Scrolls"
      Height          =   255
      Index           =   4
      Left            =   4860
      TabIndex        =   9
      Top             =   780
      Value           =   1  'Checked
      Width           =   915
   End
   Begin VB.CheckBox chkItemTypes 
      Caption         =   "Containers"
      Height          =   255
      Index           =   3
      Left            =   4860
      TabIndex        =   8
      Top             =   480
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox chkItemTypes 
      Caption         =   "Keys"
      Height          =   255
      Index           =   2
      Left            =   6060
      TabIndex        =   6
      Top             =   480
      Value           =   1  'Checked
      Width           =   795
   End
   Begin VB.CheckBox chkItemTypes 
      Caption         =   "Armour "
      Height          =   255
      Index           =   1
      Left            =   3660
      TabIndex        =   5
      Top             =   780
      Value           =   1  'Checked
      Width           =   915
   End
   Begin VB.CheckBox chkItemTypes 
      Caption         =   "Weapons"
      Height          =   255
      Index           =   0
      Left            =   3660
      TabIndex        =   3
      Top             =   480
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.TextBox txtPctLimit 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "20"
      Top             =   1080
      Width           =   795
   End
   Begin VB.TextBox txtDropPct 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2580
      TabIndex        =   0
      Text            =   "20"
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Value"
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
      Left            =   2580
      TabIndex        =   16
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Choose option..."
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
      Left            =   60
      TabIndex        =   15
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   3660
      TabIndex        =   12
      Top             =   1800
      Width           =   2115
   End
   Begin VB.Label lblExtra 
      Caption         =   "(AND if % is greater than 0)"
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
      Left            =   120
      TabIndex        =   7
      Top             =   1500
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Do it for these item types:"
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
      Left            =   3660
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Condition: Only if drop % is..."
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
      Left            =   60
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
End
Attribute VB_Name = "frmMonsterItemDropPct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Private Sub cmbOnlyIfModifier_Click()
'lblExtra.Caption = ""
'Select Case cmbOnlyIfModifier.ListIndex
'    Case 0: '=
'    Case 1: ' <=
'        lblExtra.Caption = "(AND greater than 0)"
'    Case 2: ' >=
'    Case 3: ' !=
'End Select
End Sub

Private Sub cmbOption_Click()
If cmbOption.ListIndex = 1 Then
    MsgBox "Put in a positive or negative value in the box.  The result will not be set to less than 1 or greater than 100", vbOKOnly + vbInformation
End If
End Sub

Private Sub cmdGo_Click()
Dim fso As FileSystemObject, fil As String, ts As TextStream
Dim nStatus As Integer, nStatus2 As Integer, recnum As Long, x As Long, sTemp As String, y As Integer, z As Integer
Dim bLogOnly As Boolean, nValue As Integer, nLimit As Integer, bMatch As Boolean, nMonRegen As Integer
Dim bIsWeapon As Boolean, bIsArmour As Boolean, bIsContainer As Boolean, bIsScroll As Boolean, bIsKey As Boolean
                
On Error GoTo error:

If Val(txtPctLimit.Text) > 100 Or Val(txtPctLimit.Text) < 0 Then
    MsgBox "Error with limit value.", vbOKOnly + vbExclamation, "Error"
    Exit Sub
End If

If Val(txtDropPct.Text) > 100 Or (cmbOption.ListIndex = 0 And Val(txtDropPct.Text) < 0) _
    Or (cmbOption.ListIndex = 1 And Val(txtDropPct.Text) = 0) Then
    
    MsgBox "Error with set value.", vbOKOnly + vbExclamation, "Error"
    Exit Sub

End If

nLimit = Val(txtPctLimit.Text)
nValue = Val(txtDropPct.Text)
nMonRegen = Val(txtMonRegen.Text)

If nValue = 0 And cmbOption.ListIndex = 0 Then
    nStatus = MsgBox("Current settings will result in setting matched item drop pct to 0 and then will not drop.  This may be hard to undo.  Are you sure you want to continue?", vbYesNo + vbDefaultButton2 + vbQuestion, "Continue?")
    If nStatus <> vbYes Then Exit Sub
End If

recnum = 1
nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Monsters: Could not get first record, Error: " & BtrieveErrorCode(nStatus), vbExclamation + vbOKOnly
    Exit Sub
End If

Set fso = CreateObject("Scripting.FileSystemObject")
If Right(App.Path, 1) = "\" Then
    fil = App.Path & "NMR-Log_MonItemPct.txt"
Else
    fil = App.Path & "\NMR-Log_MonItemPct.txt"
End If
If fso.FileExists(fil) = True Then fso.DeleteFile fil, True
Set ts = fso.OpenTextFile(fil, ForWriting, True)

ts.WriteLine ("Monster Item Drop Percentage Modifier " & Date & " @ " & Time)

If chkLogOnly.Value = 1 Then bLogOnly = True
If bLogOnly Then ts.WriteLine ("** LOGGING ONLY, NO CHANGES EXECUTED **")

ts.WriteBlankLines (1)

Me.Enabled = False
frmMain.Enabled = False

lblStatus.Caption = recnum

Do While nStatus = 0
    
    RowToStruct Monsterdatabuf.buf, MonsterFldMap, Monsterrec, LenB(Monsterrec)
    
    recnum = Monsterrec.Number
    lblStatus.Caption = recnum
    
    If chkMonRegen.Value = 1 And Monsterrec.RegenTime <= nMonRegen Then GoTo skip_mon:
    
    For x = 0 To 9
        bMatch = False
        bIsWeapon = False
        bIsArmour = False
        bIsContainer = False
        bIsScroll = False
        bIsKey = False
        
        If Monsterrec.ItemNumber(x) > 0 And Monsterrec.ItemDropPer(x) > 0 Then
            Select Case cmbOnlyIfModifier.ListIndex
                Case 0: ' =
                    If Monsterrec.ItemDropPer(x) = nLimit Then bMatch = True
                Case 1: ' <=
                    If Monsterrec.ItemDropPer(x) <= nLimit Then bMatch = True
                Case 2: ' >=
                    If Monsterrec.ItemDropPer(x) >= nLimit Then bMatch = True
                Case 3: ' !=
                    If Not Monsterrec.ItemDropPer(x) = nLimit Then bMatch = True
            End Select
        End If
        
        If Monsterrec.ItemDropPer(x) = nValue Then bMatch = False
        
        If bMatch Then
            
            nStatus2 = BTRCALL(BGETEQUAL, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), Monsterrec.ItemNumber(x), Len(Monsterrec.ItemNumber(x)), 0)
            If Not nStatus2 = 0 Then
                ts.WriteLine ("ERROR: Monster #" & Monsterrec.Number & " - " & ClipNull(Monsterrec.Name) & _
                    ": Item #" & Monsterrec.ItemNumber(x) & " - ITEM NOT FOUND")
                bMatch = False
            Else
                ItemRowToStruct Itemdatabuf.buf
                
                Select Case Itemrec.Type
                    Case 0: 'armour
                        If Itemrec.WornOn = 0 Then
                            'nowhere
                            bMatch = False
                        Else
                            bIsArmour = True
                        End If
                        
                    Case 1: 'weapons
                        bIsWeapon = True
                        
                    Case 2: '"Projectile"
                        bMatch = False
                    Case 3: '"Sign"
                        bMatch = False
                    Case 4: '"Food"
                        bMatch = False
                    Case 5: '"Drink"
                        bMatch = False
                    Case 6: '"Light"
                        bMatch = False
                    Case 7: '"Key"
                        bIsKey = True
                    Case 8: '"Container"
                        bIsContainer = True
                    Case 9: '"Scroll"
                        bIsScroll = True
                    Case 10: '"Special"
                        bMatch = False
                    Case Else:
                        bMatch = False
                End Select
                
                If bMatch Then
                    If chkItemTypes(0).Value = 0 And bIsWeapon Then bMatch = False
                    If chkItemTypes(1).Value = 0 And bIsArmour Then bMatch = False
                    If chkItemTypes(2).Value = 0 And bIsKey Then bMatch = False
                    If chkItemTypes(3).Value = 0 And bIsContainer Then bMatch = False
                    If chkItemTypes(4).Value = 0 And bIsScroll Then bMatch = False
                End If
                
                If bMatch And (bIsWeapon Or bIsArmour) And chkExcludeDelMain.Value = 1 Then
                    For y = 0 To 19
                        If Itemrec.AbilityA(y) = 119 Then
                            bMatch = False
                            Exit For
                        End If
                    Next y
                End If
            End If
        End If
        
        If bMatch Then
            sTemp = "Monster #" & Monsterrec.Number & " - " & ClipNull(Monsterrec.Name) & ": Item #" & _
                Monsterrec.ItemNumber(x) & " - " & GetItemName(Monsterrec.ItemNumber(x)) & _
                ": Drop % changed from " & Monsterrec.ItemDropPer(x) & " to "
            
            If cmbOption.ListIndex = 0 Then 'set to
                sTemp = sTemp & nValue
                z = nValue
            Else
                z = Monsterrec.ItemDropPer(x) + nValue
                If z < 1 Then z = 1
                If z > 100 Then z = 100
                
                sTemp = sTemp & z
            End If
            
            If bLogOnly Then
                sTemp = "LOG_ONLY: " & sTemp
                ts.WriteLine (sTemp)
            Else
                Monsterrec.ItemDropPer(x) = z
                nStatus = UpdateMonster
                If Not nStatus = 0 Then
                    ts.WriteLine (sTemp)
                    ts.WriteLine ("ERROR Updating Monster Record # " & Monsterrec.Number & " - " & BtrieveErrorCode(nStatus))
                Else
                    ts.WriteLine (sTemp)
                End If
            End If
            
        End If
    Next
    
skip_mon:
    nStatus = BTRCALL(BGETNEXT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)

    If Not bUseCPU Then DoEvents
Loop

ts.WriteBlankLines (1)
ts.WriteLine ("Finished - " & Date & " @ " & Time)
ts.Close

nStatus = MsgBox("Finished. Open Log?", vbYesNo + vbDefaultButton1 + vbQuestion, "Complete")
If nStatus = vbYes Then
    If fso.FileExists(fil) = True Then
        Call ShellExecute(0&, "open", fil, vbNullString, vbNullString, vbNormalFocus)
    Else
        MsgBox fil & " was not found.", vbInformation
    End If
End If

out:
On Error Resume Next
Me.Enabled = True
frmMain.Enabled = True
ts.Close
Set ts = Nothing
Exit Sub
error:
Call HandleError("cmdGo_Click")
Resume out: End Sub

Private Sub Form_Load()
On Error Resume Next
Dim nStatus As Integer

Me.Top = ReadINI("Windows", "MonItemPctTop")
Me.Left = ReadINI("Windows", "MonItemPctLeft")
cmbOption.ListIndex = 0
cmbOnlyIfModifier.ListIndex = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
        If Me.WindowState = vbMinimized Then Exit Sub
        Call WriteINI("Windows", "MonItemPctTop", frmMonsterItemDropPct.Top)
        Call WriteINI("Windows", "MonItemPctLeft", frmMonsterItemDropPct.Left)
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtDropPct_Change()
If txtDropPct.Text = "" Or txtDropPct.Text = " " Then Exit Sub
If Val(txtDropPct.Text) > 100 Then txtDropPct.Text = 100
If Val(txtDropPct.Text) < -100 Then txtDropPct.Text = -100
End Sub

Private Sub txtMonRegen_Change()
If txtMonRegen.Text = "" Or txtMonRegen.Text = " " Then Exit Sub
If Val(txtMonRegen.Text) > 100 Then txtMonRegen.Text = 100
If Val(txtMonRegen.Text) < 1 Then txtMonRegen.Text = 0
End Sub

Private Sub txtMonRegen_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtPctLimit_Change()
If txtPctLimit.Text = "" Or txtPctLimit.Text = " " Then Exit Sub
If Val(txtPctLimit.Text) > 100 Then txtPctLimit.Text = 100
If Val(txtPctLimit.Text) < 1 Then txtPctLimit.Text = 0
End Sub

Private Sub txtPctLimit_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtDropPct_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub
