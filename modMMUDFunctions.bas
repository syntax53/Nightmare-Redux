Attribute VB_Name = "modMMUDFunctions"
Option Explicit
Option Base 0

Public Type RoomExitType
    Map As Long
    Room As Long
    ExitType As String
End Type

Public Function ExtractMapRoom(ByVal sExit As String) As RoomExitType
Dim x As Integer, y As Integer, i As Integer

On Error GoTo error:

ExtractMapRoom.Map = 0
ExtractMapRoom.Room = 0
ExtractMapRoom.ExitType = 0

x = InStr(1, sExit, "/")
Do While x - 1 > 0 'gets where the map number starts
    Select Case Mid(sExit, x - 1, 1)
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0":
            i = x - 1
        Case Else:
            Exit Do
    End Select
    x = x - 1
Loop

'For i = 1 To Len(sExit) - 1 'gets where the first number is
'    Select Case Mid(sExit, i, 1)
'        Case "1", "2", "3", "4", "5", "6", "7", "8", "9": Exit For
'    End Select
'Next

x = InStr(1, sExit, "/")
If x = 0 Then Exit Function
If x = Len(sExit) Then Exit Function

ExtractMapRoom.Map = Val(Mid(sExit, i, x - 1))

y = InStr(x, sExit, " ")
If y = 0 Then
    ExtractMapRoom.Room = Val(Mid(sExit, x + 1))
Else
    ExtractMapRoom.Room = Val(Mid(sExit, x + 1, y - 1))
    ExtractMapRoom.ExitType = Mid(sExit, y + 1)
End If

Exit Function

error:
Call HandleError("ExtractMapRoom")

End Function

Public Function GetMonType(ByVal Index As Integer) As String

On Error GoTo error:

Select Case Index
    Case 0: GetMonType = GetMonType & "Lair"
    Case 1: GetMonType = GetMonType & "Wanderer"
    Case 2: GetMonType = GetMonType & "NPC"
    Case 3: GetMonType = GetMonType & "Living"
    Case 4: GetMonType = GetMonType & "Random"
    Case 5: GetMonType = GetMonType & "Guard"
    Case Is >= 6, Is <= 35: GetMonType = GetMonType & "Group " & Index - 5
    Case 36: GetMonType = GetMonType & "Arena"
    Case 37: GetMonType = GetMonType & "Angel"
    Case 38: GetMonType = GetMonType & "Quest"
    Case 39: GetMonType = GetMonType & "Other"
    Case Else: GetMonType = GetMonType & "unknown"
End Select

out:
Exit Function
error:
Call HandleError("GetMonType")
Resume out:

End Function
Public Function GetRoomType(ByVal Index As Integer) As String

On Error GoTo error:

Select Case Index
    Case 0: GetRoomType = "Normal"
    Case 1: GetRoomType = "Shop"
    Case 2: GetRoomType = "Arena"
    Case 3: GetRoomType = "Lair"
    Case 4: GetRoomType = "Hotel"
    Case 5: GetRoomType = "Colliseum"
    Case 6: GetRoomType = "Jail"
    Case 7: GetRoomType = "Library"
    Case Else: GetRoomType = "unknown"
End Select

out:
Exit Function
error:
Call HandleError("GetRoomType")
Resume out:

End Function

Public Function GetRoomExitType(ByVal nNum As Integer) As String
On Error GoTo error:

Select Case nNum
    Case 0: GetRoomExitType = "Normal"
    Case 1: GetRoomExitType = "Spell"
    Case 2: GetRoomExitType = "Key"
    Case 3: GetRoomExitType = "Item"
    Case 4: GetRoomExitType = "Toll"
    Case 5: GetRoomExitType = "Action"
    Case 6: GetRoomExitType = "Hidden"
    Case 7: GetRoomExitType = "Door"
    Case 8: GetRoomExitType = "map Change"
    Case 9: GetRoomExitType = "Trap"
    Case 10: GetRoomExitType = "Text"
    Case 11: GetRoomExitType = "Gate"
    Case 12: GetRoomExitType = "Remote Action"
    Case 13: GetRoomExitType = "Class"
    Case 14: GetRoomExitType = "Race"
    Case 15: GetRoomExitType = "Level"
    Case 16: GetRoomExitType = "Timed"
    Case 17: GetRoomExitType = "Ticket"
    Case 18: GetRoomExitType = "User Count"
    Case 19: GetRoomExitType = "Block Guard"
    Case 20: GetRoomExitType = "Alignment"
    Case 21: GetRoomExitType = "Delay"
    Case 22: GetRoomExitType = "Cast"
    Case 23: GetRoomExitType = "Ability"
    Case 24: GetRoomExitType = "Spell Trap"
    Case Else:
        GetRoomExitType = "Unknown (" & nNum & ")"
End Select

out:
Exit Function
error:
Call HandleError("GetRoomExitType")
Resume out:
End Function
Public Function GetShopType(ByVal nNum As Long) As String
On Error GoTo error:

Select Case nNum
    Case 0: GetShopType = "General"
    Case 1: GetShopType = "Weapons"
    Case 2: GetShopType = "Armour"
    Case 3: GetShopType = "Items"
    Case 4: GetShopType = "Spells"
    Case 5: GetShopType = "Hospital"
    Case 6: GetShopType = "Tavern"
    Case 7: GetShopType = "Bank"
    Case 8: GetShopType = "Training"
    Case 9: GetShopType = "Inn"
    Case 10: GetShopType = "Specific"
    Case 11: GetShopType = "Gang Shop"
    Case 12: GetShopType = "Deed Shop"
    Case Else: GetShopType = "Unknown (" & nNum & ")"
End Select

out:
Exit Function
error:
Call HandleError("GetShopType")
Resume out:
End Function
Public Function GetItemType(ByVal ItemType As Integer) As String
On Error GoTo error:

Select Case ItemType
    Case 0: GetItemType = "Armour"
    Case 1: GetItemType = "Weapon"
    Case 2: GetItemType = "Projectile"
    Case 3: GetItemType = "Sign"
    Case 4: GetItemType = "Food"
    Case 5: GetItemType = "Drink"
    Case 6: GetItemType = "Light"
    Case 7: GetItemType = "Key"
    Case 8: GetItemType = "Container"
    Case 9: GetItemType = "Scroll"
    Case 10: GetItemType = "Special"
    
    Case Else: GetItemType = ItemType
End Select

out:
Exit Function
error:
Call HandleError("GetItemType")
Resume out:
End Function
Public Function GetWornType(ByVal nNum As Integer) As String
On Error GoTo error:

Select Case nNum
    Case 0: GetWornType = "Nowhere"
    Case 1: GetWornType = "Everywhere"
    Case 2: GetWornType = "Head"
    Case 3: GetWornType = "Hands"
    Case 4: GetWornType = "Finger"
    Case 5: GetWornType = "Feet"
    Case 6: GetWornType = "Arms"
    Case 7: GetWornType = "Back"
    Case 8: GetWornType = "Neck"
    Case 9: GetWornType = "Legs"
    Case 10: GetWornType = "Waist"
    Case 11: GetWornType = "Torso"
    Case 12: GetWornType = "Off-Hand"
    Case 13: GetWornType = "Finger"
    Case 14: GetWornType = "Wrist"
    Case 15: GetWornType = "Ears"
    Case 16: GetWornType = "Worn"
    Case 17: GetWornType = "Wrist"
    Case 18: GetWornType = "Eyes"
    Case 19: GetWornType = "Face"
    
    Case Else: GetWornType = "Unknown (" & nNum & ")"
End Select

out:
Exit Function
error:
Call HandleError("GetWornType")
Resume out:
End Function
Public Function GetWeaponType(ByVal nNum As Integer) As String
On Error GoTo error:

Select Case nNum
    Case 0: GetWeaponType = "1H Blunt"
    Case 1: GetWeaponType = "2H Blunt"
    Case 2: GetWeaponType = "1H Sharp"
    Case 3: GetWeaponType = "2H Sharp"
    Case Else: GetWeaponType = "Unknown (" & nNum & ")"
End Select

out:
Exit Function
error:
Call HandleError("GetWeaponType")
Resume out:
End Function
Public Function GetClassWeaponType(nNum As Integer) As String
On Error GoTo error:

Select Case nNum
    Case 0: GetClassWeaponType = "1H Blunt"
    Case 1: GetClassWeaponType = "2H Blunt"
    Case 2: GetClassWeaponType = "1H Sharp"
    Case 3: GetClassWeaponType = "2H Sharp"
    Case 4: GetClassWeaponType = "Any 1H"
    Case 5: GetClassWeaponType = "Any 2H"
    Case 6: GetClassWeaponType = "Any Sharp"
    Case 7: GetClassWeaponType = "Any Blunt"
    Case 8: GetClassWeaponType = "Any Weapon"
    Case 9: GetClassWeaponType = "Staff"
    Case Else: GetClassWeaponType = nNum
End Select

out:
Exit Function
error:
Call HandleError("GetClassWeaponType")
Resume out:
End Function

Public Function GetArmourType(nNum As Integer) As String
On Error GoTo error:

Select Case nNum
    Case 0: GetArmourType = "Natural"
    Case 1: GetArmourType = "Silk"      '"Robes"
    Case 2: GetArmourType = "Ninja"     '"Padded"
    Case 3, 4, 5, 6: GetArmourType = "Leather" '"Soft Leather","Soft Studded","Rigid Leather","Rigid Studded"
    Case 7: GetArmourType = "Chainmail"
    Case 8: GetArmourType = "Scalemail"
    Case 9: GetArmourType = "Platemail"
    Case Else: GetArmourType = nNum
End Select

out:
Exit Function
error:
Call HandleError("GetArmourType")
Resume out:
End Function
Public Function GetMagery(nNum As Integer, Optional nLevel As Integer) As String
On Error GoTo error:

Select Case nNum
    Case 0: GetMagery = "None"
    Case 1: GetMagery = "Mage"
    Case 2: GetMagery = "Priest"
    Case 3: GetMagery = "Druid"
    Case 4: GetMagery = "Bard"
    Case 5: GetMagery = "Kai"
    Case Else: GetMagery = nNum
End Select

If Not nNum = 0 Then
    GetMagery = GetMagery & "-" & nLevel
End If

out:
Exit Function
error:
Call HandleError("GetMagery")
Resume out:

End Function
Public Function CreateAccessTables(ByVal DataSource As String, ByVal UseMultiplier As Boolean) As Boolean
On Error GoTo error:
Dim catNewDB As ADOX.Catalog, tabNewRooms As ADOX.Table, tabNewMonsters As ADOX.Table, tabNewSpells As ADOX.Table
Dim tabNewClasses As ADOX.Table, tabNewRaces As ADOX.Table, tabNewItems As ADOX.Table, tabNewShops As ADOX.Table
Dim tabNewMessages As ADOX.Table, tabNewActions As ADOX.Table, tabNewTextblocks As ADOX.Table, tabNewInfo As ADOX.Table
Dim pkClasses As New ADOX.Key
Dim pkRaces As New ADOX.Key
Dim pkItems As New ADOX.Key
Dim pkMonsters As New ADOX.Key
Dim pkShops As New ADOX.Key
Dim pkSpells As New ADOX.Key
Dim pkMessages As New ADOX.Key
Dim pkActions As New ADOX.Key
Dim idxTextblocks As New ADOX.Index
Dim idxRooms As New ADOX.Index
Dim x As Integer

CreateAccessTables = False

Set catNewDB = New ADOX.Catalog
Set tabNewRooms = New ADOX.Table
Set tabNewRaces = New ADOX.Table
Set tabNewClasses = New ADOX.Table
Set tabNewMessages = New ADOX.Table
Set tabNewActions = New ADOX.Table
Set tabNewSpells = New ADOX.Table
Set tabNewShops = New ADOX.Table
Set tabNewMonsters = New ADOX.Table
Set tabNewItems = New ADOX.Table
Set tabNewTextblocks = New ADOX.Table
Set tabNewInfo = New ADOX.Table

'open the database
catNewDB.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & DataSource
'create new table object
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
With tabNewRooms
    .Name = "Rooms"
    .Columns.Append "Map Number", adInteger
    .Columns.Append "Room Number", adInteger
    .Columns.Append "Name", adVarWChar, 53
    For x = 0 To 6
        .Columns.Append CStr("Desc " & x), adVarWChar, 71
    Next
    .Columns.Append "AnsiMap", adVarWChar, 13
    .Columns.Append "Type", adInteger
    .Columns.Append "Shop Number", adInteger
    .Columns.Append "Gang House Number", adInteger
    .Columns.Append "Mon Type", adInteger
    .Columns.Append "Min Index", adInteger
    .Columns.Append "Max Index", adInteger
    .Columns.Append "Max Regen", adInteger
    .Columns.Append "Max Area", adInteger
    .Columns.Append "Control Room", adInteger
    .Columns.Append "Light", adInteger
    .Columns.Append "Runic", adInteger
    .Columns.Append "Platinum", adInteger
    .Columns.Append "Gold", adInteger
    .Columns.Append "Silver", adInteger
    .Columns.Append "Copper", adInteger
    .Columns.Append "InvisRunic", adInteger
    .Columns.Append "InvisPlatinum", adInteger
    .Columns.Append "InvisGold", adInteger
    .Columns.Append "InvisSilver", adInteger
    .Columns.Append "InvisCopper", adInteger
    .Columns.Append "Attributes", adInteger
    .Columns.Append "Death Room", adInteger
    .Columns.Append "Exit Room", adInteger
    .Columns.Append "Command Text", adInteger
    .Columns.Append "Delay", adInteger
    .Columns.Append "Perm NPC", adInteger
    .Columns.Append "Spell", adInteger
    For x = 0 To 9
        .Columns.Append CStr("Exit " & x), adInteger
        .Columns.Append CStr("Type " & x), adInteger
        .Columns.Append CStr("Para1 " & x), adInteger
        .Columns.Append CStr("Para2 " & x), adInteger
        .Columns.Append CStr("Para3 " & x), adInteger
        .Columns.Append CStr("Para4 " & x), adInteger
    Next
    For x = 0 To 16
        .Columns.Append CStr("Room Item " & x), adInteger
        .Columns.Append CStr("Room Item " & x & " USES"), adInteger
        .Columns.Append CStr("Room Item " & x & " QTY"), adInteger
    Next
    For x = 0 To 14
        .Columns.Append CStr("Hidden Item " & x), adInteger
        .Columns.Append CStr("Hidden Item " & x & " USES"), adInteger
        .Columns.Append CStr("Hidden Item " & x & " QTY"), adInteger
    Next
    For x = 0 To 9
        .Columns.Append CStr("Placed Item " & x), adInteger
    Next
    For x = 0 To 14
        .Columns.Append CStr("CurrentRoomMon " & x), adInteger
    Next
End With
'add the table to database
catNewDB.Tables.Append tabNewRooms
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
With tabNewClasses
    .Name = "Classes"
    .Columns.Append "Number", adInteger
    .Columns.Append "Name", adVarWChar, 29
    .Columns.Append "Min HP", adInteger
    .Columns.Append "Max HP", adInteger
    .Columns.Append "EXP %", adInteger
    .Columns.Append "Magic Type", adInteger
    .Columns.Append "Magic LVL", adInteger
    .Columns.Append "Weapon", adInteger
    .Columns.Append "Armour", adInteger
    .Columns.Append "Combat", adInteger
    .Columns.Append "Title Text", adInteger
    
    For x = 0 To 9
        .Columns.Append CStr("Ability " & x), adInteger
        .Columns.Append CStr("Ability Value " & x), adInteger
    Next
    
End With

catNewDB.Tables.Append tabNewClasses
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
With tabNewRaces
    .Name = "Races"
    .Columns.Append "Number", adInteger
    .Columns.Append "Name", adVarWChar, 29
    .Columns.Append "Min INT", adInteger
    .Columns.Append "Min WIL", adInteger
    .Columns.Append "Min STR", adInteger
    .Columns.Append "Min HEA", adInteger
    .Columns.Append "Min AGL", adInteger
    .Columns.Append "Min CHM", adInteger
    .Columns.Append "Max INT", adInteger
    .Columns.Append "Max WIL", adInteger
    .Columns.Append "Max STR", adInteger
    .Columns.Append "Max HEA", adInteger
    .Columns.Append "Max AGL", adInteger
    .Columns.Append "Max CHM", adInteger
    .Columns.Append "HP Bonus", adInteger
    .Columns.Append "CP", adInteger
    .Columns.Append "EXP %", adInteger
    
    For x = 0 To 9
        .Columns.Append CStr("Ability " & x), adInteger
        .Columns.Append CStr("Ability Value " & x), adInteger
    Next
    
End With

catNewDB.Tables.Append tabNewRaces
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
With tabNewSpells
    .Name = "Spells"
    .Columns.Append "Number", adInteger
    .Columns.Append "Name", adVarWChar, 29
    .Columns.Append "Short Name", adVarWChar, 5
    .Columns.Append "Level", adInteger
    .Columns.Append "Desc 1", adVarWChar, 50
    .Columns.Append "Desc 2", adVarWChar, 50
    .Columns.Append "Cast MSG A", adInteger
    .Columns.Append "Cast MSG B", adInteger
    .Columns.Append "MSG Style", adInteger
    .Columns.Append "Energy", adInteger
    .Columns.Append "Mana", adInteger
    .Columns.Append "Min", adInteger
    .Columns.Append "Max", adInteger
    .Columns.Append "Spell Type", adInteger
    .Columns.Append "Type of Resists", adInteger
    .Columns.Append "Difficulty", adInteger
    .Columns.Append "Target", adInteger
    .Columns.Append "Duration", adInteger
    .Columns.Append "Attack Type", adInteger
    .Columns.Append "Resist Ability", adInteger
    .Columns.Append "Magery A", adInteger
    .Columns.Append "Magery B", adInteger
    .Columns.Append "Level Cap", adInteger
    .Columns.Append "LVLS Max Increase", adInteger
    .Columns.Append "Max Increase", adInteger
    .Columns.Append "LVLS Min Increase", adInteger
    .Columns.Append "Min Increase", adInteger
    .Columns.Append "LVLS Dur Increase", adInteger
    .Columns.Append "Dur Increase", adInteger
    .Columns.Append "UNDEFINED01", adInteger
    .Columns.Append "UNDEFINED02", adInteger
    
    For x = 0 To 9
        .Columns.Append CStr("Ability " & x), adInteger
        .Columns.Append CStr("Ability Value " & x), adInteger
    Next
End With

catNewDB.Tables.Append tabNewSpells
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
With tabNewMonsters
    .Name = "Monsters"
    .Columns.Append "Number", adInteger
    .Columns.Append "Name", adVarWChar, 29
    .Columns.Append "Group", adInteger
    .Columns.Append "Index", adInteger
    .Columns.Append "Weapon Number", adInteger
    .Columns.Append "AC", adInteger
    .Columns.Append "DR", adInteger
    .Columns.Append "Follow", adInteger
    .Columns.Append "MR", adInteger
    .Columns.Append "Experience", adDouble
    If eDatFileVersion >= v111j And UseMultiplier = True Then .Columns.Append "Exp Multiplier", adDouble
    .Columns.Append "Hit Points", adInteger
    .Columns.Append "HP Regen", adInteger
    .Columns.Append "Energy", adInteger
    .Columns.Append "Game Limit", adInteger
    .Columns.Append "Charm LvL", adInteger
    .Columns.Append "Charm RES", adInteger
    .Columns.Append "BS Defense", adInteger
    .Columns.Append "Active", adInteger
    .Columns.Append "Type", adInteger
    .Columns.Append "Undead", adInteger
    .Columns.Append "Alignment", adInteger
    .Columns.Append "Gender", adInteger
    .Columns.Append "Regen Time", adInteger
    .Columns.Append "Date Killed", adInteger
    .Columns.Append "Time Killed", adInteger
    .Columns.Append "Runic", adInteger
    .Columns.Append "Platinum", adInteger
    .Columns.Append "Gold", adInteger
    .Columns.Append "Silver", adInteger
    .Columns.Append "Copper", adInteger
    .Columns.Append "Move Msg", adInteger
    .Columns.Append "Death Msg", adInteger
    .Columns.Append "Greet Txt", adInteger
    .Columns.Append "Desc Txt", adInteger
    .Columns.Append "Talk Txt", adInteger
    .Columns.Append "Death Spell", adInteger
    .Columns.Append "Create Spell", adInteger
    .Columns.Append "Desc 1", adVarWChar, 70
    .Columns.Append "Desc 2", adVarWChar, 70
    .Columns.Append "Desc 3", adVarWChar, 70
    .Columns.Append "Desc 4", adVarWChar, 70

    For x = 0 To 4
        .Columns.Append CStr("Attack Type " & x), adInteger
        .Columns.Append CStr("Attack Accu/Spell " & x), adInteger
        .Columns.Append CStr("Attack % " & x), adInteger
        .Columns.Append CStr("Attack Min Hit/Cast % " & x), adInteger
        .Columns.Append CStr("Attack Max Hit/Cast LVL " & x), adInteger
        .Columns.Append CStr("Attack Hit Msg " & x), adInteger
        .Columns.Append CStr("Attack Dodge Msg " & x), adInteger
        .Columns.Append CStr("Attack Miss Msg " & x), adInteger
        .Columns.Append CStr("Attack Energy " & x), adInteger
        .Columns.Append CStr("Attack Hit Spell " & x), adInteger
        .Columns.Append CStr("Spell Number " & x), adInteger
        .Columns.Append CStr("Spell Cast % " & x), adInteger
        .Columns.Append CStr("Spell Cast LVL " & x), adInteger
    Next
    
    For x = 0 To 9
        .Columns.Append CStr("Item Number " & x), adInteger
        .Columns.Append CStr("Item Uses " & x), adInteger
        .Columns.Append CStr("Item Drop % " & x), adInteger
    Next
    
    For x = 0 To 9
        .Columns.Append CStr("Ability " & x), adInteger
        .Columns.Append CStr("Ability Value " & x), adInteger
    Next

End With

catNewDB.Tables.Append tabNewMonsters
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
With tabNewItems
    .Name = "Items"
    .Columns.Append "Number", adInteger
    .Columns.Append "Name", adVarWChar, 29
    .Columns.Append "Game Limit", adInteger
    .Columns.Append "Desc1", adVarWChar, 60
    .Columns.Append "Desc2", adVarWChar, 60
    .Columns.Append "Desc3", adVarWChar, 60
    .Columns.Append "Desc4", adVarWChar, 60
    .Columns.Append "Desc5", adVarWChar, 60
    .Columns.Append "Desc6", adVarWChar, 60
    .Columns.Append "Weight", adInteger
    .Columns.Append "Type", adInteger
    .Columns.Append "Uses", adInteger
    .Columns.Append "Cost", adInteger
    .Columns.Append "Cost Type", adInteger
    .Columns.Append "Min Hit", adInteger
    .Columns.Append "Max Hit", adInteger
    .Columns.Append "AC", adInteger
    .Columns.Append "DR", adInteger
    .Columns.Append "Weapon", adInteger
    .Columns.Append "Armour", adInteger
    .Columns.Append "Worn On", adInteger
    .Columns.Append "Accuracy", adInteger
    .Columns.Append "Gettable", adInteger
    .Columns.Append "Req Str", adInteger
    .Columns.Append "Speed", adInteger
    .Columns.Append "Robable", adInteger
    .Columns.Append "Hit Msg", adInteger
    .Columns.Append "Miss Msg", adInteger
    .Columns.Append "Read Msg", adInteger
    .Columns.Append "Distruct Msg", adInteger
    .Columns.Append "Not Droppable", adInteger
    .Columns.Append "Destroy On Death", adInteger
    .Columns.Append "Retain After Uses", adInteger
    .Columns.Append "OpenRunic", adInteger
    .Columns.Append "OpenPlatinum", adInteger
    .Columns.Append "OpenGold", adInteger
    .Columns.Append "OpenSilver", adInteger
    .Columns.Append "OpenCopper", adInteger
    
    For x = 0 To 9
        .Columns.Append ("Class " & x), adInteger
    Next
    
    For x = 0 To 9
        .Columns.Append ("Race " & x), adInteger
    Next
    
    For x = 0 To 9
        .Columns.Append ("Negate " & x), adInteger
    Next

    For x = 0 To 19
        .Columns.Append ("Ability " & x), adInteger
        .Columns.Append ("Ability Value " & x), adInteger
    Next
    
End With
catNewDB.Tables.Append tabNewItems
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
With tabNewShops
    .Name = "Shops"
    .Columns.Append "Number", adInteger
    .Columns.Append "Name", adVarWChar, 39
    .Columns.Append "Desc A", adVarWChar, 52
    .Columns.Append "Desc B", adVarWChar, 52
    .Columns.Append "Desc C", adVarWChar, 52
    .Columns.Append "Type", adInteger
    .Columns.Append "Min Lvl", adInteger
    .Columns.Append "Max Lvl", adInteger
    .Columns.Append "MarkUp", adInteger
    .Columns.Append "Class Limit", adInteger
    
    For x = 0 To 19
        .Columns.Append CStr("Item " & x), adInteger
        .Columns.Append CStr("Max " & x), adInteger
        .Columns.Append CStr("Normal " & x), adInteger
        .Columns.Append CStr("Regen Time " & x), adInteger
        .Columns.Append CStr("Regen Number" & x), adInteger
        .Columns.Append CStr("Regen %" & x), adInteger
    Next
End With
catNewDB.Tables.Append tabNewShops
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
With tabNewMessages
    .Name = "Messages"
    .Columns.Append "Number", adInteger
    .Columns.Append "Line 1", adVarWChar, 74
    .Columns.Append "Line 2", adVarWChar, 74
    .Columns.Append "Line 3", adVarWChar, 74
End With
catNewDB.Tables.Append tabNewMessages
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
With tabNewTextblocks
    .Name = "Textblocks"
    .Columns.Append "Number", adInteger
    .Columns.Append "Part #", adInteger
    .Columns.Append "Link To", adInteger
'    .Columns.Append "Data Part 1", adVarWChar, 250
'    .Columns.Append "Data Part 2", adVarWChar, 250
'    .Columns.Append "Data Part 3", adVarWChar, 250
'    .Columns.Append "Data Part 4", adVarWChar, 250
'    .Columns.Append "Data Part 5", adVarWChar, 250
'    .Columns.Append "Data Part 6", adVarWChar, 250
'    .Columns.Append "Data Part 7", adVarWChar, 250
'    .Columns.Append "Data Part 8", adVarWChar, 250
    .Columns.Append "Data", adLongVarWChar, 2000
End With
catNewDB.Tables.Append tabNewTextblocks
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
With tabNewActions
    .Name = "Actions"
    .Columns.Append "Action", adVarWChar, 29
    .Columns.Append "Single to User", adVarWChar, 74
    .Columns.Append "Single to Room", adVarWChar, 74
    .Columns.Append "User to User", adVarWChar, 74
    .Columns.Append "User to Other User", adVarWChar, 74
    .Columns.Append "User to Room", adVarWChar, 74
    .Columns.Append "Monster to User", adVarWChar, 74
    .Columns.Append "Monster to Room", adVarWChar, 74
    .Columns.Append "Inventory to User", adVarWChar, 74
    .Columns.Append "Inventory to Room", adVarWChar, 74
    .Columns.Append "Floor Item to User", adVarWChar, 74
    .Columns.Append "Floor Item to Room", adVarWChar, 74
End With
catNewDB.Tables.Append tabNewActions
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
With tabNewInfo
    .Name = "Info"
    .Columns.Append "NMR Version", adVarWChar, 250
    .Columns.Append "Dat File Version", adVarWChar, 250
    .Columns.Append "Date", adVarWChar, 250
    .Columns.Append "Time", adVarWChar, 250
    .Columns.Append "Custom", adVarWChar, 250
End With
catNewDB.Tables.Append tabNewInfo

DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
With pkActions
    .Name = "pkActions"
    .Type = adKeyPrimary
    .RelatedTable = "Actions"
    .Columns.Append "Action"
    .Columns("Action").RelatedColumn = "Action"
    .UpdateRule = adRINone
End With
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
With pkMessages
    .Name = "pkMessages"
    .Type = adKeyPrimary
    .RelatedTable = "Messages"
    .Columns.Append "Number"
    .Columns("Number").RelatedColumn = "Number"
    .UpdateRule = adRINone
End With
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
With pkSpells
    .Name = "pkSpells"
    .Type = adKeyPrimary
    .RelatedTable = "Spells"
    .Columns.Append "Number"
    .Columns("Number").RelatedColumn = "Number"
    .UpdateRule = adRINone
End With
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
With pkShops
    .Name = "pkShops"
    .Type = adKeyPrimary
    .RelatedTable = "Shops"
    .Columns.Append "Number"
    .Columns("Number").RelatedColumn = "Number"
    .UpdateRule = adRINone
End With
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
With pkMonsters
    .Name = "pkMonsters"
    .Type = adKeyPrimary
    .RelatedTable = "Monsters"
    .Columns.Append "Number"
    .Columns("Number").RelatedColumn = "Number"
    .UpdateRule = adRINone
End With
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
With pkItems
    .Name = "pkItems"
    .Type = adKeyPrimary
    .RelatedTable = "Items"
    .Columns.Append "Number"
    .Columns("Number").RelatedColumn = "Number"
    .UpdateRule = adRINone
End With
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
With pkRaces
    .Name = "pkRaces"
    .Type = adKeyPrimary
    .RelatedTable = "Races"
    .Columns.Append "Number"
    .Columns("Number").RelatedColumn = "Number"
    .UpdateRule = adRINone
End With
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
With pkClasses
    .Name = "pkClasses"
    .Type = adKeyPrimary
    .RelatedTable = "Classes"
    .Columns.Append "Number"
    .Columns("Number").RelatedColumn = "Number"
    .UpdateRule = adRINone
End With
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
With idxTextblocks
    .Name = "idxTextblocks"
    .Columns.Append "Number"
    .Columns.Append "Part #"
    .Unique = False
End With
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
With idxRooms
    .Name = "idxRooms"
    .Columns.Append "Map Number"
    .Columns.Append "Room Number"
    .Unique = False
End With
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
tabNewRaces.Keys.Append pkRaces
tabNewClasses.Keys.Append pkClasses
tabNewItems.Keys.Append pkItems
tabNewMonsters.Keys.Append pkMonsters
tabNewMessages.Keys.Append pkMessages
tabNewSpells.Keys.Append pkSpells
tabNewShops.Keys.Append pkShops
tabNewActions.Keys.Append pkActions
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
tabNewRooms.Indexes.Append idxRooms
tabNewTextblocks.Indexes.Append idxTextblocks
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents
Set pkRaces = Nothing
Set pkClasses = Nothing
Set pkShops = Nothing
Set pkMonsters = Nothing
Set pkMessages = Nothing
Set pkActions = Nothing
Set pkSpells = Nothing
Set pkItems = Nothing

DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents

Set idxTextblocks = Nothing
Set idxRooms = Nothing

Set tabNewInfo = Nothing
Set tabNewRooms = Nothing
Set tabNewClasses = Nothing
Set tabNewRaces = Nothing
Set tabNewSpells = Nothing
Set tabNewShops = Nothing
Set tabNewTextblocks = Nothing
Set tabNewItems = Nothing
Set tabNewActions = Nothing
Set tabNewMonsters = Nothing
Set tabNewMessages = Nothing

Set catNewDB = Nothing

CreateAccessTables = True
DoEvents 'If Not chkUseCPU.value = 1 Then DoEvents

Exit Function
error:
Call HandleError
Set pkRaces = Nothing
Set pkClasses = Nothing
Set pkShops = Nothing
Set pkMonsters = Nothing
Set pkMessages = Nothing
Set pkActions = Nothing
Set pkSpells = Nothing
Set pkItems = Nothing

Set idxTextblocks = Nothing
Set idxRooms = Nothing

Set tabNewRooms = Nothing
Set tabNewClasses = Nothing
Set tabNewRaces = Nothing
Set tabNewSpells = Nothing
Set tabNewShops = Nothing
Set tabNewTextblocks = Nothing
Set tabNewItems = Nothing
Set tabNewActions = Nothing
Set tabNewMonsters = Nothing
Set tabNewMessages = Nothing
Set tabNewInfo = Nothing

Set catNewDB = Nothing

End Function

Public Sub LookupAbility(ByVal nAbil As Integer, ByVal nValue As Long)

On Error GoTo error:

Select Case nAbil
    Case 0: 'nothing
    
    Case 143, 158, 185: 'clearitem, req to hit, badattack
        Call frmItem.GotoItem(nValue)
        
    Case 12, 146: '12-summon, 146-mon guard
        Call frmMonster.GotoMonster(nValue)
        
    Case 42, 43, 73, 122, 151, 153, 160: '42-learnsp, 43-castssp, 73-dispell magic, 122-remove spell,151-end cast, 153-killspell, 160-GiveTempSpell
         Call frmSpell.GotoSpell(nValue)
        
    Case 148, 155: '148-textblock, 155-deathtext
        Call frmTextblock.GotoTB(nValue)
    
    Case 101, 115, 120: '101-confuse msg, 115-desc msg, 120-start msg
        Call frmMessage.GotoMSG(nValue)
        
End Select

out:
Exit Sub
error:
Call HandleError("LookupAbility")
Resume out:

End Sub

Public Function DecryptTextblock(ByVal sData As String) As String
Dim sString As String, sDecrypted As String, sChar As String, x As Long

On Error GoTo error:

sString = sData
For x = 1 To Len(sString)
    sChar = Asc(Mid(sString, x, 1))
    If sChar >= 32 Then
        sDecrypted = sDecrypted & Chr(Asc(Mid(sString, x, 1)) - 32)
    End If
Next
sDecrypted = ClipNull(sDecrypted, Len(sDecrypted))

DecryptTextblock = sDecrypted

out:
Exit Function
error:
Call HandleError("DecryptTextblock")
Resume out:

End Function
Public Function EncryptTextblock(ByVal sData As String) As String
Dim decrypted As String, stri As String, stri2 As String
Dim i As Integer

On Error GoTo error:

decrypted = sData
For i = 1 To Len(decrypted)
    stri = Asc(Mid(decrypted, i, 1))
    If stri <= 223 Then stri2 = stri2 & Chr(Asc(Mid(decrypted, i, 1)) + 32)
Next

If Len(stri2) > Len(TextblockRec.Data) Then stri2 = Left(stri2, Len(TextblockRec.Data))

EncryptTextblock = stri2 & String$(Len(TextblockRec.Data) - Len(stri2), Chr(0))

out:
Exit Function
error:
Call HandleError("EncryptTextblock")
Resume out:
        
End Function

Public Function GetMonsterAttackName(ByVal nMonsterNumber As Long, ByVal nAttack As Integer, Optional nLength As Integer = 49) As String
Dim sTemp As String, y As Integer, z As Integer
Dim nStatus As Integer
On Error GoTo error:

If nMonsterNumber = 0 Then
    GetMonsterAttackName = "none"
    Exit Function
Else
    nStatus = BTRCALL(BGETEQUAL, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), nMonsterNumber, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        GetMonsterAttackName = "unknown"
        Exit Function
    End If
End If

MonsterRowToStruct Monsterdatabuf.buf
        
If Monsterrec.AttackType(nAttack) = 0 Then
    GetMonsterAttackName = "None"
ElseIf Monsterrec.AttackType(nAttack) = 1 Then
    sTemp = GetMessages(Monsterrec.AttackHitMsg(nAttack), 1)
    y = InStr(1, sTemp, "%s ", vbTextCompare) + 3
    z = InStr(y, sTemp, " for ", vbTextCompare)
    If z = 0 Then z = InStr(y, sTemp, " ", vbTextCompare)
    
    If y > 3 And z > y Then
        If Mid(sTemp, y, z - y) = "all-out" And InStr(y + Len("all-out") + 1, sTemp, " ", vbTextCompare) > 0 Then
            z = InStr(y + Len("all-out") + 1, sTemp, " ", vbTextCompare)
            sTemp = Mid(sTemp, y, z - y)
        Else
            sTemp = Mid(sTemp, y, z - y)
        End If
    Else
        sTemp = "Physical"
    End If
    GetMonsterAttackName = sTemp
ElseIf Monsterrec.AttackType(nAttack) = 2 Then
    sTemp = GetSpellName(Monsterrec.AttackAccuSpell(nAttack))
    If sTemp = "SPELL NOT IN DATABASE!" Then sTemp = "Spell " & Monsterrec.AttackAccuSpell(nAttack)
    GetMonsterAttackName = sTemp
ElseIf Monsterrec.AttackType(nAttack) = 1 Then
    GetMonsterAttackName = "Rob"
Else
    GetMonsterAttackName = "Unknown"
End If

If Len(GetMonsterAttackName) > nLength Then GetMonsterAttackName = Left(GetMonsterAttackName, nLength - 3) & "..."

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetMonsterAttackName")
Resume out:
End Function
Public Function GetMonGroupName(ByVal nNum As Integer) As String

On Error GoTo error:

Select Case nNum
    Case 0: 'Lair
        GetMonGroupName = "Lair"
    Case 1: 'Wanderer
        GetMonGroupName = "Wanderer"
    Case 2: 'NPC
        GetMonGroupName = "NPC"
    Case 3: 'Living
        GetMonGroupName = "Living"
    Case 4: 'Random
        GetMonGroupName = "Random"
    Case 5: 'Guard
        GetMonGroupName = "Guard"
    Case 6 To 35: 'Group
        GetMonGroupName = "Group " & nNum - 5
    
    Case 36: 'Arena
        GetMonGroupName = "Arena"
    Case 37: 'Angel
        GetMonGroupName = "Angel"
    Case 38: 'Quest
        GetMonGroupName = "Quest"
    Case 39: 'Other
        GetMonGroupName = "Other"
    Case Else:
        GetMonGroupName = "unknown"
        
End Select

out:
Exit Function
error:
Call HandleError("GetMonGroupName")
Resume out:

End Function

Public Function GetRoomExits(ByVal Number As Integer, Optional LongForm As Boolean) As String
On Error GoTo error:

If LongForm = False Then
    Select Case Number
        Case 5: GetRoomExits = "NW"
        Case 0: GetRoomExits = "N"
        Case 4: GetRoomExits = "NE"
        Case 3: GetRoomExits = "W"
        Case 2: GetRoomExits = "E"
        Case 7: GetRoomExits = "SW"
        Case 1: GetRoomExits = "S"
        Case 6: GetRoomExits = "SE"
        Case 8: GetRoomExits = "U"
        Case 9: GetRoomExits = "D"
    End Select
Else
    Select Case Number
        Case 5: GetRoomExits = "northwest"
        Case 0: GetRoomExits = "north"
        Case 4: GetRoomExits = "northeast"
        Case 3: GetRoomExits = "west"
        Case 2: GetRoomExits = "east"
        Case 7: GetRoomExits = "southwest"
        Case 1: GetRoomExits = "south"
        Case 6: GetRoomExits = "southeast"
        Case 8: GetRoomExits = "up"
        Case 9: GetRoomExits = "down"
    End Select
End If

out:
Exit Function
error:
Call HandleError("GetRoomExits")
Resume out:
End Function
Public Function GetRoomName(ByVal MapNum As Long, ByVal RoomNum As Long) As String
Dim nStatus As Integer

On Error GoTo error:

RoomKeyStruct.MapNum = MapNum
RoomKeyStruct.RoomNum = RoomNum

nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    GetRoomName = "ROOM DOESN'T EXIST"
Else
    RoomRowToStruct Roomdatabuf.buf
    GetRoomName = ClipNull(Roomrec.Name)
End If

out:
Exit Function
error:
Call HandleError("GetRoomName")
Resume out:

End Function

Public Function GetTotalActions(ByVal MapNum As Long, ByVal RoomNum As Long, _
    ByVal nExitNum As Integer) As Integer
Dim nStatus As Integer

On Error GoTo error:

RoomKeyStruct.MapNum = MapNum
RoomKeyStruct.RoomNum = RoomNum

nStatus = BTRCALL(BGETEQUAL, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyStruct, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    GetTotalActions = 0
Else
    RoomRowToStruct Roomdatabuf.buf
    If Roomrec.RoomType(nExitNum) = 6 Then 'hidden/action
        GetTotalActions = Abs(Roomrec.Para2(nExitNum))
        If GetTotalActions > 9 Then GetTotalActions = -1
    Else
        GetTotalActions = -1
    End If
End If

out:
Exit Function
error:
Call HandleError("GetTotalActions")
Resume out:

End Function

Public Function GetMonsterName(ByVal nMonsterNumber As Long) As String
Dim nStatus As Integer

On Error GoTo error:

If nMonsterNumber = 0 Then
    GetMonsterName = "none"
Else
    nStatus = BTRCALL(BGETEQUAL, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), nMonsterNumber, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nStatus = 4 Then
            GetMonsterName = "Monster NOT IN DATABASE!"
        Else
            GetMonsterName = "Error: " & BtrieveErrorCode(nStatus)
        End If
    Else
        MonsterRowToStruct Monsterdatabuf.buf
        GetMonsterName = ClipNull(Monsterrec.Name)
    End If
End If

out:
Exit Function
error:
Call HandleError("GetMonsterName")
Resume out:

End Function

Public Function GetItemName(ByVal nItemNumber As Long) As String
Dim nStatus As Integer
On Error GoTo error:

If nItemNumber = 0 Then
    GetItemName = "none"
Else
    nStatus = BTRCALL(BGETEQUAL, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), nItemNumber, Len(nItemNumber), 0)
    If Not nStatus = 0 Then
        If nStatus = 4 Then
            GetItemName = "ITEM NOT IN DATABASE!"
        Else
            MsgBox "Function GetItemName, BGETEQUAL, Error: " & BtrieveErrorCode(nStatus)
        End If
    Else
        ItemRowToStruct Itemdatabuf.buf
        GetItemName = ClipNull(Itemrec.Name)
    End If
End If

out:
Exit Function
error:
Call HandleError("GetItemName")
Resume out:
End Function

Public Function GetItemUses(ByVal nItemNumber As Long) As Integer
Dim nStatus As Integer
On Error GoTo error:

If nItemNumber = 0 Then
    GetItemUses = -1
Else
    nStatus = BTRCALL(BGETEQUAL, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), nItemNumber, Len(nItemNumber), 0)
    If Not nStatus = 0 Then
        If nStatus = 4 Then
            GetItemUses = -1
        Else
            MsgBox "Function GetItemUses, BGETEQUAL, Error: " & BtrieveErrorCode(nStatus)
        End If
    Else
        ItemRowToStruct Itemdatabuf.buf
        GetItemUses = Itemrec.Uses
    End If
End If

out:
Exit Function
error:
Call HandleError("GetItemUses")
Resume out:
End Function

Public Function GetRaceName(ByVal nRaceNumber As Long) As String
Dim nStatus As Integer
On Error GoTo error:

If nRaceNumber = 0 Then
    GetRaceName = "none"
Else
    nStatus = BTRCALL(BGETEQUAL, RacePosBlock, Racedatabuf, Len(Racedatabuf), nRaceNumber, Len(nRaceNumber), 0)
    If Not nStatus = 0 Then
        If nStatus = 4 Then
            GetRaceName = "???"
        Else
            MsgBox "Function GetRaceName, BGETEQUAL, Error: " & BtrieveErrorCode(nStatus)
        End If
    Else
        RaceRowToStruct Racedatabuf.buf
        GetRaceName = ClipNull(Racerec.Name)
    End If
End If

out:
Exit Function
error:
Call HandleError("GetRaceName")
Resume out:
End Function

Public Function GetClassName(ByVal nClassNumber As Long) As String
Dim nStatus As Integer
On Error GoTo error:

If nClassNumber = 0 Then
    GetClassName = "none"
Else
    nStatus = BTRCALL(BGETEQUAL, ClassPosBlock, Classdatabuf, Len(Classdatabuf), nClassNumber, Len(nClassNumber), 0)
    If Not nStatus = 0 Then
        If nStatus = 4 Then
            GetClassName = "???"
        Else
            MsgBox "Function GetClassName, BGETEQUAL, Error: " & BtrieveErrorCode(nStatus)
        End If
    Else
        ClassRowToStruct Classdatabuf.buf
        GetClassName = ClipNull(Classrec.Name)
    End If
End If

out:
Exit Function
error:
Call HandleError("GetClassName")
Resume out:
End Function

Public Function GetShopName(ByVal nShopNumber As Long) As String
Dim nStatus As Integer
On Error GoTo error:

If nShopNumber = 0 Then
    GetShopName = "none"
Else
    nStatus = BTRCALL(BGETEQUAL, ShopPosBlock, Shopdatabuf, Len(Shopdatabuf), nShopNumber, Len(nShopNumber), 0)
    If Not nStatus = 0 Then
        If nStatus = 4 Then
            GetShopName = "SHOP NOT IN DATABASE!"
        Else
            MsgBox "Function GetShopName, BGETEQUAL, Error: " & BtrieveErrorCode(nStatus)
        End If
    Else
        ShopRowToStruct Shopdatabuf.buf
        GetShopName = ClipNull(Shoprec.Name)
    End If
End If

out:
Exit Function
error:
Call HandleError("GetShopName")
Resume out:
End Function

Public Function GetSpellName(ByVal nSpellNumber As Integer) As String
Dim nStatus As Integer
On Error GoTo error:

If nSpellNumber = 0 Then
    GetSpellName = "none"
Else
    nStatus = BTRCALL(BGETEQUAL, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), nSpellNumber, Len(nSpellNumber), 0)
    If Not nStatus = 0 Then
        If nStatus = 4 Then
            GetSpellName = "SPELL NOT IN DATABASE!"
        Else
            MsgBox "Function GetSpellName, BGETEQUAL, Error: " & BtrieveErrorCode(nStatus)
        End If
    Else
        SpellRowToStruct Spelldatabuf.buf
        GetSpellName = ClipNull(Spellrec.Name)
    End If
End If

out:
Exit Function
error:
Call HandleError("GetSpellName")
Resume out:
End Function

Public Function GetSpell(ByVal nSpellNumber As Long) As Integer
Dim nStatus As Integer
On Error GoTo error:

If nSpellNumber = 0 Then
    GetSpell = 1
Else
    nStatus = BTRCALL(BGETEQUAL, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), nSpellNumber, Len(nSpellNumber), 0)
    If Not nStatus = 0 Then
        If nStatus = 4 Then
            GetSpell = 1
        Else
            MsgBox "Function GetSpellName, BGETEQUAL, Error: " & BtrieveErrorCode(nStatus)
            GetSpell = 1
        End If
    Else
        SpellRowToStruct Spelldatabuf.buf
        GetSpell = 0
    End If
End If

out:
Exit Function
error:
Call HandleError("GetSpell")
GetSpell = 1
Resume out:
End Function

Public Function GetSpellRange(ByVal nSpellNumber As Long, ByVal nCastLevel As Long, _
    Optional ByVal nEnergyUsed As Integer) As String
Dim nMin As Currency, nMax As Currency, nDur As Currency, nEnergy As Currency, x As Integer
Dim nMRVal As Currency
Dim nStatus As Integer
On Error GoTo error:

If nSpellNumber = 0 Then Exit Function
nStatus = BTRCALL(BGETEQUAL, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), nSpellNumber, Len(nSpellNumber), 0)
If Not nStatus = 0 Then Exit Function

SpellRowToStruct Spelldatabuf.buf

nEnergy = IIf(nEnergyUsed > 0, nEnergyUsed, Spellrec.Energy)
If nEnergy < 1000 And nEnergy > 0 Then
    nEnergy = Fix(1000 / nEnergy)
    If nEnergy < 2 Then nEnergy = 0
Else
    nEnergy = 0
End If

nMin = GetSpellMinDamage(nSpellNumber, nCastLevel)
nMax = GetSpellMaxDamage(nSpellNumber, nCastLevel)

If Spellrec.LVLSDurIncr = 0 Then
    nDur = Spellrec.duration
Else
    nDur = Spellrec.duration + Fix((Spellrec.DurIncrease / Spellrec.LVLSDurIncr) * nCastLevel)
End If

If nEnergy > 0 Then
    nMin = nMin * nEnergy
    nMax = nMax * nEnergy
End If
GetSpellRange = IIf(nEnergy > 0, "(x" & nEnergy & "): ", "") _
    & nMin & " to " & nMax & IIf(nDur <> 0, ", " & nDur & " rnds", "")

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetSpellRange")
Resume out:

End Function

Public Function GetSpellMinDamage(ByVal nSpellNumber As Long, Optional ByVal nCastLevel As Integer = 0) As Long
Dim nStatus As Integer
On Error GoTo error:

GetSpellMinDamage = 0

If nSpellNumber = 0 Then Exit Function
nStatus = BTRCALL(BGETEQUAL, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), nSpellNumber, Len(nSpellNumber), 0)
If Not nStatus = 0 Then Exit Function

SpellRowToStruct Spelldatabuf.buf

If Spellrec.LVLSMinIncr = 0 Or nCastLevel <= 0 Then
    GetSpellMinDamage = Spellrec.Min
Else
    GetSpellMinDamage = Spellrec.Min + Fix((Spellrec.MinIncrease / Spellrec.LVLSMinIncr) * nCastLevel)
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetSpellMinDamage")
Resume out:
End Function

Public Function GetSpellMaxDamage(ByVal nSpellNumber As Long, Optional ByVal nCastLevel As Integer = 0) As Long
Dim nStatus As Integer
On Error GoTo error:

GetSpellMaxDamage = 0

If nSpellNumber = 0 Then Exit Function
nStatus = BTRCALL(BGETEQUAL, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), nSpellNumber, Len(nSpellNumber), 0)
If Not nStatus = 0 Then Exit Function

SpellRowToStruct Spelldatabuf.buf

If Spellrec.LVLSMaxIncr = 0 Or nCastLevel <= 0 Then
    GetSpellMaxDamage = Spellrec.Max
Else
    GetSpellMaxDamage = Spellrec.Max + Fix((Spellrec.MaxIncrease / Spellrec.LVLSMaxIncr) * nCastLevel)
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetSpellMaxDamage")
Resume out:
End Function

Public Function GetSpellDuration(ByVal nSpellNumber As Long, Optional ByVal nCastLevel As Integer = 0) As Long
Dim nStatus As Integer
On Error GoTo error:

GetSpellDuration = 0

If nSpellNumber = 0 Then Exit Function
nStatus = BTRCALL(BGETEQUAL, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), nSpellNumber, Len(nSpellNumber), 0)
If Not nStatus = 0 Then Exit Function

SpellRowToStruct Spelldatabuf.buf

If Spellrec.LVLSDurIncr = 0 Or nCastLevel <= 0 Then
    GetSpellDuration = Spellrec.duration
Else
    GetSpellDuration = Spellrec.duration + Fix((Spellrec.DurIncrease / Spellrec.LVLSDurIncr) * nCastLevel)
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetSpellDuration")
Resume out:
End Function

Public Function GetShortSpellName(ByVal nSpellNumber As Long) As String
Dim nStatus As Integer
On Error GoTo error:

If nSpellNumber = 0 Then
    GetShortSpellName = "none"
Else
    nStatus = BTRCALL(BGETEQUAL, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), nSpellNumber, Len(nSpellNumber), 0)
    If Not nStatus = 0 Then
        GetShortSpellName = "unkwn"
    Else
        SpellRowToStruct Spelldatabuf.buf
        GetShortSpellName = ClipNull(Spellrec.ShortName)
    End If
End If

out:
Exit Function
error:
Call HandleError("GetShortSpellName")
Resume out:
End Function

Public Function SpellHasAbility(ByVal nSpellNumber As Long, ByVal nAbility As Integer) As Integer
Dim nStatus As Integer, x As Integer
On Error GoTo error:

'-1 = does not have
'>=0 = value of ability

SpellHasAbility = -1
If nAbility <= 0 Or nSpellNumber <= 0 Then Exit Function

nStatus = BTRCALL(BGETEQUAL, SpellPosBlock, Spelldatabuf, Len(Spelldatabuf), nSpellNumber, Len(nSpellNumber), 0)
If Not nStatus = 0 Then Exit Function

SpellRowToStruct Spelldatabuf.buf

For x = 0 To 9
    If Spellrec.AbilityA(x) = nAbility Then
        SpellHasAbility = Spellrec.AbilityB(x)
        Exit Function
    End If
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("SpellHasAbility")
Resume out:
End Function

Public Function GetTextblockCMDS(ByVal TextBlockNumber As Long, Optional ByVal nMaxLength As Integer) As String
Dim nStatus As Integer, x1 As Integer, x2 As Integer, sDecrypted As String

On Error GoTo error:

If TextBlockNumber = 0 Then GetTextblockCMDS = "none": Exit Function

TextblockKey.Number = TextBlockNumber
TextblockKey.PartNum = 0
getlinkto:
    
nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKeyStructToRow(), KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    If nStatus = 4 Then
        GetTextblockCMDS = "TEXTBLOCK (or LinkTo) NOT IN DATABASE"
    Else
        MsgBox "Function GetTextblock, BGETEQUAL, Error: " & BtrieveErrorCode(nStatus)
    End If
Else
    TextblockRowToStruct TextblockDataBuf.buf
    
'    If TextblockRec.LinkTo <> 0 Then
'        TextblockKey.Number = TextblockRec.LinkTo
'        GoTo getlinkto:
'    End If
    
    sDecrypted = DecryptTextblock(TextblockRec.Data)
    
    x1 = 1
    x1 = InStr(x1, sDecrypted, ":")
    If x1 = 0 Then GetTextblockCMDS = "none": Exit Function
    
    GetTextblockCMDS = Mid(sDecrypted, 1, x1 - 1)
    
    x1 = x1 + 1
    Do While x1 < Len(sDecrypted)
        x1 = InStr(x1, sDecrypted, Chr(10)) + 1
        If x1 = 1 Then GoTo done:
        
        x2 = InStr(x1, sDecrypted, ":")
        If x2 = 0 Then GoTo done:
        GetTextblockCMDS = GetTextblockCMDS & ", " & Mid(sDecrypted, x1, x2 - x1)
        
        x1 = x2 + 1
    Loop
    
End If

done:
If nMaxLength > 0 And Len(GetTextblockCMDS) > nMaxLength Then
    GetTextblockCMDS = Left(GetTextblockCMDS, nMaxLength - 1) & "+"
End If

out:
Exit Function
error:
Call HandleError("GetTextblockCMDS")
Resume out:

End Function
Public Function GetTextblock(ByVal TextBlockNumber As Long) As String
Dim nStatus As Integer

On Error GoTo error:

If TextBlockNumber = 0 Then GetTextblock = "none": Exit Function

TextblockKey.Number = TextBlockNumber
TextblockKey.PartNum = 0
getlinkto:
    
nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKeyStructToRow(), KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    If nStatus = 4 Then
        GetTextblock = "TEXTBLOCK (or LinkTo) NOT IN DATABASE"
    Else
        MsgBox "Function GetTextblock, BGETEQUAL, Error: " & BtrieveErrorCode(nStatus)
    End If
Else
    TextblockRowToStruct TextblockDataBuf.buf
    
    If TextblockRec.LinkTo <> 0 Then
        TextblockKey.Number = TextblockRec.LinkTo
        GoTo getlinkto:
    End If
    
    GetTextblock = DecryptTextblock(TextblockRec.Data)
    
    If Len(GetTextblock) > 59 Then
        GetTextblock = Left(GetTextblock, 60) & "..."
    End If
    
End If

out:
Exit Function
error:
Call HandleError("GetTextblock")
Resume out:

End Function

Public Function GetTextblockLink(ByVal TextBlockNumber As Long) As Long
Dim nStatus As Integer

On Error GoTo error:

If TextBlockNumber = 0 Then GetTextblockLink = 0: Exit Function

TextblockKey.Number = TextBlockNumber
TextblockKey.PartNum = 0
getlinkto:
    
nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKeyStructToRow(), KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    If nStatus = 4 Then
        GetTextblockLink = 0
    Else
        MsgBox "Function GetTextblockLink, BGETEQUAL, Error: " & BtrieveErrorCode(nStatus)
    End If
Else
    TextblockRowToStruct TextblockDataBuf.buf
    GetTextblockLink = TextblockRec.LinkTo
End If

Exit Function
error:
Call HandleError("GetTextblockLink")

End Function

Public Function GetMessages(ByVal nMessageNumber As Long, ByVal nWhichMessage As Integer) As String
Dim nStatus As Integer, sTemp As String

'enter < 1 for all messages

On Error GoTo error:

If nMessageNumber = 0 Then
    GetMessages = "none"
Else
    nStatus = BTRCALL(BGETEQUAL, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), nMessageNumber, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nStatus = 4 Then
            GetMessages = "MSG NOT IN DATABASE!"
        Else
            MsgBox "Function GetMessageName, BGETEQUAL, Error: " & BtrieveErrorCode(nStatus)
        End If
    Else
        MessageRowToStruct Messagedatabuf.buf
        
        If nWhichMessage < 1 Then
            sTemp = ClipNull(Messagerec.MessageLine1)
            If Not sTemp = "" Then GetMessages = sTemp
            
            sTemp = ClipNull(Messagerec.MessageLine2)
            If Not sTemp = "" Then
                If Not GetMessages = "" Then GetMessages = GetMessages & ", "
                GetMessages = GetMessages & sTemp
            End If
            
            sTemp = ClipNull(Messagerec.MessageLine3)
            If Not sTemp = "" Then
                If Not GetMessages = "" Then GetMessages = GetMessages & ", "
                GetMessages = GetMessages & sTemp
            End If
        Else
            Select Case nWhichMessage
                Case Is = 1
                    GetMessages = ClipNull(Messagerec.MessageLine1)
                Case Is = 2
                    GetMessages = ClipNull(Messagerec.MessageLine2)
                Case Is = 3
                    GetMessages = ClipNull(Messagerec.MessageLine3)
            End Select
        End If
    End If
End If

out:
Exit Function
error:
Call HandleError("GetMessages")
Resume out:
End Function

Public Function GetAlignmentValue(ByVal nEvilPoints As Long) As String

'Saint = -201
'Good = -200 to -51
'Neut = -50 to 29
'Seedy = 30 to 39
'Outlaw = 40 to 79
'Criminal = 80 to 119
'Villain = 120 to 209
'Fiend = 210

On Error GoTo error:

Select Case nEvilPoints
    Case Is <= -201: 'Saint = -201
        GetAlignmentValue = "Saint"
    Case -200 To -51: 'Good = -200 to -51
        GetAlignmentValue = "Good"
    Case -50 To 29: 'Neut = -50 to 29
        GetAlignmentValue = "Neutral"
    Case 30 To 39: 'Seedy = 30 to 39
        GetAlignmentValue = "Seedy"
    Case 40 To 79: 'Outlaw = 40 to 79
        GetAlignmentValue = "Outlaw"
    Case 80 To 119: 'Criminal = 80 to 119
        GetAlignmentValue = "Criminal"
    Case 120 To 209: 'Villain = 120 to 209
        GetAlignmentValue = "Villain"
    Case Is >= 210: 'Fiend = 210
        GetAlignmentValue = "Fiend"
    Case Else:
        GetAlignmentValue = nEvilPoints
End Select

out:
Exit Function
error:
Call HandleError("GetAlignmentValue")
Resume out:

End Function
Public Sub CreateMGIL()
Dim nStatus As Integer, x As Integer

On Error GoTo error:

nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Could not get first monster record, Error: " & BtrieveErrorCode(nStatus), vbOKOnly, "Creating Monster Group/Index List"
    Exit Sub
End If

Erase MGIL()
ReDim MGIL(39, 9999)

Do While nStatus = 0
    MonsterRowToStruct Monsterdatabuf.buf
    
    frmProgressBar.lblPanel(1).Caption = Monsterrec.Number
    Call frmProgressBar.IncreaseProgress
    
    If Left(Monsterrec.Name, 3) = "sdf" Then GoTo Skip:
    If Monsterrec.Index < 0 Then GoTo Skip:
    
    If UBound(MGIL(), 2) < Monsterrec.Index Then ReDim Preserve MGIL(UBound(MGIL(), 1), Monsterrec.Index)
    'If UBound(MGIL(), 3) < Monsterrec.Number Then ReDim Preserve MGIL(UBound(MGIL(), 1), UBound(MGIL(), 2), Monsterrec.Number)
    
    For x = 0 To 10 '20
        If MGIL(Monsterrec.Group, Monsterrec.Index).nNumber(x) = 0 Then
            MGIL(Monsterrec.Group, Monsterrec.Index).nNumber(x) = Monsterrec.Number
            'MGIL(Monsterrec.Group, Monsterrec.Index).sName(x) = ClipNull(Monsterrec.Name)
            GoTo Skip:
        End If
    Next
    
Skip:
    nStatus = BTRCALL(BGETNEXT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

out:
Exit Sub
error:
Call HandleError("CreateMGIL")
Resume out:

End Sub
Public Function TestPasteChar(ByVal sTestChar As String) As Boolean
On Error GoTo error:

TestPasteChar = True

Select Case LCase(sTestChar)
    Case "a":
    Case "e":
    Case "i":
    Case "o":
    Case "u":
    Case "y":
    
    Case "b":
    Case "c":
    Case "d":
    Case "f":
    Case "g":
    Case "h":
    Case "j":
    Case "k":
    Case "l":
    Case "m":
    Case "n":
    Case "p":
    Case "q":
    Case "r":
    Case "s":
    Case "t":
    Case "v":
    Case "w":
    Case "x":
    Case "z":
    
    Case "(":
    Case ")":
    
    Case "-":
    Case "_":
    Case ",":
    Case ":":
    Case " ":
    Case "'":
    Case """":
    Case ".":
    Case "`":
    
    Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0":
    
    Case Else: TestPasteChar = False
End Select

Exit Function
error:
Call HandleError
End Function
Public Function TestAlphaChar(ByVal sTestChar As String) As Boolean
On Error GoTo error:

TestAlphaChar = True

Select Case LCase(sTestChar)
    Case "a":
    Case "e":
    Case "i":
    Case "o":
    Case "u":
    Case "y":
    
    Case "b":
    Case "c":
    Case "d":
    Case "f":
    Case "g":
    Case "h":
    Case "j":
    Case "k":
    Case "l":
    Case "m":
    Case "n":
    Case "p":
    Case "q":
    Case "r":
    Case "s":
    Case "t":
    Case "v":
    Case "w":
    Case "x":
    Case "z":
    
    Case Else: TestAlphaChar = False
End Select

Exit Function
error:
Call HandleError
End Function

Public Function CalcMarkup(ByVal nCost As Double, ByVal nMarkUp As Double) As Double
On Error GoTo error:

If nCost <= 0 Then CalcMarkup = 0: Exit Function

If nMarkUp > 0 Then
    CalcMarkup = nCost + Round(nCost * (nMarkUp / 100))
Else
    CalcMarkup = nCost
End If

out:
Exit Function
error:
Call HandleError("CalcMarkup")
Resume out:
    
End Function

Public Function GetItem(ByVal nItemNumber As Long) As Integer
Dim nStatus As Integer
On Error GoTo error:

If nItemNumber = 0 Then
    GetItem = 1
Else
    nStatus = BTRCALL(BGETEQUAL, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), nItemNumber, Len(nItemNumber), 0)
    If Not nStatus = 0 Then
        If nStatus = 4 Then
            GetItem = 1
        Else
            MsgBox "Function GetItemName, BGETEQUAL, Error: " & BtrieveErrorCode(nStatus)
            GetItem = 1
        End If
    Else
        ItemRowToStruct Itemdatabuf.buf
        GetItem = 0
    End If
End If

out:
Exit Function
error:
Call HandleError("GetItem")
GetItem = 1
Resume out:
End Function

Public Function GetItemLimit(ByVal nItemNumber As Long) As Integer
Dim nStatus As Integer
On Error GoTo error:

GetItemLimit = -1

If nItemNumber = 0 Then Exit Function
    
nStatus = BTRCALL(BGETEQUAL, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), nItemNumber, Len(nItemNumber), 0)
If Not nStatus = 0 Then
    If nStatus <> 4 Then
        MsgBox "Function GetItemName, BGETEQUAL, Error: " & BtrieveErrorCode(nStatus)
    End If
    Exit Function
End If

ItemRowToStruct Itemdatabuf.buf
GetItemLimit = Itemrec.GameLimit

out:
Exit Function
error:
Call HandleError("GetItem")
GetItemLimit = -1
Resume out:
End Function

Public Function ItemHasAbility(ByVal nItemNumber As Long, ByVal nAbility As Integer) As Integer
Dim nStatus As Integer, x As Integer
On Error GoTo error:

'-1 = does not have
'>=0 = value of ability

ItemHasAbility = -1
If nAbility <= 0 Or nItemNumber <= 0 Then Exit Function

nStatus = BTRCALL(BGETEQUAL, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), nItemNumber, Len(nItemNumber), 0)
If Not nStatus = 0 Then Exit Function

ItemRowToStruct Itemdatabuf.buf

For x = 0 To 19
    If Itemrec.AbilityA(x) = nAbility Then
        ItemHasAbility = Itemrec.AbilityB(x)
        Exit Function
    End If
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("ItemHasAbility")
Resume out:
End Function

Public Function GetItemCost(ByVal nNum As Long, Optional ByVal nMarkUp As Integer) As String
Dim nStatus As Integer

On Error GoTo error:

If nNum = 0 Then GetItemCost = "0": Exit Function

nStatus = BTRCALL(BGETEQUAL, ItemPosBlock, Itemdatabuf, Len(Itemdatabuf), nNum, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then GetItemCost = "?": Exit Function
Call ItemRowToStruct(Itemdatabuf.buf)

If Itemrec.Cost = 0 Then
    GetItemCost = "Free"
Else
    If nMarkUp > 0 Then
        GetItemCost = (Itemrec.Cost + Round(Itemrec.Cost * (nMarkUp / 100)))
    Else
        GetItemCost = Itemrec.Cost
    End If
    
    GetItemCost = PutCommas(Val(GetItemCost)) & " " & GetCostType(Itemrec.CostType)
End If

out:
Exit Function
error:
Call HandleError("GetItemCost")
Resume out:

End Function

Public Function GetCostType(ByVal nNum As Integer) As String
On Error GoTo error:

Select Case nNum
    Case 0: GetCostType = "Copper"
    Case 1: GetCostType = "Silver"
    Case 2: GetCostType = "Gold"
    Case 3: GetCostType = "Platinum"
    Case 4: GetCostType = "Runic"
    Case Else: GetCostType = "Unknown (" & nNum & ")"
End Select

out:
Exit Function
error:
Call HandleError("GetCostType")
Resume out:
End Function


Public Function CalcTrueAverage(ByVal nSwings As Double, ByVal nHitP As Double, ByVal nHitA As Long, _
    ByVal nCritP As Double, ByVal nCritA As Long, ByVal nExtraP As Double, ByVal nExtraA As Long) As Double

On Error GoTo error:

If nSwings <= 0 Then CalcTrueAverage = -1: Exit Function
If nSwings > 5 Then nSwings = 5

nHitP = nHitP / 100
nCritP = nCritP / 100
nExtraP = nExtraP / 100
'((HIT% * HITAVE) + (CRIT% * CRITAVE) + (HIT% * EXTRA% * EXTRAAVE) + (CRIT% * EXTRA% * EXTRAAVE)) * SWINGS
'CalcTrueAverage = Round(((nHitP * nHitA) + (nCritP * nCritA) + (nHitP * nExtraP * nExtraA) + (nCritP * nExtraP * nExtraA)) * nSwings, 2)
CalcTrueAverage = Round(((nHitP * nHitA) + (nCritP * nCritA) + ((nHitP + nCritP) * nExtraP * nExtraA)) * nSwings, 2)

Exit Function
error:
Call HandleError("CalcTrueAverage")

End Function


