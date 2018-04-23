Attribute VB_Name = "modSettings"
Option Base 0
Option Explicit

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public sINIFileName As String
Public bINIReadOnly As Boolean
Private ret As String

Public Function ReadINI(ByVal Section As String, ByVal Key As String, _
    Optional ByVal sAlternateFile As String, Optional ByVal sDefaultVal As String = "0") As Variant
Dim retlen As Integer
On Error GoTo error:

If sAlternateFile = "" Then sAlternateFile = sINIFileName

reread:
ret = Space$(255)
retlen = GetPrivateProfileString(Section, Key, "", ret, Len(ret), ByVal sAlternateFile)
If retlen = 0 Then
    
    Call WriteINI(Section, Key, sDefaultVal, ByVal sAlternateFile)
    
    If UCase(Section) = UCase("Windows") Then
        If UCase(Right(Key, 3)) = UCase("Top") Or UCase(Right(Key, 4)) = UCase("Left") Then _
            Call WriteINI(Section, Key, 1, ByVal sAlternateFile)
        If UCase(Right(Key, 5)) = UCase("Width") Or UCase(Right(Key, 6)) = UCase("Height") Then _
            Call WriteINI(Section, Key, 4000, ByVal sAlternateFile)
        
    ElseIf UCase(Section) = UCase("Settings") Then
        If UCase(Key) = UCase("eDatFileVersion" & IIf(WorksWithN = True, "_n", IIf(WorksWithWG = True, "_wg", ""))) Then _
            Call WriteINI(Section, Key, 1, ByVal sAlternateFile)
        If UCase(Key) = UCase("ShowAbilityEditWarning") Then _
            Call WriteINI(Section, Key, 1, ByVal sAlternateFile)
        If UCase(Key) = UCase("DatCallLetters" & IIf(WorksWithN = True, "_n", IIf(WorksWithWG = True, "_wg", ""))) Then _
            Call WriteINI(Section, Key, "cc", ByVal sAlternateFile)
        If UCase(Key) = UCase("WGPath" & IIf(WorksWithN = True, "_n", IIf(WorksWithWG = True, "_wg", ""))) Then _
            Call WriteINI(Section, Key, "c:\wgserv\", ByVal sAlternateFile)
        
    ElseIf UCase(Section) = UCase("Options") Then
        If UCase(Key) = UCase("MapSmallMap") Then _
            Call WriteINI(Section, Key, 1, ByVal sAlternateFile)
        If UCase(Key) = UCase("MMECustomName") Then _
            Call WriteINI(Section, Key, "My BBS Name", ByVal sAlternateFile)
        If UCase(Key) = UCase("ImportPath") Then _
            Call WriteINI(Section, Key, App.Path, ByVal sAlternateFile)
        If UCase(Key) = UCase("ExportPath") Then _
            Call WriteINI(Section, Key, App.Path, ByVal sAlternateFile)
        If Right(UCase(Key), 3) = UCase("All") Then _
            Call WriteINI(Section, Key, 1, ByVal sAlternateFile)

    End If
    
    If Not bINIReadOnly Then GoTo reread:
End If

ret = Left$(ret, retlen)
ReadINI = ret
Exit Function
error:
Call HandleError
End Function
Public Sub CheckINIReadOnly()
Dim fso As FileSystemObject, nYesNo As Integer, oFile As File
On Error GoTo error:

bINIReadOnly = False

Set fso = CreateObject("Scripting.FileSystemObject")
Set oFile = fso.GetFile(sINIFileName)

If oFile.Attributes And ReadOnly Then
    bINIReadOnly = True
    nYesNo = MsgBox("settings file is marked 'read only,' attempt to fix?" & vbCrLf & "(settings cannot be saved otherwise)", vbYesNo, "settings file is read-only...")
    If Not nYesNo = vbNo Then
        oFile.Attributes = oFile.Attributes - 1
        bINIReadOnly = False
    End If
End If

Set oFile = Nothing
Set fso = Nothing
Exit Sub

error:
Call HandleError
Set oFile = Nothing
Set fso = Nothing
End Sub
Public Sub WriteINI(ByVal Section As String, ByVal Key As String, ByVal Text As String, _
    Optional ByVal sAlternateFile As String)
On Error GoTo error:

If bINIReadOnly Then Exit Sub
If sAlternateFile = "" Then sAlternateFile = sINIFileName

Call WritePrivateProfileString(Section, Key, Text, ByVal sAlternateFile)

Exit Sub
error:
Call HandleError
End Sub

Public Sub CreateSettings()
Dim fso As FileSystemObject
On Error GoTo error:

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(sINIFileName) = True Then fso.DeleteFile sINIFileName, True
fso.CreateTextFile sINIFileName, True

Call WriteINI("Settings", "AutoCompile", 0)
Call WriteINI("Settings", "eDatFileVersion_n", "6")
Call WriteINI("Settings", "eDatFileVersion", "7")
Call WriteINI("Settings", "DatCallLetters", "cc")
Call WriteINI("Settings", "DatCallLetters_n", "cc")
Call WriteINI("Settings", "WGPath", "c:\wgserv\")
Call WriteINI("Settings", "WGPath_n", "c:\wgserv\")
Call WriteINI("Settings", "ShowAbilityEditWarning", 1)
Call WriteINI("Options", "ImportPath", App.Path)
Call WriteINI("Options", "ExportPath", App.Path)
Call WriteINI("Options", "MapNoColors", 0)
Call WriteINI("Options", "MapSmallMap", 1)
Call WriteINI("Options", "MapFollowMapChanges", 0)
Call WriteINI("Options", "MapEditorDontAskAnything", 0)
Call WriteINI("Options", "MapEditorAutoSame", 0)
Call WriteINI("Options", "MapEditorAutoNext", 0)
Call WriteINI("Options", "MapEditorAutoCreate", 0)
Call WriteINI("Options", "MapEditorAutoNew", 0)
Call WriteINI("Options", "MapEditorAutoExisting", 0)
Call WriteINI("Options", "MapEditorNextNewRoom", 9000)
Call WriteINI("Options", "ExportRoomsAll", 1)
Call WriteINI("Options", "ExportRoomsFrom", 1)
Call WriteINI("Options", "ExportRoomsTo", 9999)
Call WriteINI("Options", "ExportRoomsMap", 1)
Call WriteINI("Options", "ExportItemsAll", 1)
Call WriteINI("Options", "ExportItemsFrom", 1)
Call WriteINI("Options", "ExportItemsTo", 9999)
Call WriteINI("Options", "ExportSpellsAll", 1)
Call WriteINI("Options", "ExportSpellsFrom", 1)
Call WriteINI("Options", "ExportSpellsTo", 9999)
Call WriteINI("Options", "ExportMonstersAll", 1)
Call WriteINI("Options", "ExportMonstersFrom", 1)
Call WriteINI("Options", "ExportMonstersTo", 9999)
Call WriteINI("Options", "ExportShopsAll", 1)
Call WriteINI("Options", "ExportShopsFrom", 1)
Call WriteINI("Options", "ExportShopsTo", 9999)
Call WriteINI("Options", "ExportTextblocksAll", 1)
Call WriteINI("Options", "ExportTextblocksFrom", 0)
Call WriteINI("Options", "ExportTextblocksTo", 9999)
Call WriteINI("Options", "ExportRacesAll", 1)
Call WriteINI("Options", "ExportRacesFrom", 1)
Call WriteINI("Options", "ExportRacesTo", 9999)
Call WriteINI("Options", "ExportClassesAll", 1)
Call WriteINI("Options", "ExportClassesFrom", 1)
Call WriteINI("Options", "ExportClassesTo", 9999)
Call WriteINI("Options", "ExportMessagesAll", 1)
Call WriteINI("Options", "ExportMessagesFrom", 1)
Call WriteINI("Options", "ExportMessagesTo", 9999)
Call WriteINI("Windows", "MapLegendTop", 0)
Call WriteINI("Windows", "MapLegendLeft", 0)
Call WriteINI("Windows", "BankTop", 0)
Call WriteINI("Windows", "BankLeft", 0)
Call WriteINI("Windows", "RaceTop", 0)
Call WriteINI("Windows", "RaceLeft", 0)
Call WriteINI("Windows", "ClassTop", 0)
Call WriteINI("Windows", "ClassLeft", 0)
Call WriteINI("Windows", "SpellTop", 0)
Call WriteINI("Windows", "SpellLeft", 0)
Call WriteINI("Windows", "MonsterTop", 0)
Call WriteINI("Windows", "MonsterLeft", 0)
Call WriteINI("Windows", "ItemTop", 0)
Call WriteINI("Windows", "ItemLeft", 0)
Call WriteINI("Windows", "ShopTop", 0)
Call WriteINI("Windows", "ShopLeft", 0)
Call WriteINI("Windows", "RoomTop", 0)
Call WriteINI("Windows", "RoomLeft", 0)
Call WriteINI("Windows", "MapEditorTop", 0)
Call WriteINI("Windows", "MapEditorLeft", 0)
Call WriteINI("Windows", "MapEditorSmall", 0)
Call WriteINI("Windows", "ItemTop", 0)
Call WriteINI("Windows", "ItemLeft", 0)
Call WriteINI("Windows", "MessageTop", 0)
Call WriteINI("Windows", "MessageLeft", 0)
Call WriteINI("Windows", "TextTop", 675)
Call WriteINI("Windows", "TextLeft", 255)
Call WriteINI("Windows", "UserTop", 0)
Call WriteINI("Windows", "UserLeft", 0)
Call WriteINI("Windows", "ItemTop", 0)
Call WriteINI("Windows", "ItemLeft", 0)
Call WriteINI("Windows", "GangTop", 0)
Call WriteINI("Windows", "GangLeft", 0)
Call WriteINI("Windows", "ActionTop", 0)
Call WriteINI("Windows", "ActionLeft", 0)
Call WriteINI("Windows", "AbEdTop", 0)
Call WriteINI("Windows", "AbEdLeft", 0)
Call WriteINI("Windows", "QuestOrgTop", 0)
Call WriteINI("Windows", "QuestOrgLeft", 0)
Call WriteINI("Windows", "PreviewTop", 0)
Call WriteINI("Windows", "PreviewLeft", 0)
Call WriteINI("Windows", "PreviewWidth", 9720)
Call WriteINI("Windows", "PreviewHeight", 5565)
Call WriteINI("Windows", "ToolsTop", 0)
Call WriteINI("Windows", "ToolsLeft", 0)
Call WriteINI("Windows", "HelpGeneralTop", 0)
Call WriteINI("Windows", "HelpGeneralLeft", 0)
Call WriteINI("Windows", "HelpGeneralWidth", 7560)
Call WriteINI("Windows", "HelpGeneralHeight", 8310)
Call WriteINI("Windows", "HelpMessagesTop", 0)
Call WriteINI("Windows", "HelpMessagesLeft", 0)
Call WriteINI("Windows", "HelpMessagesWidth", 7560)
Call WriteINI("Windows", "HelpMessagesHeight", 8310)
Call WriteINI("Windows", "HelpMonstersTop", 0)
Call WriteINI("Windows", "HelpMonstersLeft", 0)
Call WriteINI("Windows", "HelpMonstersWidth", 7560)
Call WriteINI("Windows", "HelpMonstersHeight", 8310)
Call WriteINI("Windows", "HelpRoomsTop", 0)
Call WriteINI("Windows", "HelpRoomsLeft", 0)
Call WriteINI("Windows", "HelpRoomsWidth", 7560)
Call WriteINI("Windows", "HelpRoomsHeight", 8310)
Call WriteINI("Windows", "HelpTextblocksTop", 0)
Call WriteINI("Windows", "HelpTextblocksLeft", 0)
Call WriteINI("Windows", "HelpTextblocksWidth", 7935)
Call WriteINI("Windows", "HelpTextblocksHeight", 8220)

Set fso = Nothing

Exit Sub
error:
Call HandleError
Set fso = Nothing
End Sub

