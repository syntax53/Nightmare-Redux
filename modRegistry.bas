Attribute VB_Name = "modRegistry"
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal Hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
    "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, _
    ByVal cbData As Long) As Long

Private Const KEY_READ = &H20019
Private Const KEY_WRITE = &H20006  '((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Private Const REG_SZ = 1
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003

' Yes, there's a double-space in "Version  6.15"...
Public Const BTRV_CREATE_KEY_PATH = "SOFTWARE\\Btrieve Technologies\\Microkernel Workstation Engine\\Version  6.15\\Settings"
Public Const BTRV_CREATE_KEY_VALUE_NAME = "Create 5x Files"

' Write or Create a Registry value
' returns False if successful
'
' Use KeyName = "" for the default value
'
' Value can be an integer value (REG_DWORD), a string (REG_SZ)
' or an array of binary (REG_BINARY). Raises an error otherwise.

Private Function SetRegistryValue(ByVal Hkey As Long, ByVal KeyName As String, _
    ByVal ValueName As String, value As Variant) As Boolean
    Dim Handle As Long
    Dim lngValue As Long
    Dim str As String
    Dim binValue() As Byte
    Dim length As Long
    Dim retVal As Long
    retVal = 1 'defailt as failed
    
    ' Open the key, exit if not found
    If RegOpenKeyEx(Hkey, KeyName, 0, KEY_WRITE, Handle) = 0 Then
        ' three cases, according to the data type in Value
        Select Case VarType(value)
            Case vbInteger, vbLong
                lngValue = value
                retVal = RegSetValueEx(Handle, ValueName, 0, REG_DWORD, lngValue, 4)
            Case vbString
                str = value
                retVal = RegSetValueEx(Handle, ValueName, 0, REG_SZ, ByVal str, Len(str))
            Case vbArray + vbByte
                binValue = value
                length = UBound(binValue) - LBound(binValue) + 1
                retVal = RegSetValueEx(Handle, ValueName, 0, REG_BINARY, _
                    binValue(LBound(binValue)), length)
            Case Else
                RegCloseKey Handle
                Err.Raise 1001, , "SetRegistryValue: Unsupported Value Type"
        End Select
        ' Close the key and signal success
        RegCloseKey Handle
    End If

    SetRegistryValue = (retVal <> 0)
End Function

Public Function GetBTRCreateIsV5() As Boolean
    Dim Handle As Long
    Dim str As String
    Dim LenValue As Long
    Dim dummy As Integer
    str = "0"
    
    If (RegOpenKeyEx(HKEY_LOCAL_MACHINE, BTRV_CREATE_KEY_PATH, 0, KEY_READ, Handle) = 0) Then
        dummy = RegQueryValueEx(Handle, BTRV_CREATE_KEY_VALUE_NAME, 0, REG_SZ, vbNullString, LenValue)
        If LenValue > 0 Then
            str = Space(LenValue)
            Call RegQueryValueEx(Handle, BTRV_CREATE_KEY_VALUE_NAME, 0, REG_SZ, ByVal str, Len(str))
            Call RegCloseKey(Handle)
            ' trim and remove any nulls...
            str = Trim$(str)
            If Mid$(str, 1, 1) = Chr$(0) Then str = Mid$(str, 2)
            If Mid$(str, Len(str), 1) = Chr$(0) Then str = Left$(str, Len(str) - 1)
        End If
    End If
    
    GetBTRCreateIsV5 = CBool(CInt(str) * -1)
End Function

Public Function SetBTRCreateV5(is_v5 As Boolean) As Boolean
    SetBTRCreateV5 = SetRegistryValue(HKEY_LOCAL_MACHINE, BTRV_CREATE_KEY_PATH, BTRV_CREATE_KEY_VALUE_NAME, CStr(Abs(CInt(is_v5))))
End Function

