VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegistryRoutines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*****************************************************************************
'While most of this code came from multiple source one being www.vb2themax.com
'I could not find any single source that put EVERYTHING together into one easy
'to use package. So here it is. Read the comments for help. The code should be
'pretty much error free but should you find a bug kindly email me and let me know.
' Kiser_Donald@hotmail.com
'Revision: 2.0.0
'Released: 10/30/01
'Author: Don Kiser
'Revised 9/05/02
'        * Added Error Checking to GetRegistryValue
'        * Added Remote Registry Read and Write
'        * fixed Error_more_data
'*******************************************************************************

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hkey As Long, phkResult As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal numBytes As Long)

Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted                                  ' KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Const SYNCHRONIZE = &H100000
Const READ_CONTROL = &H20000
Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Const STANDARD_RIGHTS_ALL = &H1F0000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Const KEY_EXECUTE = (KEY_READ)
Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Const ERROR_MORE_DATA = 234
Const ERROR_NO_MORE_ITEMS = &H103
Const ERROR_KEY_NOT_FOUND = &H2

Enum DataType
 REG_SZ = &H1
 REG_EXPAND_SZ = &H2
 REG_BINARY = &H3
 REG_DWORD = &H4
 REG_MULTI_SZ = &H7
End Enum

Enum hkey
 HKEY_CLASSES_ROOT = &H80000000
 HKEY_CURRENT_USER = &H80000001
 HKEY_LOCAL_MACHINE = &H80000002
 HKEY_USERS = &H80000003
 HKEY_PERFORMANCE_DATA = &H80000004
 HKEY_CURRENT_CONFIG = &H80000005
 HKEY_DYN_DATA = &H80000006
End Enum

Dim mvarhKeySet As Long
Dim mvarKeyRoot As String
Dim mvarSubKey As String
Dim Security As SECURITY_ATTRIBUTES

Public Property Get hkey() As hkey
    hkey = mvarhKeySet
End Property

Public Property Let hkey(ByVal vData As hkey)
    mvarhKeySet = vData
End Property

Public Property Get KeyRoot() As String
    KeyRoot = mvarKeyRoot
End Property

Public Property Let KeyRoot(ByVal vData As String)
    mvarKeyRoot = vData
End Property

Public Property Get Subkey() As String
    Subkey = mvarSubKey
End Property

Public Property Let Subkey(ByVal vData As String)
    mvarSubKey = vData
End Property
'****************************************************************************
'       Check to see if Registry key exists
'       Inputs: None
'       Class Properties: Classname.hkey, Classname.keyroot, Classname.subkey
'       Return: True if key exists
'****************************************************************************
Public Function KeyExists() As Boolean
    Dim handle As Long
    Dim ret As Long
        If RegOpenKeyEx(mvarhKeySet, mvarKeyRoot & "\" & mvarSubKey, 0, KEY_READ, handle) Then
              KeyExists = False
              Exit Function
        End If
        KeyExists = True
End Function
'****************************************************************************
'       Create a key in the registry
'       Inputs: KeyName
'       Class Properties: if Input Empty Classname.subkey
'       Return: 0 if successful
'****************************************************************************
Public Function CreateKey(Optional KeyName As Variant) As String
    Dim handle As Long
    Dim disp As Long
    Dim RetVal As Long
        KeyName = IIf(IsMissing(KeyName), mvarSubKey, CStr(KeyName))
        RetVal = RegCreateKeyEx(mvarhKeySet, mvarKeyRoot & "\" & KeyName, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, Security, handle, disp)
        If RetVal Then Exit Function
        RegCloseKey (handle)
        CreateKey = RetVal
End Function
'****************************************************************************
'       Delete a key from the registry
'       Inputs: SubKey
'       Class Properties: Classname.hkey, Classname.keyroot
'       Returns: 0 if successful
'****************************************************************************
Public Function DeleteKey(KeyName As String) As Long
    Dim RetVal As Long
    Dim handle As Long
        RetVal = RegDeleteKey(mvarhKeySet, mvarKeyRoot & "\" & KeyName)
        If RetVal Then Exit Function
        RegCloseKey (handle)
        DeleteKey = RetVal
End Function
'****************************************************************************
'       Delete the value of a key
'       Inputs: Value Name
'       Class Properties: Classname.hkey, Classname.keyroot, Classname.subkey
'       Return: 0 if successful
'****************************************************************************
Public Function DeleteValue(ValueName As String) As Long
    Dim RetVal As Long
    Dim handle As Long
        RetVal = RegOpenKeyEx(mvarhKeySet, mvarKeyRoot & "\" & mvarSubKey, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, handle)
        If RetVal <> 0 Then 'Operation Failed
            DeleteValue = RetVal
            Exit Function
        End If
        DeleteValue = RegDeleteValue(handle, ValueName)
        RegCloseKey (handle)
End Function
'****************************************************************************
'       Enumerate Value Names under a given key
'       Inputs: Key Root, Key Name
'       Return: a collection of strings
'       Source: Slightly modified from www.vb2themax.com EnumRegistryKeys
'****************************************************************************
Public Function EnumRegistryKeys(ByVal hkey As hkey, ByVal KeyName As String) As _
                Collection
    Dim handle As Long
    Dim length As Long
    Dim index As Long
    Dim subkeyName As String
    Dim fFiletime As FILETIME
         ' initialize the result collection
         Set EnumRegistryKeys = New Collection
         
         ' Open the key, exit if not found
         If Len(KeyName) Then
             If RegOpenKeyEx(hkey, KeyName, 0, KEY_READ, handle) Then Exit Function
             ' in all case the subsequent functions use hKey
             hkey = handle
         End If
         
         Do
             ' this is the max length for a key name
             length = 260
             subkeyName = Space$(length)
             ' get the N-th key, exit the loop if not found
             If RegEnumKeyEx(hkey, index, subkeyName, length, 0, "", vbNull, fFiletime) = ERROR_NO_MORE_ITEMS Then Exit Do
             ' add to the result collection
             subkeyName = Left$(subkeyName, InStr(subkeyName, vbNullChar) - 1)
             EnumRegistryKeys.Add subkeyName, subkeyName
             ' prepare to query for next key
             index = index + 1
         Loop
        
         ' Close the key, if it was actually opened
         If handle Then RegCloseKey handle
        
End Function
'****************************************************************************
'       Enumerate values under a given registry key
'       Inputs: Key Root, Key Name
'       Return: a collection, where each element of the collection
'               is a 2-element array of Variants:
'               element(0) is the value name, element(1) is the value's value
'       Source: Slightly Modified from www.vb2themax.com EnumRegistryValues
'****************************************************************************
Function EnumRegistryValues(ByVal hkey As hkey, ByVal KeyName As String) As _
    Collection
    Dim handle As Long
    Dim index As Long
    Dim valueType As Long
    Dim name As String
    Dim nameLen As Long
    Dim resLong As Long
    Dim resString As String
    Dim length As Long
    Dim valueInfo(0 To 1) As Variant
    Dim RetVal As Long
    Dim i As Integer
    Dim vTemp As Variant
    
    ' initialize the result
    Set EnumRegistryValues = New Collection
    
    ' Open the key, exit if not found.
    If Len(KeyName) Then
        If RegOpenKeyEx(hkey, KeyName, 0, KEY_READ, handle) Then Exit Function
        ' in all cases, subsequent functions use hKey
        hkey = handle
    End If
    
    Do
        ' this is the max length for a key name
        nameLen = 260
        name = Space$(nameLen)
        ' prepare the receiving buffer for the value
        length = 4096
        ReDim resBinary(0 To length - 1) As Byte
        
        ' read the value's name and data
        ' exit the loop if not found
        RetVal = RegEnumValue(hkey, index, name, nameLen, ByVal 0&, valueType, _
            resBinary(0), length)
        
        ' enlarge the buffer if you need more space
        If RetVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To length - 1) As Byte
            RetVal = RegEnumValue(hkey, index, name, nameLen, ByVal 0&, _
                valueType, resBinary(0), length)
        End If
        ' exit the loop if any other error (typically, no more values)
        If RetVal Then Exit Do
        
        ' retrieve the value's name
        valueInfo(0) = Left$(name, nameLen)
        
        ' return a value corresponding to the value type
        Select Case valueType
            
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                valueInfo(1) = resLong
            
            Case REG_SZ
                ' copy everything but the trailing null char
                If length <> 0 Then
                    resString = Space$(length - 1)
                    CopyMemory ByVal resString, resBinary(0), length - 1
                    valueInfo(1) = resString
                Else
                    valueInfo(1) = ""
                End If
                
            Case REG_EXPAND_SZ
                ' copy everything but the trailing null char
                ' expand the environment variable to it's value
                ' Ignore a Blank String
                If length <> 0 Then
                    resString = Space$(length - 1)
                    CopyMemory ByVal resString, resBinary(0), length - 1
                    length = ExpandEnvironmentStrings(resString, resString, Len(resString))
                    valueInfo(1) = TrimNull(resString)
                Else
                    valueInfo(1) = ""
                End If

            Case REG_BINARY
                ' shrink the buffer if necessary
                If length < UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To length - 1) As Byte
                End If
                 'Convert to display as string like this: 00 01 01 00 01
                    For i = 0 To UBound(resBinary)
                         resString = resString & " " & Format(Trim(Hex(resBinary(i))), "0#")
                    Next i
                    valueInfo(1) = LTrim(resString) 'Get rid of leading space
            
            Case REG_MULTI_SZ
                ' copy everything but the 2 trailing null chars
                resString = Space$(length - 2)
                CopyMemory ByVal resString, resBinary(0), length - 2
                
                'convert from null-delimited (vbNullChar) stream of strings
                'to comma delimited stream of strings
                'The listview likes it better that way
                resString = Replace(resString, vbNullChar, ",", , , vbBinaryCompare)
                valueInfo(1) = resString
            
            Case Else
                ' Unsupported value type - do nothing
        End Select
        
        ' add the array to the result collection
        ' the element's key is the value's name
        EnumRegistryValues.Add valueInfo, valueInfo(0)
        
        index = index + 1
    Loop
   
    ' Close the key, if it was actually opened
    If handle Then RegCloseKey handle
        
End Function
'****************************************************************************
'       Read a Registry value
'
'       Inputs: Use KeyName = "" for the Default value
'                If the value isn't there, it returns the DefaultValue
'                argument passed in, or Empty if the argument has been omitted
'       Return: Variant
'
'               REG_DWORD: Long
'               REG_SZ: String
'               REG_EXPAND_SZ: String with Expanded Environment variable
'               REG_BINARY: Byte Array
'               REG_MULTI_SZ: null-delimited (vbNullChar) stream of strings
'                   (VB6 users can use Split to convert to an array of string)
'                    Split(expression[, delimiter[, count[, compare]]])
'       Source: Slightly modified from www.vb2themax GetRegistryValue
'****************************************************************************
Public Function GetRegistryValue(ByVal ValueName As String, Optional DefaultValue As Variant) As Variant
    Dim handle As Long
    Dim resLong As Long
    Dim resString As String
    Dim TestString As String
    Dim resBinary() As Byte
    Dim length As Long
    Dim RetVal As Long
    Dim valueType As Long
    
        ' Prepare the default result
        GetRegistryValue = IIf(IsMissing(DefaultValue), Empty, DefaultValue)
        
        ' Open the key, exit if not found.
        If RegOpenKeyEx(mvarhKeySet, mvarKeyRoot & "\" & mvarSubKey, REG_OPTION_NON_VOLATILE, KEY_READ, handle) Then
           'Don 't overwrite the default value!
           'GetRegistryValue = CVar("Error!")
           Exit Function
        End If
        
        ' prepare a 1K receiving resBinary
        length = 1024
        ReDim resBinary(0 To length - 1) As Byte
        
        ' read the registry key
        RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
            length)
        ' if resBinary was too small, try again
        If RetVal = ERROR_MORE_DATA Then
            ' enlarge the resBinary, and read the value again
            ReDim resBinary(0 To length - 1) As Byte
            RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
                length)
        End If
        
        'Added 11/5/01 Don Kiser
        If RetVal = ERROR_KEY_NOT_FOUND Then
                 RegCloseKey (handle)
                 Exit Function
        End If
        
        ' return a value corresponding to the value type
        Select Case valueType
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                GetRegistryValue = resLong
            
            Case REG_SZ
                ' copy everything but the trailing null char
                ' Ignore Blank Strings
                If length <> 0 Then
                    resString = Space$(length - 1)
                    CopyMemory ByVal resString, resBinary(0), length - 1
                    GetRegistryValue = resString
                End If
            
            Case REG_EXPAND_SZ
                ' copy everything but the trailing null char
                ' expand the environment variable to it's value
                ' Ignore a Blank String
                If length <> 0 Then
                    resString = Space$(length - 1)
                    CopyMemory ByVal resString, resBinary(0), length - 1
                    'This is here only to illustrate the example
                      frmIceCreamParlor.txtExpandActual.Text = resString
                    'End illustration
                    
                    length = ExpandEnvironmentStrings(resString, resString, Len(resString))
                    GetRegistryValue = Left$(resString, length)
                    
                End If
            
            Case REG_BINARY
                ' resize the result resBinary
                If length <> UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To length - 1) As Byte
                End If
                GetRegistryValue = resBinary()
            
            Case REG_MULTI_SZ
                ' copy everything but the 2 trailing null chars
                resString = Space$(length - 2)
                CopyMemory ByVal resString, resBinary(0), length - 2
                'A nonexistant value for REG_MULTI_SZ will return a string of nulls
                'with a length = 1022
                'This is because at the beginging of the routine we define Length = 1024
                ' resString = Space$(length -2) = 1022
                'So If we trims all nulls and are left with an empty string then
                'the value doesn't exist so the defualt value is returned
                'Set resstring to a temporary variable because trimnull will truncate it
                TestString = resString
                If Len(TrimNull(TestString)) > 0 Then GetRegistryValue = resString
                
            Case Else
                ' Unsupported value type - do nothing
                ' Shouldn't ever get here
        End Select
        
        ' close the registry key
     RegCloseKey (handle)
   
End Function

'****************************************************************************
'       Write or Create a Registry value
'
'       Inputs: ValueName, Value, Data Type
'       Class Properties: Classname.hkey, Classname.Keyroot, Classname.subkey
'       Return: True if successful
'
'       Use KeyName = "" for the default value
'       Supports:
'       REG_DWORD      -Integer or Long
'       REG_SZ         -String
'       REG_EXPAND_SZ  -String with Environment Variable Ex. %SystemDrive%
'       REG_BINARY     -an array of binary
'       REG_MULTI_SZ   -Null delimited String with double null terminator
'       Source: Slightly modified from www.vb2themax.com SetRegistryValue
'****************************************************************************
Public Function SetRegistryValue(ByVal ValueName As String, Value As Variant, DType As DataType) As Boolean
    Dim handle As Long
    Dim lngValue As Long
    Dim strValue As String
    Dim binValue() As Byte
    Dim length As Long
    Dim RetVal As Long
    
    ' Open the key, exit if not found
    If RegOpenKeyEx(hkey, mvarKeyRoot & "\" & mvarSubKey, REG_OPTION_NON_VOLATILE, KEY_WRITE, handle) Then
       SetRegistryValue = False 'CVar("Error!")
       Exit Function
    End If

    ' three cases, according to the data type passed
    Select Case DType
        Case REG_DWORD
            lngValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_DWORD, lngValue, 4)
        Case REG_SZ
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_SZ, ByVal strValue, _
                Len(strValue))
        Case REG_BINARY
            binValue = Value
            length = UBound(binValue) - LBound(binValue) + 1
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_BINARY, _
                                   binValue(LBound(binValue)), length)
        Case REG_EXPAND_SZ
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_EXPAND_SZ, ByVal strValue, _
                Len(strValue))
        
        Case REG_MULTI_SZ
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_MULTI_SZ, ByVal strValue, _
                Len(strValue))
        
        Case Else
            ' Unsupported value type - do nothing
            ' Shouldn't ever get here
    End Select
    
    ' Close the key and signal success
     RegCloseKey (handle)
    ' signal success if the value was written correctly
    SetRegistryValue = (RetVal = 0)
    
End Function
'****************************************************************************
'       Read a Value from a Remote Registry
'
'       Inputs: Remote Computer Name, Keyroot,Subkey,ValueName
'       Class Properties: Classname.hkey, Classname.Keyroot, Classname.subkey
'       Return: True if successful
'
'       Use KeyName = "" for the default value
'       Supports:
'       REG_DWORD      -Integer or Long
'       REG_SZ         -String
'       REG_EXPAND_SZ  -String with Environment Variable Ex. %SystemDrive%
'       REG_BINARY     -an array of binary
'       REG_MULTI_SZ   -Null delimited String with double null terminator
'       Source: Slightly modified from www.vb2themax.com GetRegistryValue
'               Addition of API Call for Remote Registry Connection
'****************************************************************************
Public Function ReadRemoteRegistryValue(ByVal sRemoteComputer As String, ByVal hkey As hkey, ByVal ValueName As String, Optional KeyPath As String) As Variant
    
    Dim handle As Long
    Dim lReturnCode, lHive, lhRemoteRegistry As Long
    Dim valueType As Long
    Dim resLong As Long
    Dim resString As String
    Dim TestString As String
    Dim resBinary() As Byte
    Dim length As Long
    Dim RetVal As Long
    Dim RegPath As String
    
    
        RegPath = IIf(IsMissing(KeyPath), mvarKeyRoot & "\" & mvarSubKey, KeyPath)
                
        If RegConnectRegistry(sRemoteComputer, hkey, lhRemoteRegistry) Then
            ReadRemoteRegistryValue = CVar("Error!")
            Exit Function
        End If
        lReturnCode = RegOpenKeyEx(lhRemoteRegistry, RegPath, 0, KEY_ALL_ACCESS, handle)
        
        ' prepare a 1K receiving resBinary
        length = 1024
        ReDim resBinary(0 To length - 1) As Byte
        
        ' read the registry key
        RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
            length)
        ' if resBinary was too small, try again
        
        If RetVal = ERROR_MORE_DATA Then
            ' enlarge the resBinary, and read the value again
            ReDim resBinary(0 To length - 1) As Byte
            RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
                length)
        End If
        
        'Added 11/5/01 for error handling Don Kiser
        If RetVal = ERROR_KEY_NOT_FOUND Then
            RegCloseKey (handle)
            Exit Function
        End If
        
        ' return a value corresponding to the value type
        Select Case valueType
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                ReadRemoteRegistryValue = resLong
            
            Case REG_SZ
                ' copy everything but the trailing null char
                ' Ignore Blank Strings
                If length <> 0 Then
                    resString = Space$(length - 1)
                    CopyMemory ByVal resString, resBinary(0), length - 1
                    ReadRemoteRegistryValue = resString
                End If
            
            Case REG_EXPAND_SZ
                ' copy everything but the trailing null char
                ' expand the environment variable to it's value
                ' Ignore a Blank String
                If length <> 0 Then
                    resString = Space$(length - 1)
                    CopyMemory ByVal resString, resBinary(0), length - 1
                    'This is here only to illustrate the example
                      frmIceCreamParlor.txtExpandActual.Text = resString
                    'End illustration
                    
                    length = ExpandEnvironmentStrings(resString, resString, Len(resString))
                    ReadRemoteRegistryValue = Left$(resString, length)
                    
                End If
            
            Case REG_BINARY
                ' resize the result resBinary
                If length <> UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To length - 1) As Byte
                End If
                ReadRemoteRegistryValue = resBinary()
            
            Case REG_MULTI_SZ
                ' copy everything but the 2 trailing null chars
                resString = Space$(length - 2)
                CopyMemory ByVal resString, resBinary(0), length - 2
                'A nonexistant value for REG_MULTI_SZ will return a string of nulls
                'with a length = 1022
                'This is because at the beginging of the routine we define Length = 1024
                ' resString = Space$(length -2) = 1022
                'So If we trims all nulls and are left with an empty string then
                'the value doesn't exist so the defualt value is returned
                'Set resstring to a temporary variable because trimnull will truncate it
                TestString = resString
                If Len(TrimNull(TestString)) > 0 Then ReadRemoteRegistryValue = resString
                
            Case Else
                ' Unsupported value type - do nothing
                ' Shouldn't ever get here
        End Select
        
        ' close the registry key
        RegCloseKey (handle)

End Function

'****************************************************************************
'       Write to a Remote Registry value
'
'       Inputs: Remote Computer Name,hKey, ValueName, Value, Data Type
'       Class Properties:  Classname.Keyroot, Classname.subkey
'       Return: True if successful
'
'       Use KeyName = "" for the default value
'       Supports:
'       REG_DWORD      -Integer or Long
'       REG_SZ         -String
'       REG_EXPAND_SZ  -String with Environment Variable Ex. %SystemDrive%
'       REG_BINARY     -an array of binary
'       REG_MULTI_SZ   -Null delimited String with double null terminator
'       Source: Slightly modified from www.vb2themax.com SetRegistryValue
'               Added ability to connect to Remote Machine
'****************************************************************************
Public Function WriteRemoteRegistryValue(ByVal sRemoteComputer As String, ByVal hkey As hkey, ByVal ValueName As String, Value As Variant, DType As DataType, Optional KeyPath As String) As Boolean
    Dim handle As Long
    Dim lngValue As Long
    Dim strValue As String
    Dim binValue() As Byte
    Dim length As Long
    Dim RetVal As Long
    Dim RegPath As String
    Dim lhRemoteRegistry As Long
    Dim lReturnCode     As Long
    
            
        RegPath = IIf(IsMissing(KeyPath), mvarKeyRoot & "\" & mvarSubKey, KeyPath)
                
        If RegConnectRegistry(sRemoteComputer, hkey, lhRemoteRegistry) Then
            WriteRemoteRegistryValue = CVar("Error!")
            Exit Function
        End If
        lReturnCode = RegOpenKeyEx(lhRemoteRegistry, RegPath, 0, KEY_ALL_ACCESS, handle)
        
    ' three cases, according to the data type passed
    Select Case DType
        Case REG_DWORD
            lngValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_DWORD, lngValue, 4)
        Case REG_SZ
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_SZ, ByVal strValue, _
                Len(strValue))
        Case REG_BINARY
            binValue = Value
            length = UBound(binValue) - LBound(binValue) + 1
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_BINARY, _
                                   binValue(LBound(binValue)), length)
        Case REG_EXPAND_SZ
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_EXPAND_SZ, ByVal strValue, _
                Len(strValue))
        
        Case REG_MULTI_SZ
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_MULTI_SZ, ByVal strValue, _
                Len(strValue))
        
        Case Else
            ' Unsupported value type - do nothing
            ' Shouldn't ever get here
    End Select
    
    ' Close the key and signal success
     RegCloseKey (handle)
    ' signal success if the value was written correctly
    WriteRemoteRegistryValue = (RetVal = 0)
    
End Function
'****************************************************************************
' Trim to first Null character
' Inputs: String with null characaters
' Return: String up to where first null character occured
'****************************************************************************
Public Function TrimNull(item As String) As String
    Dim pos As Integer
        pos = InStr(item, Chr$(0))
        If pos Then item = Left$(item, pos - 1)
        TrimNull = item
End Function
