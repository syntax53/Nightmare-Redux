Attribute VB_Name = "modErrors"
Option Explicit

Public Sub HandleError(ByVal strMessage As String, ByVal strSubject As String, ParamArray InputParms())
On Error GoTo ErrorHandler
Dim lngMax As Long
Dim strParm As String
Dim lngStep As Long
Dim lngLen As Long
    lngMax = UBound(InputParms)
    On Error Resume Next
    For lngStep = 0 To lngMax
        Select Case TypeName(strParm) 'see what type of object the parameter is
            Case "String", "Integer", "Currency", "Date", "Long", "Double", "Boolean", "Byte", "Decimal"
                lngLen = Len(InputParms(lngStep))
                If Err.Number = 450 Then 'most likely an object if cannot do a LEN on it
                    strParm = strParm & "[OBJECT],"
                    Err.Clear 'clear error
                Else
                    If lngLen = 0 Then
                        strParm = strParm & "NULL,"
                    Else
                        strParm = strParm & InputParms(lngStep) & ","
                    End If
                End If
            Case Else ' all others
                strParm = strParm & "[OBJECT],"
        End Select
    Next
    On Error GoTo ErrorHandler
    strParm = Left(strParm, Len(strParm) - 1) 'remove trailing comma
    MsgBox ("An error occured" & vbCrLf & strSubject & vbCrLf & strMessage & vbCrLf & "Input Parameters: " & strParm)
    
    Exit Sub
    
ErrorHandler:
End Sub


Public Function InDesign() As Boolean
    InDesign = False: Exit Function
    
    Static lintCallCount As Integer
    Static lblnReturn As Boolean
    
    lintCallCount = lintCallCount + 1
    
    Select Case lintCallCount
        Case 1:
            Debug.Assert InDesign()
        Case 2:
            lblnReturn = True
    End Select
    
    InDesign = lblnReturn
    lintCallCount = 0
    
End Function
