VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Permissions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2002-2003. All Rights Reserved
'   File:       Permissions.cls
'   Author:     Richard Meinesz, September 2002
'   Purpose:
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   NCJ 27 Nov 03 - Added class Terminate
'   NCJ 3 Dec 03 - Check for empty permissions collection
'   DPH 24/03/2004 - Clean up collections in class_terminate
'----------------------------------------------------------------------------------------'

Option Explicit

Private mcolUserFunctions As Collection
Private msUserRole As String

'----------------------------------------------------------------------------------------'
Friend Property Get UserFunctions() As Collection
'----------------------------------------------------------------------------------------'

    Set UserFunctions = mcolUserFunctions

End Property

'----------------------------------------------------------------------------------------'
Friend Property Get UserRole() As String
'----------------------------------------------------------------------------------------'

    UserRole = msUserRole

End Property

'----------------------------------------------------------------------------------------'
Friend Function CheckPermission(ByVal sPermission As String) As Boolean
'----------------------------------------------------------------------------------------'
'REM 23/09/02
'Checks if a user has a specific permission
'----------------------------------------------------------------------------------------'
    
    On Error GoTo Errorlabel

    CheckPermission = CollectionMember(mcolUserFunctions, sPermission, False)
    
Exit Function
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Permissions.CheckPermission"
End Function

'----------------------------------------------------------------------------------------'
Friend Sub LoadUserRoleFunctions(oSecCon As Connection, sUserRole As String, ByRef sSecCon As String)
'----------------------------------------------------------------------------------------'
'REM 20/09/02
'Load up a collecton of permissions (functions) accociated with a specific user role, and sets the user role property
'MLM 03/07/03: To address Oracle case sensitivity, additionally pass in connection string here and use for database-specific SQL.
'----------------------------------------------------------------------------------------'
Dim rsUserFunctions As ADODB.Recordset
Dim sSQL As String
Dim sFnCode As String

    On Error GoTo Errorlabel

    Select Case Connection_Property(CONNECTION_PROVIDER, sSecCon)
    Case CONNECTION_MSDAORA, CONNECTION_ORAOLEDB_ORACLE
        sSQL = "SELECT FunctionCode FROM RoleFunction WHERE upper(RoleCode) = upper('" & sUserRole & "')"
    Case Else
        sSQL = "SELECT FunctionCode FROM RoleFunction " _
            & " WHERE RoleCode = '" & sUserRole & "'"
    End Select

    Set rsUserFunctions = New ADODB.Recordset
    rsUserFunctions.Open sSQL, oSecCon, adOpenKeyset, adLockPessimistic, adCmdText
    
    Set mcolUserFunctions = New Collection

    'set up the user functions collection property
    With rsUserFunctions
        Do Until .EOF = True
            sFnCode = rsUserFunctions!FunctionCode

            mcolUserFunctions.Add sFnCode, sFnCode
               .MoveNext
        Loop
    End With
    
    'set the user role property
    msUserRole = sUserRole

    rsUserFunctions.Close
    Set rsUserFunctions = Nothing

Exit Sub
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Permissions.LoadUserRoleFunctions"
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'
Dim vFunctions As Variant
Dim i As Long
    
    msUserRole = PropBag.ReadProperty("R")
    vFunctions = Split(PropBag.ReadProperty("V"), ",")
    Set mcolUserFunctions = New Collection
    For i = 0 To UBound(vFunctions)
        mcolUserFunctions.Add vFunctions(i), vFunctions(i)
    Next
    
End Sub

'----------------------------------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------------------------------
' NCJ 27 Nov 03 - Tidy up on termination
' DPH 24/03/2004 - remove all from collection
'----------------------------------------------------------------------

    Call CollectionRemoveAll(mcolUserFunctions)
    Set mcolUserFunctions = Nothing
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------------------------------------------'
' Store comma delimited list in PropBag
' NCJ 3 Dec 03 - Must check for empty collection
'----------------------------------------------------------------------------------------'
Dim s As String
Dim i As Long

    s = ""
    ' NCJ 3 Dec 03 - Check for empty collection
    If mcolUserFunctions.Count > 0 Then
        s = mcolUserFunctions(1)
        For i = 2 To mcolUserFunctions.Count
            s = s & "," & mcolUserFunctions(i)
        Next
    End If
    PropBag.WriteProperty "V", s
    PropBag.WriteProperty "R", msUserRole
        
End Sub
