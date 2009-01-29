VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APIAssociations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------
' File: APIAssociations.cls
' Copyright: InferMed Ltd 2008 All Rights Reserved
' Author: Nicky Johns Mar 2008
' Purpose: User role import/export routines for MACRO Interface
'       NB This is just a wrapper for the .NET C# "MACROSSURBS30" DLL
'------------------------------------------------------------------
' REVISIONS
' NCJ 6-17 Mar 08 : Initial development
'
'------------------------------------------------------------------

Option Explicit

'------------------------------------------------------------------
Friend Function ExportAssociations(ByVal sSerialisedUser As String, _
                                ByVal sXMLAssocRequest As String, _
                                ByRef sXMLAssociations As String) As Boolean
'------------------------------------------------------------------
' Export user role associations contained in sXMLAssocRequest specification
' We use the SSURIO class in the C# "MACROSSURBS30" DLL
'------------------------------------------------------------------
Dim oUser As MACROUser
Dim oSSURIO As SSURIO

    On Error GoTo ErrLabel
    
    ' Create the MACRO User
    Set oUser = New MACROUser
    Call oUser.SetStateHex(sSerialisedUser)
    
    ' Call C# dll to do assoc export
    Set oSSURIO = New SSURIO
    sXMLAssociations = oSSURIO.ExportAssocsXML(sXMLAssocRequest, oUser.CurrentDBConString)
    Set oSSURIO = Nothing
    
    Set oUser = Nothing

    ' Succeed if we got something back
    ExportAssociations = (sXMLAssociations <> "")
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|APIAssociations.ExportAssociations"
    
End Function

'------------------------------------------------------------------
Friend Function ImportAssociations(ByVal sSerialisedUser As String, _
                              ByVal sXMLAssociations As String, _
                              ByRef sMessage As String) As eImportResult
'------------------------------------------------------------------
' Export user role associations contained in sXMLAssocRequest specification
' We use the SSURIO class in the C# "MACROSSURBS30" DLL
'------------------------------------------------------------------
Dim oUser As MACROUser
Dim oSSURIO As SSURIO
Dim sSecCon As String

    On Error GoTo ErrLabel
    
    sMessage = ""
    
    ' Create the MACRO User
    Set oUser = New MACROUser
    Call oUser.SetStateHex(sSerialisedUser)

    ' We have to check the current user's permissions
    If oUser.CheckPermission(gsFnChangeAccessRights) Then
            ' Call C# dll to do assoc export
            Set oSSURIO = New SSURIO
            ' Get decrypted Security DB connection string
            sSecCon = DecryptString(oUser.SecurityDBConEncrypted)
            ImportAssociations = oSSURIO.ImportAssocsXML(sXMLAssociations, oUser.CurrentDBConString, _
                                oUser.DatabaseCode, sSecCon, oUser.UserName, sMessage)
            Set oSSURIO = Nothing
    Else
        ' No permission error
        ImportAssociations = eImportResult.PermissionDenied
        sMessage = "User does not have Change Access Rights permission"
    End If
    
    Set oUser = Nothing

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|APIAssociations.ImportAssociations"
    
End Function

