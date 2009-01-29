VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APICategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------
' File: APICategories.cls
' Copyright: InferMed Ltd 2007 All Rights Reserved
' Author: Nicky Johns Nov 2007
' Purpose: Category import/export routines for MACRO Interface
'       NB This is just a wrapper for the .NET C# "MACROCATBS30" DLL
'------------------------------------------------------------------
' REVISIONS
' NCJ 23 Nov - 6 Dec 07 : Initial development
'
'------------------------------------------------------------------

Option Explicit

'------------------------------------------------------------------
Friend Function ImportCategories(ByVal sSerialisedUser As String, _
                                ByVal sXMLCatsInput As String, _
                                ByRef sXMLReport As String) As Long
'------------------------------------------------------------------
' Import categories contained in sXMLCatsInput specification
' We use the TopCats class in the C# "MACROCATBS30" DLL
' Errors will be caught in the MACROAPI class
'------------------------------------------------------------------
Dim oUser As MACROUser
Dim oTopCats As TopCats

    On Error GoTo ErrLabel
    
    ' Create the MACRO User
    Set oUser = New MACROUser
    Call oUser.SetStateHex(sSerialisedUser)
    
    ' Call C# dll to do category import
    Set oTopCats = New TopCats
    ImportCategories = oTopCats.ImportCats(sXMLCatsInput, oUser.CurrentDBConString, oUser.UserName, sXMLReport)
    Set oTopCats = Nothing
    Set oUser = Nothing
        
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|APICategories.ImportCategories"
    
End Function

'------------------------------------------------------------------
Friend Function ExportCategories(ByVal sSerialisedUser As String, _
                                ByVal sXMLCatRequest As String, _
                                ByRef sXMLReport As String) As Boolean
'------------------------------------------------------------------
' Export categories contained in sXMLCatRequest request string
' We use the TopCats class in the C# "MACROCATBS30" DLL
' Errors will be caught in the MACROAPI class
'------------------------------------------------------------------
Dim oUser As MACROUser
Dim oTopCats As TopCats

    On Error GoTo ErrLabel
    
    ' Create the MACRO User
    Set oUser = New MACROUser
    Call oUser.SetStateHex(sSerialisedUser)
    
    ' Call C# dll to do category export
    Set oTopCats = New TopCats
    ExportCategories = oTopCats.ExportCats(sXMLCatRequest, oUser.CurrentDBConString, oUser.UserName, sXMLReport)
    Set oTopCats = Nothing
    
    Set oUser = Nothing

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|APICategories.ExportCategories"
    
End Function



