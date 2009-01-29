VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998-2000. All Rights Reserved
'   File:       clsSite.cls
'   Author:     Paul Norris, 26/07/99
'   Purpose:    Class to wrap up rules about a site, reading and writing site details.
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'   Revisions:
'   1   Paul Norris     07/09/99    Upgrade database access code from DAO to ADO
'   2   PN              15/09/99    Changed call to ADODBConnection() to MacroADODBConnection()
'   NCJ 18 Jan 00   Deal with single quotes in site description
'   NCJ 29 Apr 00 - SR 3357
'   NCJ 6/7/00 SR3660 Use ReplaceQuotes when saving edited Site Description
'   DPH 06/08/2002 - Added new field SiteLocation
'   ZA  27/09/2002 - Added Locale, Country and Time zone properties
'  ASH Added optional parameter in sub Validate for GetMacroDBSetting
'--------------------------------------------------------------------------------
Option Explicit

Public Event IsValid(bIsValid As Boolean)

Public Enum SiteSaveResult
    ESiteSaved
    ESiteAlreadyExists
End Enum

Public Enum SiteLocation
    ESiteNoLocation = -1
    ESiteServer = 0
    ESiteRemote = 1
End Enum

Private msDescription As String
Private msCode As String
Private mbNewRecord As Boolean
Private mbHasChanges As Boolean
Private mbIsValid As Boolean
Private mnLocation As Integer
Private mlSiteLocale As Long
Private mlCountryCode As Long
Private mlTimeZone As Long

'--------------------------------------------------------------------------------
Public Function Load() As Boolean
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
' REVISIONS
' DPH 06/08/2002 - Added SiteLocation
' ASH 19/2/2003 - Removed SiteLocale and SiteTimeZone
'--------------------------------------------------------------------------------
Dim oSiteProperties As ADODB.Recordset
Dim sSQL As String
    
    On Error GoTo ErrHandler

     ' PN 07/09/99
    ' upgrade to ado from dao
    sSQL = "select site, SiteDescription, SiteLocation, SiteCountry from site where site='" & msCode & "'"
    Set oSiteProperties = New ADODB.Recordset
    With oSiteProperties
        .Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
        If .RecordCount > 0 Then
            msCode = .Fields("Site")
            If IsNull(.Fields("SiteDescription")) Then
                msDescription = vbNullString
            Else
                msDescription = .Fields("SiteDescription")
            End If
            If IsNull(.Fields("SiteLocation")) Then
                mnLocation = SiteLocation.ESiteNoLocation
            Else
                mnLocation = .Fields("SiteLocation")
            End If
            
            'mlSiteLocale = ConvertFromNull(.Fields("SiteLocale").Value, vbLong)
            mlCountryCode = ConvertFromNull(.Fields("SiteCountry").Value, vbLong)
            'mlTimeZone = ConvertFromNull(.Fields("SiteTimeZone").Value, vbLong)
            
            Load = True
            mbNewRecord = False
        Else
            Load = False
        End If
    End With
           
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Load", "clsSite")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function
'--------------------------------------------
Public Function Save() As SiteSaveResult
'--------------------------------------------
' Save site details to database
'--------------------------------------------
' REVISIONS
' DPH 06/08/2002 - Added SiteLocation
' ASH 19/2/2003 - Removed SiteLocale and SiteTimeZone
'--------------------------------------------
Dim sSQL As String
    
    On Error GoTo ErrHandler
    
    If Not mbIsValid Then Err.Raise 454
    
    If mbNewRecord Then
        If Not DoesNameExist(msCode) Then
            ' commit the new site to the db
            
            ' PN 07/09/99
            ' upgrade to ado from dao
            ' NCJ 18/1/00 - Deal with quotes in description
            sSQL = "INSERT INTO Site(Site,SiteStatus, SiteDescription,SiteLocation,  " & _
                   "SiteCountry ) values('"
            sSQL = sSQL & msCode & "',0,'"
            sSQL = sSQL & ReplaceQuotes(msDescription) & "'," & mnLocation & ", " & _
                     mlCountryCode & ")"
            MacroADODBConnection.Execute sSQL
            
            Save = ESiteSaved
        Else
            Save = ESiteAlreadyExists
        End If
    Else
        ' commit the edited site details
        
        ' PN 07/09/99
        ' upgrade to ado from dao
        ' NCJ 6/7/00 SR3660 Use ReplaceQuotes around msDescription
'        sSQL = "update site set SiteDescription = '" & msDescription & "' where site ='" & msCode & "'"
        sSQL = "UPDATE Site SET SiteDescription = '" & ReplaceQuotes(msDescription) & "'" _
                & ", SiteLocation = " & mnLocation _
                & ", SiteCountry = " & ConvertToNull(mlCountryCode, vbLong) _
                & " WHERE Site = '" & msCode & "'"
        MacroADODBConnection.Execute sSQL, dbFailOnError
        
        Save = ESiteSaved
    End If
               
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "load", "clsSite")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select


End Function

'--------------------------------------------------------------------------------
Private Function DoesNameExist(sNewSite As String) As Boolean
'--------------------------------------------------------------------------------
' Check whether the site name already exists
'--------------------------------------------------------------------------------
Dim oExistingNames As ADODB.Recordset
Dim sSQL As String
    
    On Error GoTo ErrHandler
    
    sSQL = "SELECT * FROM Site WHERE Site = '" & sNewSite & "'"
        
    ' run the sql to determine if the SiteDescription exists in the site table
    Set oExistingNames = New ADODB.Recordset
    
    ' PN 07/09/99
    ' upgrade to ado from dao
    oExistingNames.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
    With oExistingNames
        If .RecordCount > 0 Then
            DoesNameExist = True
        Else
            DoesNameExist = False
        End If
        .Close
        
    End With
    
    Set oExistingNames = Nothing
               
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                    "DoesNameExist", "clsSite")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function

'--------------------------------------------------------------------------------
Public Property Get Code() As String
'--------------------------------------------------------------------------------
' The site code
'--------------------------------------------------------------------------------
    
    Code = msCode
    
End Property

'--------------------------------------------------------------------------------
Public Property Let Code(sCode As String)
'--------------------------------------------------------------------------------
' The site code (max. 8 chars)
'--------------------------------------------------------------------------------

    If Len(sCode) > 8 Then
        Err.Raise vbObjectError + 1001, , "String too long"
    End If
    
    ' check that the character entered is valid
    If Not gblnValidString(sCode, valAlpha + valNumeric) Then
        Err.Raise vbObjectError + 1001, , "Invalid string"
    End If

    ' check that the first char is not numeric
    If sCode <> vbNullString Then
        If gblnValidString(Left$(sCode, 1), valNumeric) Then
            Err.Raise vbObjectError + 1001, , "Invalid string"
        End If
    End If
    
    If sCode <> msCode Then
        msCode = sCode
        HasChanges = True
    End If
    
End Property

'--------------------------------------------------------------------------------
Public Property Get Name() As String
'--------------------------------------------------------------------------------
    
'--------------------------------------------------------------------------------
    
    Name = msDescription
    
End Property

'--------------------------------------------------------------------------------
Public Property Let Name(sDesc As String)
'--------------------------------------------------------------------------------
' Set the name (i.e. description) of a site
'--------------------------------------------------------------------------------

    If Len(sDesc) > 255 Then
        Err.Raise vbObjectError + 1001, , "String too long"
    End If
    
    ' check that the character entered is valid
    If Not gblnValidString(sDesc, valOnlySingleQuotes) Then
        Err.Raise vbObjectError + 1001, , "Invalid string"
    End If

    If sDesc <> msDescription Then
        ' It's changed
        msDescription = sDesc
        HasChanges = True
    End If

End Property

'--------------------------------------------------------------------------------
Public Property Get HasChanges() As Boolean
'--------------------------------------------------------------------------------
    
'--------------------------------------------------------------------------------
    
    HasChanges = mbHasChanges

End Property

'--------------------------------------------------------------------------------
Private Property Let HasChanges(bHasChanges As Boolean)
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
    If mbHasChanges <> bHasChanges Then
        mbHasChanges = bHasChanges
    End If
    Call Validate
    
End Property

'--------------------------------------------------------------------------------
Public Property Get IsValid() As Boolean
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------

    IsValid = mbIsValid
    
End Property

'--------------------------------------------------------------------------------
Private Sub Validate()
'--------------------------------------------------------------------------------
' Check we've got a non-empty site code & name
'--------------------------------------------------------------------------------
' REVISIONS
' DPH 07/08/2002 - Need to check if server site / if so need location
'--------------------------------------------------------------------------------
    
    If GetMacroDBSetting("datatransfer", "dbtype", , gsSERVER) = "server" Then
        If msDescription <> vbNullString And msCode <> vbNullString And mnLocation <> SiteLocation.ESiteNoLocation Then
            mbIsValid = True
        Else
            mbIsValid = False
        End If
    Else
        If msDescription <> vbNullString And msCode <> vbNullString Then
            mbIsValid = True
        Else
            mbIsValid = False
        End If
    End If
    
    ' NCJ 29/4/00 SR3357 - Commented out these lines
'    If Len(msDescription) > Len(msCode) Then
'        msDescription = msCode
'     End If
    
    'check locale, country and time zone regardless of server/web or remote
    'ASH 19/2/2003 Removed locale and time zone
    If CountryCode = 0 Then
        mbIsValid = False
    End If
    
    RaiseEvent IsValid(mbIsValid)
    
End Sub

'--------------------------------------------------------------------------------
Public Property Get Location() As Integer
'--------------------------------------------------------------------------------
    
'--------------------------------------------------------------------------------
    
    Location = mnLocation
    
End Property

'--------------------------------------------------------------------------------
Public Property Let Location(nLoc As Integer)
'--------------------------------------------------------------------------------
' Set the name (i.e. description) of a site
'--------------------------------------------------------------------------------

    If nLoc <> mnLocation Then
        ' It's changed
        mnLocation = nLoc
        HasChanges = True
    End If

End Property

'--------------------------------------------------------------------------------
Public Property Get SiteLocale() As Long
'--------------------------------------------------------------------------------
'retrieves site locale
'--------------------------------------------------------------------------------
    
    'SiteLocale = mlSiteLocale
    
End Property

'--------------------------------------------------------------------------------
Public Property Let SiteLocale(lSiteLocale As Long)
'--------------------------------------------------------------------------------
'set site locale
'--------------------------------------------------------------------------------
    
'    If lSiteLocale <> mlSiteLocale Then
'        mlSiteLocale = lSiteLocale
'        HasChanges = True
'    End If
    
End Property

'--------------------------------------------------------------------------------
Public Property Get CountryCode() As Long
'--------------------------------------------------------------------------------
'retrieves country code for the site
'--------------------------------------------------------------------------------

    CountryCode = mlCountryCode
    
End Property

'--------------------------------------------------------------------------------
Public Property Let CountryCode(lCountryCode As Long)
'--------------------------------------------------------------------------------
'set site country code
'--------------------------------------------------------------------------------

    If lCountryCode <> mlCountryCode Then
        mlCountryCode = lCountryCode
        HasChanges = True
    End If
    
End Property

'--------------------------------------------------------------------------------
Public Property Get TimeZone() As Long
'--------------------------------------------------------------------------------
'retrieves time zone for the site
'--------------------------------------------------------------------------------

    'TimeZone = mlTimeZone
    
End Property

'--------------------------------------------------------------------------------
Public Property Let TimeZone(lTimeZone As Long)
'--------------------------------------------------------------------------------
'set time zone for the site
'--------------------------------------------------------------------------------

'    If lTimeZone <> mlTimeZone Then
'        mlTimeZone = lTimeZone
'        HasChanges = True
'    End If
    
End Property

'--------------------------------------------------------------------------------
Private Sub Class_Initialize()
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------

    mbNewRecord = True
    mnLocation = SiteLocation.ESiteNoLocation
    
End Sub
