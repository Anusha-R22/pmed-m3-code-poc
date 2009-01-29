VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommunication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       clsCommunication.cls
'   Author:     Paul Norris 23/07/99
'   Purpose:    Configuration class for the communication settings
'               stored in the TrialOffice table in Macro.mdb
'               This class is utilised by the frmCommunicationConfiguration form.
'               It will not allow saving unless the correct type of data is supplied.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   PN  30/09/99    Commented code more thoroughly and DB access from conversion DAO to ADO
'   PN  30/09/99    Changed class names
'                   clsCommunicationData to clsCommunication
'                   because prog id is too long with original name
'   AN  11/12/99    Changed Validate such that User and Password are not required.
'                   ProxyServer was missing - added in now.
'   NCJ 18/2/00     Changed date fields to use Arezzo date parsing
'                   Local date values now stored in Arezzo and Double formats
'   NCJ 7/3/00      SRs2780,3110 Restructuring
'   NCJ 13/6/00     SR 3591 Take more account of TransferData flag
'   WillC 3/8/00 SR3648 Uncommented  2 lines in TrialOffice Let to disallow invalid chars in TrialOffice name
'   NCJ 24 Sep 01 - Changed frmArezzo to goArezzo (for MACRO 2.2)
'   TA 26/9/01:     Changes following db audit
'   TA 05/10/2001:  Changed password to UserPassword
'------------------------------------------------------------------------------------
Option Explicit

' tell the container when the object is valid or invalid
Public Event IsValid(bIsValid As Boolean)

' internal properties
' NCJ 18/2/00 - Now in Arezzo date format
Private msArezzoEffectiveFrom As String
Private msArezzoEffectiveTo As String
' NCJ 18/2/00 - Store dates as doubles, too
Private mdblEffectiveFrom As Double
Private mdblEffectiveTo As Double

' NCJ 7/3/00 - Store the original values of DB key fields
Private msStoredTrialOffice As String
Private msStoredSite As String
Private mdblStoredEffectiveFrom As Double
Private mdblStoredEffectiveTo As Double
' NCJ 13/6/00
Private msStoredTransferData As String

Private msTrialOffice As String
Private msHTTPAddress As String
Private msSite As String
Private msUser As String
Private msPassword As String
Private msTransferData As String
Private msPortNumber As String
'   ATN 11/12/99
Private msProxyServer As String

' internal object management properties
Private mbUpdateRecord As Boolean
Private mbHasChanges As Boolean
Private mbIsValid As Boolean
Private mbEnableWriteOnceProperties As Boolean

Private msKey As String

'------------------------------------------------------------------------------------
Public Property Get CommKey() As String
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------

    CommKey = msKey
    
End Property

'------------------------------------------------------------------------------------
Public Property Let CommKey(sKey As String)
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------

    msKey = sKey
    
End Property

'------------------------------------------------------------------------------------
Public Property Get HasChanges() As Boolean
'------------------------------------------------------------------------------------
' allow the container to ask if changes have been made
'------------------------------------------------------------------------------------

    HasChanges = mbHasChanges
    
End Property

'------------------------------------------------------------------------------------
Public Property Get IsValid() As Boolean
'------------------------------------------------------------------------------------
' allo the container to ask if the chagnes are currently valid
'------------------------------------------------------------------------------------

    IsValid = mbIsValid
    
End Property

'------------------------------------------------------------------------------------
Private Sub Validate()
'------------------------------------------------------------------------------------
' object validation is simple so it is carried out here
' it is called each time a property is set
'------------------------------------------------------------------------------------
    ' all these fields are mandatory
'   ATN 11/12/99
'   User and password are not required.  Transferschedule not required.
    On Error GoTo ErrHandler

    If msTrialOffice <> vbNullString And _
       msArezzoEffectiveTo <> vbNullString And _
       msArezzoEffectiveFrom <> vbNullString And _
       msHTTPAddress <> vbNullString And _
       msSite <> vbNullString And _
       CStr(msTransferData) <> vbNullString And _
       CStr(msPortNumber) <> vbNullString Then
       
        mbIsValid = True
        
    Else
        mbIsValid = False
        
    End If
    
    ' this allows the contianer to enable the ok buttons
    ' it also acts as a HasChanges event since it is only ever
    ' called when changes have been made
    RaiseEvent IsValid(mbIsValid)

    Exit Sub
 
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "Validate", "clsCommunication")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'------------------------------------------------------------------------------------
Public Function HasKeyChanged() As Boolean
'------------------------------------------------------------------------------------
' Return TRUE if any of "key" fields has changed (i.e.TrialOffice, EffectiveFrom and EffectiveTo)
' Return FALSE otherwise
' NCJ 13/6/00 - SR3591 Also consider TransferData
'------------------------------------------------------------------------------------

    If mdblEffectiveFrom = mdblStoredEffectiveFrom _
     And mdblEffectiveTo = mdblStoredEffectiveTo _
     And msTransferData = msStoredTransferData _
     And msTrialOffice = msStoredTrialOffice Then
        HasKeyChanged = False
    Else
        HasKeyChanged = True
    End If

End Function

'------------------------------------------------------------------------------------
Public Property Let StoredTrialOffice(sOffice As String)
'------------------------------------------------------------------------------------
' The original DB value of TrialOffice
'------------------------------------------------------------------------------------

    msStoredTrialOffice = sOffice
    
End Property

'------------------------------------------------------------------------------------
Public Property Get StoredTrialOffice() As String
'------------------------------------------------------------------------------------
' The original DB value of TrialOffice
'------------------------------------------------------------------------------------

    StoredTrialOffice = msStoredTrialOffice
    
End Property

'------------------------------------------------------------------------------------
Public Property Let StoredSite(sSite As String)
'------------------------------------------------------------------------------------
' The original DB value of Site
'------------------------------------------------------------------------------------

    msStoredSite = sSite
    
End Property

'------------------------------------------------------------------------------------
Public Property Get StoredSite() As String
'------------------------------------------------------------------------------------
' The original DB value of Site
'------------------------------------------------------------------------------------

    StoredSite = msStoredSite
    
End Property

'------------------------------------------------------------------------------------
Public Property Get TrialOffice() As String
'------------------------------------------------------------------------------------

    TrialOffice = msTrialOffice
    
End Property

'------------------------------------------------------------------------------------
Public Property Let TrialOffice(sValue As String)
'------------------------------------------------------------------------------------
' The TrialOffice property
' Can only be set once, i.e. if there's no stored value
'------------------------------------------------------------------------------------

    If msStoredTrialOffice = "" Then
        If Len(sValue) > 255 Then
            ' string too long for database field
            Err.Raise vbObjectError + 1001, , "String too long"
        'WillC 3/8/00 SR3648 Uncommented these 2 lines to disallow invalid chars in TrialOffice name
        ElseIf Not gblnValidString(sValue, valAlpha) Then
            Err.Raise vbObjectError + 1001, , "Invalid string"
            
        End If
        
        If sValue <> msTrialOffice Then
            ' store the value and update the Changes internal variable
            msTrialOffice = sValue
            HasChanges = True
        End If
    End If
        

'    If mbEnableWriteOnceProperties Then
'        ' this property is enabled
'        ' normally it will not be but when the record is new it has
'        ' to be populated
'        If Len(sValue) > 255 Then
'            ' string too long for database field
'            Err.Raise vbObjectError + 1001, , "String too long"
'
'        End If
'
'        If sValue <> msTrialOffice Then
'            ' store the value and update the changes internal variable
'            msTrialOffice = sValue
'            HasChanges = True
'        End If
'
'    Else
''   ATn 17/12/99
''   Need to allow edits when creating a new record
'        ' raise the read only property error
'        ' because this property should not be writable once saved
''        Err.Raise 383
'        If sValue <> msTrialOffice Then
'            ' store the value and update the changes internal variable
'            msTrialOffice = sValue
'            HasChanges = True
'        End If
'    End If
    
End Property

'------------------------------------------------------------------------------------
Public Property Let DblEffectiveFrom(dblDate As Double)
'------------------------------------------------------------------------------------
' Can set date as either string or double
' Assume dblDate valid
'------------------------------------------------------------------------------------

    mdblEffectiveFrom = dblDate
    If dblDate <> 0 Then
        msArezzoEffectiveFrom = ConvertDateToProformaSyntax(dblDate)
    Else
        msArezzoEffectiveFrom = ""
    End If
    
End Property

'------------------------------------------------------------------------------------
Public Property Get DblEffectiveFrom() As Double
'------------------------------------------------------------------------------------
' NCJ 7 Mar 00
' Get EffectiveFrom as a double
'------------------------------------------------------------------------------------

    DblEffectiveFrom = mdblEffectiveFrom

End Property

'------------------------------------------------------------------------------------
Public Property Let StoredEffectiveFrom(dblDate As Double)
'------------------------------------------------------------------------------------
' Set stored (DB) value for EffectiveFrom as double
' Assume dblDate valid
'------------------------------------------------------------------------------------

    mdblStoredEffectiveFrom = dblDate
    
End Property

'------------------------------------------------------------------------------------
Public Property Get StoredEffectiveFrom() As Double
'------------------------------------------------------------------------------------
' NCJ 7 Mar 00
' Get stored (DB) EffectiveFrom as a double
'------------------------------------------------------------------------------------

    StoredEffectiveFrom = mdblStoredEffectiveFrom

End Property

'------------------------------------------------------------------------------------
Public Property Let DblEffectiveTo(dblDate As Double)
'------------------------------------------------------------------------------------
' Can set date as either string or double
' Assume dblDate valid
'------------------------------------------------------------------------------------

    mdblEffectiveTo = dblDate
    If dblDate <> 0 Then
        msArezzoEffectiveTo = ConvertDateToProformaSyntax(dblDate)
    Else
        msArezzoEffectiveTo = ""
    End If
    
End Property

'------------------------------------------------------------------------------------
Public Property Get DblEffectiveTo() As Double
'------------------------------------------------------------------------------------
' NCJ 7 Mar 00
' Get EffectiveTo as a double
'------------------------------------------------------------------------------------

    DblEffectiveTo = mdblEffectiveTo

End Property

'------------------------------------------------------------------------------------
Public Property Let StoredEffectiveTo(dblDate As Double)
'------------------------------------------------------------------------------------
' Set stored (DB) value for EffectiveTo as double
' Assume dblDate valid
'------------------------------------------------------------------------------------

    mdblStoredEffectiveTo = dblDate
    
End Property

'------------------------------------------------------------------------------------
Public Property Get StoredEffectiveTo() As Double
'------------------------------------------------------------------------------------
' NCJ 7 Mar 00
' Get stored (DB) EffectiveTo as a double
'------------------------------------------------------------------------------------

    StoredEffectiveTo = mdblStoredEffectiveTo

End Property

'------------------------------------------------------------------------------------
Public Property Get EffectiveFrom() As String
'------------------------------------------------------------------------------------
' NCJ 18/2/00 We now store the date in Arezzo format,
' but return the "correctly" formatted value
'------------------------------------------------------------------------------------

    If msArezzoEffectiveFrom > "" Then
        EffectiveFrom = goArezzo.FormatDate(msArezzoEffectiveFrom, frmMenu.DefaultDateFormat)
    Else
        EffectiveFrom = ""
    End If
    
End Property

'------------------------------------------------------------------------------------
Public Property Let EffectiveFrom(sDate As String)
'------------------------------------------------------------------------------------
' Wrap up the validation for the EffectiveFrom property
' sDate is String date value (as entered by user)
' NCJ 18/2/00 - Use new Arezzo date parsing and store double value here
' Store our local value as an Arezzo date
'------------------------------------------------------------------------------------
Dim sArezzoDate As String
Dim sNewDate As String

    sNewDate = Trim(sDate)
    If sNewDate > "" Then
        ' NCJ Read date using current default date format
        sNewDate = goArezzo.ReadValidDate(sNewDate, frmMenu.DefaultDateFormat, sArezzoDate)
'        sNewDate = frmArezzo.ReadValidDate(sNewDate, frmMenu.DefaultDateFormat, sArezzoDate)
        If sNewDate = "" Then
            Err.Raise 13
        Else
            ' See if it's changed
            If sArezzoDate <> msArezzoEffectiveFrom Then
                ' Date was accepted OK
                ' Convert to Double value, stripping off time of day
                mdblEffectiveFrom = CLng(goArezzo.ArezzoDateToDouble(sArezzoDate))
                ' Store in Arezzo format
                msArezzoEffectiveFrom = sArezzoDate
                HasChanges = True
            End If
        End If
    End If
    
End Property

'------------------------------------------------------------------------------------
Public Property Get EffectiveTo() As String
'------------------------------------------------------------------------------------
' NCJ 18/2/00 We now store the date in Arezzo format,
' but return the "correctly" formatted value
'------------------------------------------------------------------------------------

    If msArezzoEffectiveTo > "" Then
        EffectiveTo = goArezzo.FormatDate(msArezzoEffectiveTo, frmMenu.DefaultDateFormat)
    Else
        EffectiveTo = ""
    End If

'    EffectiveTo = msEffectiveTo
    
End Property

'------------------------------------------------------------------------------------
Public Property Let EffectiveTo(sDate As String)
'------------------------------------------------------------------------------------
' Wrap up the validation for the EffectiveTo property
' sDate is String date value (as entered by user)
' NCJ 18/2/00 - Use new Arezzo date parsing and store double value here
' Store our local value as an Arezzo date
'------------------------------------------------------------------------------------
Dim sArezzoDate As String
Dim sNewDate As String

    sNewDate = Trim(sDate)
    If sNewDate > "" Then
        ' NCJ Read date using current default date format
        sNewDate = goArezzo.ReadValidDate(sNewDate, frmMenu.DefaultDateFormat, sArezzoDate)
        If sNewDate = "" Then
            Err.Raise 13
        Else
            ' Date was accepted OK
            ' See if it's changed
            If sArezzoDate <> msArezzoEffectiveTo Then
                ' Convert to Double value, stripping off the time of day
                mdblEffectiveTo = CLng(goArezzo.ArezzoDateToDouble(sArezzoDate))
                ' Store in Arezzo format
                msArezzoEffectiveTo = sArezzoDate
                HasChanges = True
            End If
        End If
    End If
    
End Property

'------------------------------------------------------------------------------------
Public Property Get HTTPAddress() As String
'------------------------------------------------------------------------------------

    HTTPAddress = msHTTPAddress
    
End Property

'------------------------------------------------------------------------------------
Public Property Let HTTPAddress(sValue As String)
'------------------------------------------------------------------------------------
' wrap up the validation for the HTTPAddress property
'------------------------------------------------------------------------------------
    If Len(sValue) > 255 Then
        ' string too long for database field
        Err.Raise vbObjectError + 1001, , "String too long"
        
    End If
    
    If sValue <> msHTTPAddress Then
        ' store the value
        msHTTPAddress = sValue
        HasChanges = True
        
    End If
    
End Property

'------------------------------------------------------------------------------------
Public Property Get ProxyServer() As String
'------------------------------------------------------------------------------------

    ProxyServer = msProxyServer
    
End Property

'------------------------------------------------------------------------------------
Public Property Let ProxyServer(sValue As String)
'------------------------------------------------------------------------------------
' wrap up the validation for the HTTPAddress property
'------------------------------------------------------------------------------------
    If Len(sValue) > 255 Then
        ' string too long for database field
        Err.Raise vbObjectError + 1001, , "String too long"
        
    End If
    
    If sValue <> msProxyServer Then
        ' store the value
        msProxyServer = sValue
        HasChanges = True
        
    End If
    
End Property

'------------------------------------------------------------------------------------
Public Property Get Site() As String
'------------------------------------------------------------------------------------

    Site = msSite
    
End Property

'------------------------------------------------------------------------------------
Public Property Let Site(sValue As String)
'------------------------------------------------------------------------------------
' The Site property
' This can only be set if there's no stored value
'------------------------------------------------------------------------------------
    
    If msStoredSite = "" Then
        ' NCJ 7/3/00 - Restrict site codes to 8 characters
        If Len(sValue) > 8 Then
            ' string too long for database field
            ' NCJ 8/3/00 - But still add a truncated version (for backwards compatibility)
            sValue = Left(sValue, 8)
'            Err.Raise vbObjectError + 1001, , "Site code too long"
        ' NCJ 8/3/00 - Allow AlphaNumeric
        ElseIf Not gblnValidString(sValue, valAlpha + valNumeric) Then
            Err.Raise vbObjectError + 1001, , "Invalid string"
        
        End If
        
        If sValue <> msSite Then
            msSite = sValue
            HasChanges = True
        End If

    End If
    
'    If mbEnableWriteOnceProperties Then
'        ' this property is enabled
'        ' normally it will not be but when the record is new it has
'        ' to be populated
'        If Len(sValue) > 255 Then
'            ' string too long for database field
'            Err.Raise vbObjectError + 1001, , "String too long"
'        End If
'
'        If sValue <> msSite Then
'            msSite = sValue
'            HasChanges = True
'        End If
'
'    Else
''   ATn 17/12/99
''   Need to allow edits when creating a new record
'        ' raise the read only property error
'        ' because this property should not be writable once saved
''        Err.Raise 383
'        If sValue <> msSite Then
'            msSite = sValue
'            HasChanges = True
'        End If
'    End If
    
End Property

'------------------------------------------------------------------------------------
Public Property Get User() As String
'------------------------------------------------------------------------------------

    User = msUser
    
End Property

'------------------------------------------------------------------------------------
Public Property Let User(sValue As String)
'------------------------------------------------------------------------------------
' wrap up the validation for the User property
'------------------------------------------------------------------------------------
    If Len(sValue) > 50 Then
        ' string too long for database field
        Err.Raise vbObjectError + 1001, , "String too long"
        
    End If
    
    ' check that the character entered is valid
    If Not gblnValidString(sValue, valAlpha + valNumeric) Then
        Err.Raise vbObjectError + 1001, , "Invalid string"
        
    End If
    
    If sValue <> msUser Then
        ' store the value
        msUser = sValue
        HasChanges = True
        
    End If
    
End Property

'------------------------------------------------------------------------------------
Public Property Get Password() As String
'------------------------------------------------------------------------------------

    Password = msPassword
    
End Property

'------------------------------------------------------------------------------------
Public Property Let Password(sValue As String)
'------------------------------------------------------------------------------------
' wrap up the validation for the Password property
'------------------------------------------------------------------------------------
    If Len(sValue) > 50 Then
        ' string too long for database field
        Err.Raise vbObjectError + 1001, , "String too long"
        
    End If
    
    ' check that the character entered is valid
    If Not gblnValidString(sValue, valAlpha + valNumeric) Then
        Err.Raise vbObjectError + 1001, , "Invalid string"
        
    End If

    If sValue <> msPassword Then
        ' store the value
        msPassword = sValue
        HasChanges = True
        
    End If
    
End Property

'------------------------------------------------------------------------------------
Public Property Let StoredTransferData(sTransfer As String)
'------------------------------------------------------------------------------------
' The value of TransferData in the database
' NB This is a string even though it can be "0" or "1"
'------------------------------------------------------------------------------------

    msStoredTransferData = sTransfer
    
End Property

'------------------------------------------------------------------------------------
Public Property Get TransferData() As Integer
'------------------------------------------------------------------------------------
'changed by Mo Morris 14/12/99, from type string to integer
'------------------------------------------------------------------------------------

'   ATN 17/12/99
'   Need to set default value
    If msTransferData = "" Then
        msTransferData = "0"
    End If
    
    TransferData = msTransferData
    
    
End Property

'------------------------------------------------------------------------------------
Public Property Let TransferData(sValue As Integer)
'------------------------------------------------------------------------------------
'wrap up the validation for the TransferData property
'changed by Mo Morris 14/12/99, from type string to integer
'------------------------------------------------------------------------------------
    ' check that the character entered is valid
    If Not gblnValidString(sValue, valNumeric) Then
        Err.Raise vbObjectError + 1001, , "Invalid number"
        
    End If
    
    ' check if it is null
    'changed by Mo Morris 14/12/99, from vbNullString to Null
    If sValue <> Null Then
        sValue = CInt(sValue)
        
    End If
    
    If CStr(sValue) <> msTransferData Then
        ' store the value
        msTransferData = sValue
        HasChanges = True
        
    End If
    
End Property

'------------------------------------------------------------------------------------
Public Property Get PortNumber() As String
'------------------------------------------------------------------------------------

    PortNumber = msPortNumber
    
End Property

'------------------------------------------------------------------------------------
Public Property Let PortNumber(sValue As String)
'------------------------------------------------------------------------------------
' wrap up the validation for the PortNumber property
'------------------------------------------------------------------------------------
    ' check that the character entered is valid
    If Not gblnValidString(sValue, valNumeric) Then
        Err.Raise vbObjectError + 1001, , "Invalid number"
    End If
    
    ' check if it is null
    If sValue <> vbNullString Then
        sValue = CInt(sValue)
    End If
    
    If sValue <> msPortNumber Then
        ' store the value
        msPortNumber = sValue
        HasChanges = True
    End If
End Property

'------------------------------------------------------------------------------------
Public Property Let HasChanges(bHasChanges As Boolean)
'------------------------------------------------------------------------------------
    
    If mbHasChanges <> bHasChanges Then
        mbHasChanges = bHasChanges
    End If
    Call Validate
    
End Property

'------------------------------------------------------------------------------------
Public Sub Load(Optional sTrialOffice As String = vbNullString, _
                Optional sSite As String = vbNullString)
'------------------------------------------------------------------------------------
' read all data from the database
'------------------------------------------------------------------------------------
Dim rsProperties As ADODB.Recordset
Dim sSQL As String
Dim sWhere As String
    
    ' read the properties
    sSQL = "select * from TrialOffice"
    If sTrialOffice <> vbNullString Then
        sWhere = " where trialoffice='" & sTrialOffice & "'"
        msStoredTrialOffice = sTrialOffice
    Else
        msStoredTrialOffice = ""
    End If
    If sSite <> vbNullString Then
        If sWhere <> vbNullString Then
            sWhere = sWhere & " and Site='" & sSite & "'"
        Else
            sWhere = sWhere & " where Site='" & sSite & "'"
        End If
    End If
    If sWhere <> vbNullString Then
        sSQL = sSQL & sWhere
    End If
    
    ' PN 30/09/99 change db access from DAO to ADO
    Set rsProperties = New ADODB.Recordset
    rsProperties.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    With rsProperties
        If .RecordCount > 0 Then
            mbUpdateRecord = True
            If Not IsNull(.Fields("EffectiveFrom")) Then
'                msEffectiveFrom = Format(CDate(.Fields("EffectiveFrom")), "dd/mm/yyyy")
                ' NCJ 18/2/00
                mdblEffectiveFrom = .Fields("EffectiveFrom")
                msArezzoEffectiveFrom = ConvertDateToProformaSyntax(mdblEffectiveFrom)
            Else
                msArezzoEffectiveFrom = ""
                mdblEffectiveFrom = 0
            End If
            ' NCJ 7/3/00 - Store original value
            mdblStoredEffectiveFrom = mdblEffectiveFrom
            
            If Not IsNull(.Fields("EffectiveTo")) Then
'                msEffectiveTo = Format(CDate(.Fields("EffectiveTo")), "dd/mm/yyyy")
                ' NCJ 18/2/00
                mdblEffectiveTo = .Fields("EffectiveTo")
                msArezzoEffectiveTo = ConvertDateToProformaSyntax(mdblEffectiveTo)
            Else
                msArezzoEffectiveTo = ""
                mdblEffectiveTo = 0
            End If
            ' NCJ 7/3/00 - Store original value
            mdblStoredEffectiveTo = mdblEffectiveTo
            
            If Not IsNull(.Fields("TransferData")) Then
                msTransferData = .Fields("TransferData")
                ' NCJ 13/6/00
                msStoredTransferData = msTransferData
            End If
            
            If Not IsNull(.Fields("Site")) Then
                msSite = .Fields("Site")
            End If
            
            If Not IsNull(.Fields("HTTPAddress")) Then
                msHTTPAddress = .Fields("HTTPAddress")
            End If
            
            If Not IsNull(.Fields("ProxyServer")) Then
                msProxyServer = .Fields("ProxyServer")
            End If
            
            If Not IsNull(.Fields("TrialOffice")) Then
                msTrialOffice = .Fields("TrialOffice")
            End If
            
            If Not IsNull(.Fields("PortNumber")) Then
                msPortNumber = .Fields("PortNumber")
            End If
            
            If Not IsNull(.Fields("UserName")) Then
                msUser = .Fields("UserName")
            End If
            
            If Not IsNull(.Fields("UserPassword")) Then
                msPassword = .Fields("UserPassword")
            End If
        
        Else
            mbUpdateRecord = False
        End If
        .Close
        
    End With
    
    ' if the site is blank then allow setting it
    If msSite = vbNullString Then
        mbEnableWriteOnceProperties = True
    End If
    
    HasChanges = False
    Set rsProperties = Nothing

End Sub

'------------------------------------------------------------------------------------
Public Sub Save()
'------------------------------------------------------------------------------------
' save data to database
' NCJ 18/2/00 Store dates dircetly from local Double values
'------------------------------------------------------------------------------------
Dim sSQL As String
        
    On Error GoTo ErrHandler
    
    If mbIsValid Then
        ' save the properties
        
        'Begin transaction
        TransBegin
        
        'changed by Mo Morris 14/12/99, single quotes removed from msTransferData in following SQL statement
'        If mbUpdateRecord = False Then
        If msStoredTrialOffice = "" Then
            ' It's a new record
            sSQL = "insert into trialoffice (EffectiveFrom, EffectiveTo,  "
            sSQL = sSQL & "TransferData,TrialOffice,HTTPAddress,ProxyServer,Site,PortNumber, UserName, UserPassword) "
            sSQL = sSQL & "values (" & ConvertLocalNumToStandard(CStr(mdblEffectiveFrom)) _
                    & ", " & ConvertLocalNumToStandard(CStr(mdblEffectiveTo)) & ", "
            sSQL = sSQL & msTransferData & ", '" & msTrialOffice & "', '" & msHTTPAddress & "', '" & msProxyServer & "', '"
            sSQL = sSQL & msSite & "', " & msPortNumber & ", '" & msUser & "', '" & msPassword & "' )"
            MacroADODBConnection.Execute sSQL
            mbEnableWriteOnceProperties = False
            
        Else
            sSQL = "UPDATE trialoffice SET PortNumber=" & msPortNumber
            sSQL = sSQL & ", UserName='" & msUser & "',"
            sSQL = sSQL & " UserPassword='" & msPassword & "',"
            sSQL = sSQL & " TransferData=" & msTransferData & ","
            sSQL = sSQL & " HTTPAddress='" & msHTTPAddress & "',"
            sSQL = sSQL & " ProxyServer='" & msProxyServer & "',"
            sSQL = sSQL & " EffectiveFrom= " & ConvertLocalNumToStandard(CStr(mdblEffectiveFrom)) & ","
            sSQL = sSQL & " EffectiveTo= " & ConvertLocalNumToStandard(CStr(mdblEffectiveTo))
            sSQL = sSQL & " WHERE site='" & msSite & "' AND trialoffice='" & msTrialOffice & "'"
            ' NCJ 7/3/00 SR3110 - Must also match on original date values to avoid updating more than one field
            sSQL = sSQL & " AND EffectiveFrom = " & ConvertLocalNumToStandard(CStr(mdblStoredEffectiveFrom))
            sSQL = sSQL & " AND Effectiveto = " & ConvertLocalNumToStandard(CStr(mdblStoredEffectiveTo))
            MacroADODBConnection.Execute sSQL
    
        End If
        
        'End transaction
        TransCommit
        
        HasChanges = False
        
    End If

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                    "Save", "clsCommunication")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select


End Sub

'------------------------------------------------------------------------------------
Private Sub Class_Initialize()
'------------------------------------------------------------------------------------

    mbEnableWriteOnceProperties = False
    mbIsValid = False
    
End Sub
