VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WWWIOEform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------
'   File:       WWWIOEform.cls
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     ian curtis, august 2003
'   Purpose:    interface to html eform generating code
'----------------------------------------------------------------------------
'   revisions
'   ic 15/09/2003 added CheckArezzoEventsA() and SaveArezzoEventsA()
'   ic 16/03/2004 remove conditional ORAMA compilation
'   ic 05/07/2004 added error handling to each routine
'----------------------------------------------------------------------------
Option Explicit

'----------------------------------------------------------------------------
Public Function LoadEform(ByRef oUser As MACROUser, ByRef oSubject As StudySubject, _
    ByVal sSiteCode As String, ByVal lCRFPageTaskId As Long, ByRef sEFILockToken As String, _
    ByRef sVILockToken As String, ByRef bEFIUnavailable As Boolean, ByVal sDecimalPoint As String, _
    ByVal sThousandSeparator As String, Optional ByVal vAlerts As Variant, Optional ByVal vErrors As Variant, _
    Optional ByVal bAutoNext As Boolean = False) As String
'----------------------------------------------------------------------------
' ic 18/08/2003
' function requests eform html code for specified eform
' revisions
' ic 05/07/2004 added error handling
'----------------------------------------------------------------------------
    On Error GoTo CatchAllError

    LoadEform = GetEformBody(oUser, oSubject, oUser.DatabaseCode, sSiteCode, lCRFPageTaskId, sEFILockToken, sVILockToken, _
        bEFIUnavailable, sDecimalPoint, sThousandSeparator, eInterface.iwww, vAlerts, vErrors, bAutoNext)
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|WWWIOEform.LoadEform")
End Function

'----------------------------------------------------------------------------
Public Function SaveEform(ByRef oUser As MACROUser, ByRef oSubject As StudySubject, ByVal sCRFPageTaskId As String, _
                          ByVal sForm As String, ByRef sEFILockToken As String, ByRef sVILockToken As String, _
                          ByVal bVReadOnly As Boolean, ByVal bEReadOnly As Boolean, ByVal sLabCode As String, _
                          ByVal sDecimalPoint As String, ByVal sThousandSeparator As String, ByRef sRegister As String, _
                          ByVal sLocalDate As String, Optional ByVal nTimezoneOffset As Integer = 0) As Variant
'----------------------------------------------------------------------------
' ic 18/08/2003
' function calls saveform function
' revisions
' ic 05/07/2004 added error handling
'----------------------------------------------------------------------------
    On Error GoTo CatchAllError

    SaveEform = SaveForm(oUser, oSubject, sCRFPageTaskId, sForm, sEFILockToken, sVILockToken, _
        bVReadOnly, bEReadOnly, sLabCode, sDecimalPoint, sThousandSeparator, sRegister, sLocalDate, nTimezoneOffset)
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|WWWIOEform.SaveEform")
End Function

'ic 16/03/2004 remove conditional compilation
'#If ORAMA = 1 Then
'----------------------------------------------------------------------------
Public Function CheckArezzoEventsA(ByRef oSubject As StudySubject, ByVal sDatabase As String, _
    ByVal sEformPageTaskId As String, ByVal sNext As String, ByRef bArezzoEvents As Boolean) As String
'----------------------------------------------------------------------------
'   ic 12/09/2003
'   function calls CheckArezzoEvents function
'   revisions
'   ic 05/07/2004 added error handling
'----------------------------------------------------------------------------
    On Error GoTo CatchAllError
    
    CheckArezzoEventsA = CheckArezzoEvents(oSubject, sDatabase, sEformPageTaskId, sNext, bArezzoEvents)
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|WWWIOEform.CheckArezzoEventsA")
End Function

'----------------------------------------------------------------------------
Public Sub SaveArezzoEventsA(ByRef oSubject As StudySubject, ByVal sForm As String)
'----------------------------------------------------------------------------
'   ic 12/09/2003
'   function calls SaveArezzoEvents function
'   revisions
'   ic 05/07/2004 added error handling
'----------------------------------------------------------------------------
    On Error GoTo CatchAllError

    Call SaveArezzoEvents(oSubject, sForm)
    Exit Sub
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|WWWIOEform.SaveArezzoEventsA")
End Sub
'#End If
