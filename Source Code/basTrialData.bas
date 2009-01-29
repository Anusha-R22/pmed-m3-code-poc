Attribute VB_Name = "basTrialData"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998-2005. All Rights Reserved
'   File:       basTrialData.bas
'   Module:     basTrialData
'   Author:     Andrew Newbigging, June 1997
'   Purpose:    Common SQL functions used throughout MACRO for trial related data.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   1   Andrew Newbigging       4/06/97
'   2   Andrew Newbigging       19/06/97
'   3   Andrew Newbigging       4/07/97
'   4   Mo Morris               4/07/97
'   5   Andrew Newbigging       17/07/97
'   6   Andrew Newbigging       10/09/97
'   7   Andrew Newbigging       24/09/97
'   8   Andrew Newbigging       24/09/97
'   9   Mo Morris               24/09/97
'   10  Mo Morris               26/09/97
'   11  Mo Morris               27/11/97
'   12  Andrew Newbigging       26/01/98
'   13  Joanne Lau              8/05/98
'   14  Andrew Newbigging       15/05/98
'   15  Joanne Lau              29/07/98
'   16  Andrew Newbigging       11/11/98
'       Following routines moved into this module:
'           InsertTrial
'           CopyCRFPage,gdsCRFPage,gblnCRFPageExists
'           gdsCRFPageElementList,mnCopyCRFElement,DeleteCRFPage,mnNextCRFElementId
'           CopyDataItem,gdsDataItem,DataItemExists,gnNextDataItemId
'   17  Andrew Newbigging       16/3/99 SR 738
'       Modified CopyDataItem to cope with single quote in data item name
'   18  Paul Norris             05/08/99 SR 648
'       Added gblnDataItemExportExists() to check for duplicate DataItemExport names
'       In InsertTrial In writing to table ClinicalTrial, Sponsor removed from SQL and TrialTypeId added
'   19  NCJ 12/8/99
'       Updates to use CLM
'   20  Paul Norris     16/08/99
'       Updated CopyDataItem() to copy new fields in a data item with validations and categories
'   22  Paul Norris     27/08/99
'       Updated InsertTrial() to add defaults for stand. time and date
'   22  Paul Norris     02/09/99    Amended InsertTrial() to first close trial before
'                                   calling CreateProformaTrial() to resolve bug when
'                                   creating new trial if a trial is already open
'   23  PN          08/09/99    Changed Field Value to SpecialValue
'   24  NCJ 8 Sep 99
'       Removed calls to InsertProformaCRFElement
'   25  WillC       09/09/99    Added tables to gDeleteTrial
'   25  PN          13/09/99    Changed field name DataItem.Case to DataItemCase
'   PN  17/09/99 Changed parameters FromClinicalTrialId to ToClinicalTrialId and
'                FromVersionId to ToVersionId in CopyDataItem()
'   26  Willc       17/09/99    Added function gdsValidation to match ValidationAcitionNAmes
'       to ValidationType names due to database change (new table)
'   NCJ 17 Sept 99
'       Added DeleteProformaTrial to gDeleteTrial
'   20/9/99 Mo Morris
'   gDeleteTrial split into two routines together with additional files being deleted:-
'   DeleteTrialSD, which deletes a Study Definition
'   and DeleteTrialPRD, which deletes the Patient Response data within a Trial.
'   If both are required call DeleteTrialPRD before DeleteTrialSD
'   PN  23/09/99    Moved CopyValidations() and CopyCategories(), DeleteCRFPage(),
'                   DeleteTrialPRD(), DeleteTrialSD(), gblnCRFPageExists()
'                   gdsCRFPageElementList(), gdsDataItem(), CopyDataItem(), CopyCRFPage()
'                   InsertTrial(), mnCopyCRFElement()
'                   to modSDTrialData
'   PN 24/09/99 Moved DeleteTrialPRD() and DeleteTrialSD() from modSDTrialData module
'   Mo Morris   1/11/99
'   DAO to ADO conversion
'   SDM 10/11/99    Copied in error handling rountines
'   NCJ 30/11/99    Function GetCRFPageDateLabel
'                   Save timestamps as doubles
'   NCJ 7/12/99     Reordered elements in gdsDataValues
'   NCJ 11/12/99    Ids to Long instead of Integer
'   NCJ 13/12/99    Commented out unused routines gnNewTrialId and gnNextDataItemId
'   Mo Morris 29/4/00   Added the functions TrialNameFromId, VisitCodeFromId, CRFPageCodeFromTaskId
'                   and DataItemCodeFromTaskId
'   NCJ 19/5/00     Added DataItemNameFromTaskId
'   NCJ 7/8/00      Added Function SubjectLabelFromTrialSiteId
'   Mo Morris   4/10/00
'                   CurrentVersionId, CheckDataItemLength  and gnNewTrialId removed.
'                   End changed to Call MacroEnd in error trapping.
'                   'Clinical Trial'changed to 'Study' in glog messages.
'                   Created new functions DataItemCodeFromId and DataItemNameFromId, both
'                   of which call new function DataItemFromId
'   Mo Morris   10/10/00 new functions SiteDescriptionFromSite, TrialDescriptionFromName added
'   NCJ 19/10/00 - Removed NormalRange, RandomisationStep and StratificationFactor tables from Delete trial
'                   Also include MIMessage in DeleteTrialPRD
'                   Removed unused gdsDataFormat, gdsValidation
'   NCJ 26/10/00 - Keep trial site deletion in DeleteTrialPRD
'   NCJ 27/10/00 - Added GetSQLStringEquals
'   NCJ 30/10/00 - Moved GetSQLStringLike and GetSQLStringEquals to basCommon
'   Mo Morris   12/12/00 DeleteTrialSD re-wriiten
'   Mo Morris 30/8/01 Db Audit. New function NextMessageId, required due to table
'                   Message no longer having an AutoNumber on field MessageId
'   Mo Morris   8/1/02  New function CRFPageCodeFromId added
'   Mo Morris   19/2/02 New function DataTypeFromId added
'   Mo Morris   25/4/2002, "On Error GoTo ErrHandler" added to several functions where it was missing
'   Mo Morris   8/5/2002    General clean up of DeleteTrialSD
'   REM 13/05/02 - in gdsCRFPage changed to a readonly recordset that allows recordcount
'   REM 03/07/02 - add 3 new routines for RQG's - QGroupExists, QGroupFromId, NewQGroupId
'   MLM 11/07/02: SR 4798, CBB 2.2.18/3: Don't use the PSS to delete protocol in DeleteTrialSD.
'   ATO 23/08/2002 - Added RepeatNumber to routines DataItemCodeFromTaskId,DataItemIDFromTaskId,DataItemNameFromTaskId
'   ZA 24/09/2002 - Moved the delete statements within the Transaction in DeleteTrialSD routine
'   NCJ 25 Sept 02 - Fixed bug in QGroupExists
'   NCJ 15 Oct 02 - Add repeat number if > 1 in DataItemNameFromTaskId (and use square brackets for cycles elsewhere)
'   ASH 10/01/03 - Delete from LFMESSAGE,DATAIMPORT tables added to DeleteTrialPRD
'   Mo  4/2/2003    New functions added VisitIdFromCode, CRFPageIdFromCode, DataItemIdFromCode
'                   IdFromTrialSiteSubjectLabel, DataItemFormatFromId
'                   (for use of MACRO Query module and Batch Data Entry)
'   NCJ 6 Mar 03 - Remove registration data in DeleteTrialPRD (MACRO 3 TESTER Bug 511)
'   NCJ 30 Apr 03 - Added Timezone offset to gdsUpdateTrialStatus
'   REM 05/02/04 - Remove Batch data entry buffer entries when deleting patient data during a delete study
'   NCJ 19 Dec 05 - Added Partial Date Flag to DataItemFormatFromId
'   TA  18/01/2006 - MessageId now calculated by a sequence to avoid duplicate id problem
'   Mo 21/8/2006 - Bug 2784 new function gdsDataValuesALL added for the use of the Query Module.
'                   gdsDataValuesALL (Active and Inactive codes) copied from gdsDataValues (Active codes only).
'   ic 28/02/2007 issue 2855 clinical coding release - delete coding history data
'------------------------------------------------------------------------------------'

Option Explicit
Option Base 0
Option Compare Binary

'---------------------------------------------------------------------
Public Function gdsDataValues(ClinicalTrialId As Long, VersionId As Integer, DataItemId As Long) As ADODB.Recordset
'---------------------------------------------------------------------
'Creates a ReadOnly recordset, that supports RecordCount.
'---------------------------------------------------------------------
Dim sSQL As String
    
    On Error GoTo ErrHandler
    
    'SDM SR1303 26/10/99   Ensure only Active items are returned.
    'NCJ 7 Dec 99, SR 1492 - ORDER BY ValueOrder (not ValueId)
    sSQL = "SELECT ValueData.* FROM ValueData " _
        & " WHERE ClinicalTrialId = " & ClinicalTrialId _
        & " AND VersionId = " & VersionId _
        & " AND DataItemId = " & DataItemId _
        & " AND Active = 1" _
        & " ORDER BY ValueOrder"
    Set gdsDataValues = New ADODB.Recordset
    gdsDataValues.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsDataValues", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Public Function gdsDataValuesALL(ClinicalTrialId As Long, VersionId As Integer, DataItemId As Long) As ADODB.Recordset
'---------------------------------------------------------------------
'Creates a ReadOnly recordset, that supports RecordCount.
'Copied from gdsDataValues, but does not filter on Active = 1
'---------------------------------------------------------------------
Dim sSQL As String
    
    On Error GoTo ErrHandler
    
    sSQL = "SELECT ValueData.* FROM ValueData " _
        & " WHERE ClinicalTrialId = " & ClinicalTrialId _
        & " AND VersionId = " & VersionId _
        & " AND DataItemId = " & DataItemId _
        & " ORDER BY ValueOrder"
    Set gdsDataValuesALL = New ADODB.Recordset
    gdsDataValuesALL.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsDataValuesALL", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Public Function gnCurrentVersionId( _
    ByVal vClinicalTrialId As Long) As Integer
'---------------------------------------------------------------------

Dim rsTrial As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler
    
    sSQL = "SELECT max(VersionId) as CurrentVersionId FROM StudyDefinition " _
        & " WHERE ClinicalTrialId = " & vClinicalTrialId
    Set rsTrial = New ADODB.Recordset
    rsTrial.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
                        
    If IsNull(rsTrial!CurrentVersionId) Then
        gnCurrentVersionId = 0
    Else
        gnCurrentVersionId = rsTrial!CurrentVersionId
    End If
    
    rsTrial.Close
    Set rsTrial = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gnCurrentVersionId", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Public Function gdsStudyDefinition(ClinicalTrialId As Long, VersionId As Integer) As ADODB.Recordset
'---------------------------------------------------------------------
'Creates a Writable recordset.
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    sSQL = "SELECT * FROM StudyDefinition " _
        & " WHERE ClinicalTrialId = " & ClinicalTrialId _
        & " AND VersionId = " & VersionId
    Set gdsStudyDefinition = New ADODB.Recordset
    gdsStudyDefinition.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsStudyDefinition", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Public Sub gdsUpdateTrialStatus(ClinicalTrialId As Long, _
            VersionId As Integer, statusId As Integer)
'---------------------------------------------------------------------
'   Added by JL
'   Store timestamps as doubles - NCJ 30/11/99
'   NCJ 30 Apr 03 - Added timezone offset
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim nNextId As Integer
Dim oTimeZone As TimeZone

    On Error GoTo ErrHandler
    
    sSQL = "UPDATE ClinicalTrial SET " _
        & "  ClinicalTrial.StatusId= " & statusId _
        & " WHERE ClinicalTrialId = " & ClinicalTrialId
            
    MacroADODBConnection.Execute sSQL
    
    sSQL = "SELECT max(TrialStatusChangeId) as MaxTrialStatusChangeId FROM TrialStatusHistory " _
        & "WHERE ClinicalTrialId = " & ClinicalTrialId _
        & " AND VersionId = " & VersionId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
                                            
    If IsNull(rsTemp!MaxTrialStatusChangeId) Then
        nNextId = 1
    Else
        nNextId = rsTemp!MaxTrialStatusChangeId + 1
    End If
    
    Set oTimeZone = New TimeZone
    
    'WillC 4/2/00 changed the dates to SQLStandardNow to cope with regional settings
    'Mo Morris 30/8/01 Db Audit (UserId to UserName)
    ' NCJ 30 Apr 03 - Added StatusChangedTimestamp_TZ
    sSQL = "INSERT INTO TrialStatusHistory ( ClinicalTrialId, VersionId, " _
            & " TrialStatusChangeId, StatusId, UserName, " _
            & " StatusChangedTimestamp, StatusChangedTimestamp_TZ)" _
            & " VALUES (" & ClinicalTrialId & "," & VersionId _
            & "," & nNextId & "," & statusId & ",'" & goUser.UserName & "', " _
            & SQLStandardNow & ", " & oTimeZone.TimezoneOffset & ")"
            
    MacroADODBConnection.Execute sSQL
    
    rsTemp.Close
    Set rsTemp = Nothing

    Set oTimeZone = Nothing
    
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsUpdateTrialStatus", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'---------------------------------------------------------------------
Public Function GetCRFPageDateLabel(ByVal vClinicalTrialId As Long, _
                           ByVal vVersionId As Integer, _
                           ByVal vCRFPageId As Long) As String
'---------------------------------------------------------------------
' Read the CRFPageDateLabel from the CRFPage table
' Returns "" if no date label defined
' NCJ 30/11/99 - Based on SDM SR687, 04/11/99 (in frmStudyVisits)
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsDate As ADODB.Recordset 'SDM SR687 04/11/99

    sSQL = "SELECT CRFPageDateLabel FROM CRFPage " & _
            "WHERE ClinicalTrialId = " & vClinicalTrialId & _
            "AND VersionId = " & vVersionId & _
            "AND CRFPageId = " & vCRFPageId
            
    Set rsDate = New ADODB.Recordset
    rsDate.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ' NCJ 28/2/00 Screen out NULL values
    If Not rsDate.EOF Then
        GetCRFPageDateLabel = RemoveNull(rsDate!CRFPageDateLabel)
    End If
    
    rsDate.Close
    Set rsDate = Nothing

End Function

'---------------------------------------------------------------------
Public Function gdsCRFPage(ByVal vClinicalTrialId As Long, _
                           ByVal vVersionId As Integer, _
                           ByVal vCRFPageId As Long) As ADODB.Recordset
'---------------------------------------------------------------------
'Creates a ReadOnly recordset.
'Retrieve details of a CRF page.
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    sSQL = "SELECT * FROM CRFPage " _
        & "  WHERE   ClinicalTrialId             =  " & vClinicalTrialId _
        & "  AND     VersionId                   =  " & vVersionId _
        & "  AND     CRFPageId                   =  " & vCRFPageId
    
    Set gdsCRFPage = New ADODB.Recordset
    'REM 13/05/02 - open a readonly recordset that allows recordcount
    gdsCRFPage.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    'gdsCRFPage.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsCRFPage", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Public Function DataItemExists(ClinicalTrialId As Long, VersionId As Integer, DataItemCode As String) As Long
'---------------------------------------------------------------------
'TA 29/03/2000 Now Returns ID or -1 if not found
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler
    
    sSQL = "SELECT DataItemId FROM DataItem " _
        & " WHERE DataItemCode = '" & DataItemCode & "'" _
        & " AND DataItem.ClinicalTrialId = " & ClinicalTrialId _
        & " AND DataItem.VersionId = " & VersionId
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        DataItemExists = -1
    Else
        DataItemExists = rsTemp!DataItemId
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "DataItemExists", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Public Function QGroupExists(lClinicalTrialId As Long, nVersionId As Integer, sQGroupCode As String) As Long
'---------------------------------------------------------------------
' REM 20/02/02
' Check to see if a Question Group Exists, returns ID or -1 if not found
' NCJ 25 Sept 02 - Used GetSQLStringEquals to cope with upper/lower case comparisons
'---------------------------------------------------------------------
Dim rsQGroup As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler
    
    ' NCJ 25 Sept 02 - Removed incorrect references to DataItem table
    sSQL = "SELECT QGroupId FROM QGroup " _
        & " WHERE " & GetSQLStringEquals("QGroupCode", sQGroupCode) _
        & " AND ClinicalTrialId = " & lClinicalTrialId _
        & " AND VersionId = " & nVersionId
    
    Set rsQGroup = New ADODB.Recordset
    rsQGroup.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsQGroup.RecordCount = 0 Then
        QGroupExists = -1
    Else
        QGroupExists = rsQGroup!QGroupID
    End If
    
    rsQGroup.Close
    Set rsQGroup = Nothing

Exit Function

ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "QGroupExists", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function

'---------------------------------------------------------------------
Public Function mnNextCRFElementId(ByVal vClinicalTrialId As Long, _
                                    ByVal vVersionId As Integer, _
                                    ByVal vCRFPageId As Long) As Integer
'---------------------------------------------------------------------
'   Returns the next available unique id for a new CRF element on a CRF page
'---------------------------------------------------------------------

Dim rsTmp As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler
    
    'changed by Mo Morris 20/12/99, adOpenForwardOnly changed to adOpenStatic
    sSQL = "SELECT MAX(CRFElementId) as MaxCRFElementId FROM CRFElement " _
        & "  WHERE  ClinicalTrialId          = " & vClinicalTrialId _
        & "  AND    VersionId                = " & vVersionId _
        & "  AND    CRFPageId                = " & vCRFPageId
    Set rsTmp = New ADODB.Recordset
    rsTmp.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
    
    If IsNull(rsTmp!MaxCRFElementId) Then     'if no CRF elements on this page
        mnNextCRFElementId = gnFIRST_ID
    Else
        mnNextCRFElementId = rsTmp!MaxCRFElementId + gnID_INCREMENT
    End If
    
    rsTmp.Close
    Set rsTmp = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "mnNextCRFElementId", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Public Function gblnDataItemExportExists(iClinicalTrialId As Long, _
                                    iVersionId As Integer, sExportName As String, _
                                    sCode As String) As Boolean
'---------------------------------------------------------------------
'   PN change 18 new function
'---------------------------------------------------------------------

Dim rsTemp As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler
    
    sSQL = "SELECT 'x' FROM DataItem " _
        & " WHERE DataItem.ExportName = '" & sExportName & "'" _
        & " AND DataItem.ClinicalTrialId = " & iClinicalTrialId _
        & " AND DataItem.VersionId = " & iVersionId _
        & " AND DataItem.DataItemCode <> '" & sCode & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        gblnDataItemExportExists = False
    Else
        gblnDataItemExportExists = True
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gblnDataItemExportExists", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Public Sub DeleteTrialPRD(lTrialIdForDeletion As Long, sTrialName As String)
'---------------------------------------------------------------------
'   Delete all Patient Response Data on the specified Study Definition
'   NCJ 19/10/00 - Removed NormalRange (this was old!)
'   ASH 10/01/03 - Delete from LFMESSAGE,DATAIMPORT tables added
'   NCJ 6 Mar 03 - Delete registration data from RSSubjectIdentifier, RSNextNumber and RSUniquenessCheck
'   REM 05/02/04 - Remove Batch data entry buffer entries
'   ic 28/02/2007 issue 2855 clinical coding release - delete coding history data
'---------------------------------------------------------------------

Dim sSQL As String
    
    On Error GoTo ErrHandler
    ' Delete records from all tables containing Patient Response Data
    'Begin transaction
    TransBegin
    
    sSQL = "DELETE FROM DataItemResponseHistory " _
        & "WHERE ClinicalTrialId = " & lTrialIdForDeletion
    MacroADODBConnection.Execute sSQL
    
    sSQL = "DELETE FROM DataItemResponse " _
        & "WHERE ClinicalTrialId = " & lTrialIdForDeletion
    MacroADODBConnection.Execute sSQL
    
    sSQL = "DELETE FROM CRFPageInstance " _
        & "WHERE ClinicalTrialId = " & lTrialIdForDeletion
    MacroADODBConnection.Execute sSQL
            
    sSQL = "DELETE FROM  TrialSubject " _
        & "WHERE ClinicalTrialId = " & lTrialIdForDeletion
    MacroADODBConnection.Execute sSQL
       
    sSQL = "DELETE FROM VisitInstance " _
        & "WHERE ClinicalTrialId = " & lTrialIdForDeletion
    MacroADODBConnection.Execute sSQL
          
    sSQL = "DELETE FROM TrialSite " _
        & "WHERE ClinicalTrialId = " & lTrialIdForDeletion
    MacroADODBConnection.Execute sSQL
            
    sSQL = "DELETE FROM Message " _
        & "WHERE ClinicalTrialId = " & lTrialIdForDeletion
    MacroADODBConnection.Execute sSQL
    
    ' NCJ 19/10/00 - Delete from MIMessage table too
    sSQL = "DELETE FROM MIMessage " _
        & "WHERE MIMessageTrialName = '" & sTrialName & "'"
    MacroADODBConnection.Execute sSQL
    
    'ASH 10/01/03 - Delete from DATAIMPORT table
    sSQL = "DELETE FROM DATAIMPORT " _
        & "WHERE CLINICALTRIALNAME = '" & sTrialName & "'"
    MacroADODBConnection.Execute sSQL
    
    'ASH 10/01/03 - Delete from LFMESSAGE table
    sSQL = "DELETE FROM LFMESSAGE " _
        & "WHERE CLINICALTRIALID = " & lTrialIdForDeletion
    MacroADODBConnection.Execute sSQL

    ' NCJ 6 Mar 03 - Remove registration data too (keyed on Trial NAME not ID)
    sSQL = "DELETE FROM RSSUBJECTIDENTIFIER " _
        & "WHERE CLINICALTRIALNAME = '" & sTrialName & "'"
    MacroADODBConnection.Execute sSQL
    
    sSQL = "DELETE FROM RSNEXTNUMBER " _
        & "WHERE CLINICALTRIALNAME = '" & sTrialName & "'"
    MacroADODBConnection.Execute sSQL
    
    sSQL = "DELETE FROM RSUNIQUENESSCHECK " _
        & "WHERE CLINICALTRIALNAME = '" & sTrialName & "'"
    MacroADODBConnection.Execute sSQL
    
    'REM 05/02/04 - Remove Batch data entry buffer entries
    sSQL = "DELETE FROM BATCHRESPONSEDATA " _
        & "WHERE CLINICALTRIALID = " & lTrialIdForDeletion
    MacroADODBConnection.Execute sSQL
    
    'ic 28/02/2007 issue 2855 clinical coding release - delete coding history data
    If (gbClinicalCoding) Then
        sSQL = "DELETE FROM CODINGHISTORY " _
            & "WHERE CLINICALTRIALID = " & lTrialIdForDeletion
        MacroADODBConnection.Execute sSQL
    End If
    
    'End transaction
    TransCommit
    
    'Log the deletion
    'Mo Morris 4/10/00, message text changed from Clinical Trial to Study
    gLog gsDEL_TRIAL_PRD, "Study [" & sTrialName & "] Deleted."

Exit Sub
ErrHandler:
    'RollBack transaction
    TransRollBack

    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "DeleteTrialPRD", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'---------------------------------------------------------------------
Public Sub DeleteTrialSD(lTrialIdForDeletion As Long, sTrialName As String)
'---------------------------------------------------------------------
'   Delete all references to the specified Study Definition
'   Delete records from all tables containing Study Definition data
'---------------------------------------------------------------------
'NCJ 17/9/99        Also delete Arezzo protocol here
'Mo Morris 20/9/99  old sub gDeleteTrial split into DeleteTrialSD and DeleteTrialPRD
'Mo Morris 4/10/99  DAO to ADO conversion
'Mo Morris 21/1/00  The deletion of StudyReport and StudyReportData added
'Mo Morris 12/12/00 Re-wriiten. based on querying  MacroTable for StudyDefinition
'                   tables with a SegmentId less than 300 (i.e tables that have a
'                   key of ClinicalTrialId)
'Mo Morris 8/5/2002 Code changed back to how it used to be.
'                   Protocols Table entry no longer removed by code here, but is
'                   removed by the PSS.DLL which is called by DeleteProformaTrial
' MLM 11/07/02: SR 4798, CBB 2.2.18/3: Delete from Protocols table directly rather than using PSS,
'               so that the protocol is reinstated when rolling back MacroADODBConnection.
'               NOTE: Callers must refresh the PSS's Protocols collection themselves!
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTables As ADODB.Recordset

     On Error GoTo ErrHandler
    
    ' Delete records from all tables containing Study Definition data

    'Begin transaction
    TransBegin
    
    sSQL = "Select TableName, SegmentId From MACROTable" _
        & " WHERE STYDEF = 1" _
        & " ORDER BY SegmentId"
    Set rsTables = New ADODB.Recordset
    'TA 21/02/2001: Changed to static cursor - see Knowledge Base Q272358
    rsTables.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
    
    'Loop through rstables deleting the required table entries for the specified ClinicalTrialId
    Do Until rsTables.EOF
        'Note that rsTables is sorted on SegmentId and once segmentId goes above 300 the loop can be exited
        If rsTables!SegmentId = "300" Then
            Exit Do
        End If
        sSQL = "DELETE FROM " & rsTables!TableName _
            & " WHERE ClinicalTrialId = " & lTrialIdForDeletion
        MacroADODBConnection.Execute sSQL
        rsTables.MoveNext
    Loop
    
    rsTables.Close
    Set rsTables = Nothing
    
    'ZA 24/09/2002 - moved the following statements within the Transaction
    ' NCJ 25 Sept 02 - Use GetSQLStringEquals (copes with upper/lower case)
    sSQL = "DELETE FROM Protocols WHERE " & GetSQLStringEquals("FileName", sTrialName)
    MacroADODBConnection.Execute sSQL
    
    'End transaction
    TransCommit
    
    
    'Log the deletion
    'Mo Morris 4/10/00, message text changed from Clinical Trial to Study
    gLog gsDEL_TRIAL_SD, "Study [" & sTrialName & "] Deleted."

Exit Sub
ErrHandler:
    'RollBack transaction
    TransRollBack
    
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "DeleteTrialSD", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'---------------------------------------------------------------------
Public Function TrialNameFromId(ByVal lClinicalTrialId As Long) As String
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler

    sSQL = "SELECT ClinicalTrialName FROM ClinicalTrial " _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        TrialNameFromId = ""
    Else
        TrialNameFromId = rsTemp!ClinicalTrialName
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "TrialNameFromId", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Private Function VisitFromId(ByVal bName As Boolean, _
                                ByVal lClinicalTrialId As Long, _
                                ByVal lVisitId As Long, _
                                Optional ByVal nVisitCycle As Integer = 1) As String
'---------------------------------------------------------------------
' Get Visit Name or Code from Visit Id
' If bName = TRUE, return visit name and cycle number
' If bName = FALSE, return Visit Code
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim sTemp As String

    On Error GoTo ErrHandler

    sTemp = ""
    sSQL = "SELECT VisitCode, VisitName FROM StudyVisit " _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND VisitId = " & lVisitId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        sTemp = ""
    ElseIf bName Then
        ' Visit name
        sTemp = RemoveNull(rsTemp!VisitName)
        If nVisitCycle > 1 Then
            sTemp = sTemp & "[" & nVisitCycle & "]"
        End If
    Else
        ' Visit Code
        sTemp = RemoveNull(rsTemp!VisitCode)
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

    VisitFromId = sTemp
    
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "VisitFromId", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function

'---------------------------------------------------------------------
Public Function VisitNameFromId(ByVal lClinicalTrialId As Long, _
                                ByVal lVisitId As Long, _
                                ByVal nVisitCycle As Integer) As String
'---------------------------------------------------------------------
                           
    ' True = Visit Name
    VisitNameFromId = VisitFromId(True, _
                            lClinicalTrialId, lVisitId, nVisitCycle)
    
End Function

'---------------------------------------------------------------------
Public Function VisitCodeFromId(ByVal lClinicalTrialId As Long, _
                                ByVal lVisitId As Long) As String
'---------------------------------------------------------------------
    
    ' False = Visit Code
    VisitCodeFromId = VisitFromId(False, _
                            lClinicalTrialId, lVisitId)

End Function

'---------------------------------------------------------------------
Private Function CRFPageFromTaskId(bName As Boolean, _
                                    ByVal lClinicalTrialId As Long, _
                                    ByVal sTrialSite As String, _
                                    ByVal lPersonId As Long, _
                                    ByVal lCRFPageTaskId As Long) As String
'---------------------------------------------------------------------
' NCJ 23/5/00
' Get CRFPage Code or Name
' If bName = TRUE return Name and cycle number
' If bName = FALSE return CRF Page code only
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim sTemp As String
    
    On Error GoTo ErrHandler
    
    sTemp = ""
    
    ' Get CRF Page Name and Code
    sSQL = "SELECT CRFTitle, CRFPageCode, CRFPageCycleNumber " _
    & " FROM CRFPage, CRFPageInstance " _
    & " WHERE CRFPage.ClinicalTrialId = CRFPageinstance.ClinicalTrialId" _
    & " AND CRFPage.ClinicalTrialId = " & lClinicalTrialId _
    & " AND CRFPage.CRFPageId = CRFPageinstance.CRFPageId " _
    & " AND CRFPageInstance.CRFPageTaskId = " & lCRFPageTaskId _
    & " AND CRFPageInstance.PersonId = " & lPersonId _
    & " AND CRFPageInstance.TrialSite = '" & sTrialSite & "'"
        
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        sTemp = ""
    ElseIf bName Then
        ' The CRF Page Name
        sTemp = rsTemp!CRFTitle
        If Val(RemoveNull(rsTemp!CRFPageCycleNumber)) > 1 Then
            sTemp = sTemp & "[" & rsTemp!CRFPageCycleNumber & "]"
        End If
    Else
        ' The CRF Page Code
        sTemp = RemoveNull(rsTemp!CRFPageCode)
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

    CRFPageFromTaskId = sTemp
    
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "CRFPageFromTaskId", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function

'---------------------------------------------------------------------
Public Function CRFPageNameFromTaskId(ByVal lClinicalTrialId As Long, _
                                    ByVal sTrialSite As String, _
                                    ByVal lPersonId As Long, _
                                    ByVal lCRFPageTaskId As Long) As String
'---------------------------------------------------------------------
' Get CRF Page Name and cycle number from CRFPageTaskID
'---------------------------------------------------------------------
    
    ' True = CRF Page Name
    CRFPageNameFromTaskId = CRFPageFromTaskId(True, lClinicalTrialId, sTrialSite, _
                                lPersonId, lCRFPageTaskId)
                                
End Function

'---------------------------------------------------------------------
Public Function CRFPageCodeFromTaskId(ByVal lClinicalTrialId As Long, _
                                    ByVal sTrialSite As String, _
                                    ByVal lPersonId As Long, _
                                    ByVal lCRFPageTaskId As Long) As String
'---------------------------------------------------------------------

    ' False = CRF Page Code
    CRFPageCodeFromTaskId = CRFPageFromTaskId(False, lClinicalTrialId, sTrialSite, _
                                lPersonId, lCRFPageTaskId)

End Function

'---------------------------------------------------------------------
Private Function DataItemIdFromTaskId(ByVal lClinicalTrialId As Long, _
                                    ByVal sTrialSite As String, _
                                    ByVal lPersonId As Long, _
                                    ByVal lResponseTaskId As Long, _
                                    ByVal nRepeatNumber As Integer) As Long
'---------------------------------------------------------------------
' NCJ 19/5/00 - Copied from DataItemCodeFromTaskId
' Get data item Id from ResponseTaskId
' Returns -1 if data item not found
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    'Get the DataItemId
    sSQL = "SELECT DataItemId FROM DataItemResponse " _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND TrialSite = '" & sTrialSite & "'" _
        & " AND PersonId = " & lPersonId _
        & " AND ResponseTaskId = " & lResponseTaskId _
        & " AND RepeatNumber = " & nRepeatNumber
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount > 0 Then
        DataItemIdFromTaskId = rsTemp!DataItemId
    Else
        DataItemIdFromTaskId = glMINUS_ONE
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

End Function

'---------------------------------------------------------------------
Public Function DataItemCodeFromTaskId(ByVal lClinicalTrialId As Long, _
                                    ByVal sTrialSite As String, _
                                    ByVal lPersonId As Long, _
                                    ByVal lResponseTaskId As Long, _
                                    ByVal nRepeatNumber As Integer) As String
'---------------------------------------------------------------------
' Get data item code from ResponseTaskId
'TA 22/09/2000 : ClinicalTrialId also needed
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim lDataItemId As Long

    On Error GoTo ErrHandler
    
    ' NCJ 19/5/00 - Code moved to DataItemIdFromTaskId
    'Get the DataItemId first
    lDataItemId = DataItemIdFromTaskId(lClinicalTrialId, sTrialSite, lPersonId, lResponseTaskId, nRepeatNumber)
    
    'Now get the DataItemCode
    sSQL = "SELECT DataItemCode FROM DataItem " _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND DataItemId = " & lDataItemId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        DataItemCodeFromTaskId = ""
    Else
        DataItemCodeFromTaskId = rsTemp!DataItemCode
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "DataItemCodeFromTaskId", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Public Function DataItemNameFromTaskId(ByVal lClinicalTrialId As Long, _
                                    ByVal sTrialSite As String, _
                                    ByVal lPersonId As Long, _
                                    ByVal lResponseTaskId As Long, _
                                    ByVal nRepeatNumber As Integer) As String
'---------------------------------------------------------------------
' NCJ 19/5/00 - Based on Mo's DataItemCodeFromTaskId
' Get data item name from ResponseTaskId
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim lDataItemId As Long

    On Error GoTo ErrHandler
    
    'Get the DataItemId first
    lDataItemId = DataItemIdFromTaskId(lClinicalTrialId, sTrialSite, lPersonId, lResponseTaskId, nRepeatNumber)
    
    'Now get the DataItemName
    sSQL = "SELECT DataItemName FROM DataItem " _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND DataItemId = " & lDataItemId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount > 0 Then
        DataItemNameFromTaskId = rsTemp!DataItemName
        ' NCJ 15 Oct 02 - Add repeat number if > 1
        If nRepeatNumber > 1 Then
            DataItemNameFromTaskId = DataItemNameFromTaskId & "[" & nRepeatNumber & "]"
        End If
    Else
        DataItemNameFromTaskId = ""
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                    "DataItemNameFromTaskId", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Public Function TrialIdFromName(ByVal sClinicalTrialName As String) As Long
'---------------------------------------------------------------------
' NCJ 6/11/00 - Returns -1 if sClinicaltrialName doesn't exist
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler

    sSQL = "SELECT ClinicalTrialId FROM ClinicalTrial " _
        & " WHERE ClinicalTrialName = '" & sClinicalTrialName & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        TrialIdFromName = -1
    Else
        TrialIdFromName = rsTemp!ClinicalTrialId
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "TrialIdFromName", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function


'---------------------------------------------------------------------
Public Function SubjectLabelFromTrialSiteId(lTrialId As Long, sSite As String, _
                                            lPersonId As Long) As String
'---------------------------------------------------------------------
' NCJ 7 Aug 2000 Retrieve Subject label given TrialId, Site and PersonID
' Returns empty string if no label exists
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler

    sSQL = "SELECT LocalIdentifier1 FROM TrialSubject " _
        & " WHERE ClinicalTrialId = " & lTrialId _
        & " AND TrialSite = '" & sSite & "'" _
        & " AND PersonId = " & lPersonId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        SubjectLabelFromTrialSiteId = ""
    Else
        SubjectLabelFromTrialSiteId = RemoveNull(rsTemp!LocalIdentifier1)
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                            "SubjectLabelFromTrialSiteId", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function

'---------------------------------------------------------------------
Public Function DataItemCodeFromId(ByVal lClinicalTrialId As Long, _
                                ByVal lDataItemId As Long) As String
'---------------------------------------------------------------------
    
    ' False = DataItem Code
    DataItemCodeFromId = DataItemFromId(False, _
                            lClinicalTrialId, lDataItemId)

End Function

Public Function DataItemNameFromId(ByVal lClinicalTrialId As Long, _
                                ByVal lDataItemId As Long) As String
'---------------------------------------------------------------------
                           
    ' True = DataItem Name
    DataItemNameFromId = DataItemFromId(True, _
                            lClinicalTrialId, lDataItemId)
    
End Function

'---------------------------------------------------------------------
Public Function DataItemFromId(ByVal bName As Boolean, _
                                ByVal lClinicalTrialId As Long, _
                                ByVal lDataItemId As Long) As String
'---------------------------------------------------------------------
' Get DataItem Name or Code from DataItem Id
' If bName = TRUE, return DataItem name
' If bName = FALSE, return DataItem Code
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim sTemp As String

    On Error GoTo ErrHandler

    sTemp = ""
    ' NCJ 2/1/01 - Changed DataItemIdId to DataItemId
    sSQL = "SELECT DataItemCode, DataItemName FROM DataItem " _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND DataItemId = " & lDataItemId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        sTemp = ""
    ElseIf bName Then
        ' DataItem name
        sTemp = RemoveNull(rsTemp!DataItemName)
    Else
        ' DataItem Code
        sTemp = RemoveNull(rsTemp!DataItemCode)
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

    DataItemFromId = sTemp
    
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "DataItemFromId", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function

'---------------------------------------------------------------------
Public Function QGroupFromID(ByVal lClinicalTrialId As Long, ByVal nVersionId As Integer, _
                                 ByVal lQGroupId As Long) As ADODB.Recordset
'---------------------------------------------------------------------
' REM 20/02/02
' Returns a reordset containing the QGroupCode, QGroupName and DisplayType
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsQG As ADODB.Recordset
    
    sSQL = "SELECT QGroupCode, QGroupName, DisplayType FROM QGroup " _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND VersionId = " & nVersionId _
        & " AND QGroupId = " & lQGroupId
    Set rsQG = New ADODB.Recordset
    rsQG.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    Set QGroupFromID = rsQG
    
    Set rsQG = Nothing
    
Exit Function

ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "QGroupFromID", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function

'---------------------------------------------------------------------
Public Function NewQGroupId(ByVal lToClinicalTrialId As Long, ByVal nToVersionId As Integer) As Long
'---------------------------------------------------------------------
' REM 22/02/02
' Returns a new QGroup Id
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsQGroupID As ADODB.Recordset

    'SQL to find max QGroupID in the study being copied to
    sSQL = "SELECT MAX (QGroupID) as MaxQGroupID FROM QGroup" & _
           " WHERE ClinicalTrialID = " & lToClinicalTrialId & _
           " AND VersionID = " & nToVersionId
    
    Set rsQGroupID = New ADODB.Recordset
    rsQGroupID.Open sSQL, MacroADODBConnection
    
    'Get new QGroup ID by check max current ID and making it 1 more
    NewQGroupId = rsQGroupID.Fields!MaxQGroupID + gnID_INCREMENT

End Function

'---------------------------------------------------------------------
Public Function TrialDescriptionFromName(ByVal sClinicalTrialName As String) As String
'---------------------------------------------------------------------
'Added by Mo Morris 10/10/2000
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim sTemp As String

    On Error GoTo ErrHandler

    sSQL = "SELECT ClinicalTrialDescription FROM ClinicalTrial " _
        & " WHERE ClinicalTrialName = '" & sClinicalTrialName & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        sTemp = ""
    Else
        sTemp = rsTemp!ClinicalTrialDescription
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    TrialDescriptionFromName = sTemp

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "TrialDescriptionFromName", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Public Function SiteDescriptionFromSite(ByVal sSite As String) As String
'---------------------------------------------------------------------
'Added by Mo Morris 10/10/2000
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim sTemp As String

    On Error GoTo ErrHandler

    sSQL = "SELECT SiteDescription FROM Site " _
        & " WHERE Site = '" & sSite & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        sTemp = ""
    Else
        sTemp = rsTemp!SiteDescription
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    SiteDescriptionFromSite = sTemp

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "SiteDescriptionFromSite", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Public Sub InsertMessage(sTrialSite As String, lClinicalTrialId As Long, nMessageType As Integer, _
                         sMessageTimeStamp As String, nMessageTimeStamp_TZ As Integer, sUsername As String, sMessageBody As String, _
                         sMessageParameters As String, nMessageDirection As Integer, nMessageReceived As Integer)
'---------------------------------------------------------------------
'REM 01/09/03
'Insert a message into the Message table
'---------------------------------------------------------------------
Dim lMessageId As Long
Dim sSQL As String
Dim oTimeZone As TimeZone
Dim nTimeZoneOffSet As Integer
Dim sMessageReceivedTimeStamp As String

    On Error GoTo ErrLabel

    Set oTimeZone = New TimeZone
    
    'get the timestamp and the time-zone offset for local machine
    sMessageReceivedTimeStamp = SQLStandardNow
    nTimeZoneOffSet = oTimeZone.TimezoneOffset
    
    'get next message id
    lMessageId = NextMessageId
    
    'Insert the received message into MACRO DB Message table
    sSQL = "INSERT INTO Message (TrialSite, ClinicalTrialId, MessageType, MessageTimeStamp, UserName, MessageBody," _
        & " MessageParameters, MessageReceived, MessageDirection, MessageId, MessageReceivedTimeStamp, MessageTimeStamp_TZ, MessageReceivedTimeStamp_TZ)" _
        & "  VALUES ('" & sTrialSite & "'," & lClinicalTrialId & "," & nMessageType & "," & sMessageTimeStamp & ",'" & sUsername & "','" & sMessageBody & "','" _
        & sMessageParameters & "'," & nMessageReceived & "," & nMessageDirection & "," & lMessageId & "," _
        & sMessageReceivedTimeStamp & "," & nMessageTimeStamp_TZ & "," & nTimeZoneOffSet & ")"
    MacroADODBConnection.Execute sSQL, , adCmdText
    
    Set oTimeZone = Nothing
    
Exit Sub
ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "basCommon.InsertMessage"
End Sub

'-----------------------------------------------------
Public Function NextMessageId(Optional sConMACRO As ADODB.Connection = Nothing) As Long
'-----------------------------------------------------
'Return the next MessageId for table Message
'   TA  18/01/2006 - MessageId now calculated by a sequence to avoid duplicate id problem
'-----------------------------------------------------
Dim sNewConMACRO As ADODB.Connection
Dim xfer As SysDataXfer

    On Error GoTo ErrHandler
    
    If Not sConMACRO Is Nothing Then
        Set sNewConMACRO = sConMACRO
    Else
        Set sNewConMACRO = MacroADODBConnection
    End If
        
    Set xfer = New SysDataXfer
    NextMessageId = xfer.GetNextMessageId(sNewConMACRO)
    Set xfer = Nothing
            
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "basTrialData.NextMessageId"
        
End Function

'---------------------------------------------------------------------
Public Function CRFPageCodeFromId(ByVal lClinicalTrialId As Long, _
                                    ByVal lCRFPageId As Long) As String
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim sTemp As String

    On Error GoTo ErrHandler

    sTemp = ""
    sSQL = "SELECT CRFPageCode FROM CRFPage " _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND CRFPageId = " & lCRFPageId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        sTemp = ""
    Else
        sTemp = RemoveNull(rsTemp!CRFPageCode)
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

    CRFPageCodeFromId = sTemp

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "CRFPageCodeFromId", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function

'---------------------------------------------------------------------
Public Function DataTypeFromId(ByVal lClinicalTrialId As Long, _
                                    ByVal lDataItemId As Long) As Long
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim nTempId As Long

    On Error GoTo ErrHandler

    sSQL = "SELECT DataType FROM DataItem " _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND DataItemId = " & lDataItemId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        nTempId = 0
    Else
        nTempId = RemoveNull(rsTemp!DataType)
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

    DataTypeFromId = nTempId

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "DataTypeFromId", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function

'---------------------------------------------------------------------
Public Function IdFromTrialSiteSubjectLabel(ByVal lClinicalTrialId As Long, _
                                            ByVal sSite As String, _
                                            ByVal sSubjectLabel As String) As Long
'---------------------------------------------------------------------
'Note that MACRO has nothing in place to guarantee unique SubjectLabels.
'This function locates SubjectLabels in table TrialSubject using the calling
'Trial/Site/SubjectLabel parameters.
'If no matches are found a PersonId of 0 is returned.
'If a single match is found its PersonId is returned.
'If more than one subject is found a PersonId of -1 is returned
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim lPersonId As Long

    On Error GoTo ErrHandler
    
    sSQL = "SELECT PersonId FROM TrialSubject " _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND TrialSite = '" & sSite & "'" _
        & " AND LocalIdentifier1 = '" & sSubjectLabel & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        lPersonId = 0
    ElseIf rsTemp.RecordCount = 1 Then
        lPersonId = RemoveNull(rsTemp!PersonId)
    Else
        lPersonId = -1
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    IdFromTrialSiteSubjectLabel = lPersonId

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "IdFromTrialSiteSubjectLabel", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function

'---------------------------------------------------------------------
Public Function VisitIdFromCode(ByVal lClinicalTrialId, ByVal sVisitCode As String) As Long
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler

    sSQL = "SELECT VisitId FROM StudyVisit " _
        & " WHERE ClinicalTrialid = " & lClinicalTrialId _
        & " AND VisitCode = '" & sVisitCode & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        VisitIdFromCode = -1
    Else
        VisitIdFromCode = rsTemp!VisitId
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "VisitIdFromCode", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Public Function CRFPageIdFromCode(ByVal lClinicalTrialId, ByVal sCRFPageCode As String) As Long
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler

    sSQL = "SELECT CRFPageId FROM CRFPage " _
        & " WHERE ClinicalTrialid = " & lClinicalTrialId _
        & " AND CRFPageCode = '" & sCRFPageCode & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        CRFPageIdFromCode = -1
    Else
        CRFPageIdFromCode = rsTemp!CRFPageId
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CRFPageIdFromCode", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Public Function DataItemIdFromCode(ByVal lClinicalTrialId, ByVal sDataItemCode As String) As Long
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler

    sSQL = "SELECT DataItemId FROM DataItem " _
        & " WHERE ClinicalTrialid = " & lClinicalTrialId _
        & " AND DataItemCode = '" & sDataItemCode & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        DataItemIdFromCode = -1
    Else
        DataItemIdFromCode = rsTemp!DataItemId
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "DataItemIdFromCode", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Public Function DataItemFormatFromId(ByVal lClinicalTrialId As Long, _
                                ByVal lDataItemId As Long, _
                                Optional ByRef nPDFlag As Integer) As String
'---------------------------------------------------------------------
' NCJ 19 Dec 05 - Added nPDFlag argument (it's 1 for a Partial Date)
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler

    sSQL = "SELECT DataItemFormat, DataItemCase FROM DataItem " _
        & " WHERE ClinicalTrialid = " & lClinicalTrialId _
        & " AND DataItemId = " & lDataItemId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        DataItemFormatFromId = ""
        nPDFlag = 0
    Else
        DataItemFormatFromId = rsTemp!DataItemFormat
        nPDFlag = CInt(RemoveNull(rsTemp!DataItemCase))
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "DataItemFormatFromId", "basTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function
