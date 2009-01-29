Attribute VB_Name = "basCRFData"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       basCRFSQL.bas
'   Author:     Andrew Newbigging, July 1998
'   Purpose:    SQL routines for Case Report form.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'  1-20        Andrew Newbigging   4/06/97 - 11/11/98
'       Following routines moved form this module to TrialData module:
'           CopyCRFPage
'           gblnCRFPageExists
'           gdsCRFPage
'           gdsCRFPageElementList,mnCopyCRFElement,DeleteCRFPage,mnNextCRFElementId
'   21      Paul Norris          23/-8/99   SR 1712
'           Amended gdsCRFPageList() to include bFormAlphabeticOrder parameter
'   22      NCJ 1/9/99
'           Removed Validation from CRFElement recordset
'   23      Paul Norris          06/09/99   SR1499 saving of new RoleCode field in CRFElement table
'   24      NCJ 8 Sep 99
'           Removed calls to InsertProformaCRFElement
'   PN  24/09/99    Moved gDeleteCRFElement(), gdsCRFPageList(), gnInsertCRFElement(),
'                   mnNextCRFPageOrder(), mnNextFieldOrder(), UpdateCRFElementProperties()
'                   UpdateCRFPageTitle(), UpdateCRFPageDetails(), UpdateCRFFieldOrder(),
'                   UpdateCRFPageDisplayNumbers(), UpdateCRFPageLabel() to modSDCRFData
'   Mo Morris   9/11/99
'       DAO to ADO conversion
'   NCJ 16 Nov 99   Return DataItemCase in CRFElement sets
'  WillC    Changed the Following where present from Integer to Long  ClinicalTrialId
'           CRFPageId,VisitId,CRFElementID
'   NCJ 15/9/00 - Added LabTest things to gdsCRFPageDataItems
'   TA  19/9/00 - Added LabTest things to gdsCRFDataElements
'   NCJ 5/10/00 - ClinicalTestId -> ClinicalTestCode
'   NCJ 31/10/00 - Adjusted gdsCRFPageDataItemsYorder to cope correctly with lines
'                   (the old DemoStudy wasn't printing properly)
' MACRO 3.0
'   NCJ 13 Dec 01 - Added QGroupFieldOrder to gdsCRFPageDataItems
'------------------------------------------------------------------------------------'

Option Explicit


'---------------------------------------------------------------------
Public Function gdsCRFPageDataItems(ByVal vClinicalTrialId As Long, _
                                    ByVal vVersionId As Integer, _
                                    ByVal vCRFPageId As Long) As ADODB.Recordset
'---------------------------------------------------------------------
'Retrieve data items on a CRF page, in field order
'Creates a ReadOnly recordset, that supports RecordCount.
' NCJ 15/9/00 - Include ClinicalTestCode and TestDateExpr
' NCJ 13/12/01 - Order by QGroupFieldOrder too
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
'   ATN 21/12/99
'   Oracle doesn't like LEFT JOIN.
' NCJ 15/9/00 - Include ClinicalTestCode
' NCJ 13/12/01 - Include QGroupFieldOrder as sub-ordering
    If goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
        sSQL = "SELECT CRFElement.*, DataItemCode, DataItemName,DataType, DataItemLength, " _
            & "DataItemFormat, UnitOfMeasurement, Derivation, DataItemHelpText, " _
            & " DataItemCase, ClinicalTestCode " _
            & "FROM   CRFElement , DataItem " _
            & "WHERE     DataItem.ClinicalTrialId (+) = CRFElement.ClinicalTrialId " _
            & "AND    DataItem.VersionId (+) = CRFElement.VersionId " _
            & "AND    DataItem.DataItemId (+) = CRFElement.DataItemId " _
            & "AND  CRFElement.ClinicalTrialId = " & vClinicalTrialId _
            & " AND    CRFElement.VersionId = " & vVersionId _
            & " AND    CRFElement.CRFPageId = " & vCRFPageId _
            & " ORDER BY CRFElement.FieldOrder, CRFElement.QGroupFieldOrder"
    Else
        sSQL = "SELECT CRFElement.*, DataItemCode, DataItemName,DataType, DataItemLength, " _
            & "DataItemFormat, UnitOfMeasurement, Derivation, DataItemHelpText, " _
            & " DataItemCase, ClinicalTestCode " _
            & "FROM   CRFElement LEFT JOIN DataItem " _
            & "ON     CRFElement.ClinicalTrialId = DataItem.ClinicalTrialId " _
            & "AND    CRFElement.VersionId = DataItem.VersionId " _
            & "AND    CRFElement.DataItemId = DataItem.DataItemId " _
            & "WHERE  CRFElement.ClinicalTrialId = " & vClinicalTrialId _
            & " AND    CRFElement.VersionId = " & vVersionId _
            & " AND    CRFElement.CRFPageId = " & vCRFPageId _
            & " ORDER BY CRFElement.FieldOrder, CRFElement.QGroupFieldOrder"
    End If

    Set gdsCRFPageDataItems = New ADODB.Recordset
    gdsCRFPageDataItems.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsCRFPageDataItems", "CRFData")
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
Public Function gdsCRFDataElement(ByVal vClinicalTrialId As Long, _
                                  ByVal vVersionId As Integer, _
                                  ByVal vCRFPageId As Long, _
                                  ByVal vCRFElementId As Integer) As ADODB.Recordset
'---------------------------------------------------------------------
'Retrieve details of a data item on a CRF page
'Creates a ReadOnly recordset
'TA 18/09/2000 ClinicalTestCode now retireved as well
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    ' NCJ 16 Nov 99 - Added DataItemCase
    sSQL = "SELECT CRFElement.*, DataItemCode,DataItemName, DataItemLength, " _
        & "DataItemFormat, UnitOfMeasurement, DataType, DataItemCase, ClinicalTestCode " _
        & "  FROM   CRFElement, DataItem  " _
        & "  WHERE  CRFElement.ClinicalTrialId          = " & vClinicalTrialId _
        & "  AND    CRFElement.VersionId        = " & vVersionId _
        & " AND CRFElement.ClinicalTrialId          = DataItem.ClinicalTrialId " _
        & " AND CRFElement.VersionId                = DataItem.VersionId " _
        & " AND CRFElement.DataItemId               = DataItem.DataItemId " _
        & " AND CRFElement.CRFPageId                = " & vCRFPageId _
        & " AND CRFElementId                        = " & vCRFElementId
    
    Set gdsCRFDataElement = New ADODB.Recordset
    gdsCRFDataElement.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

    Exit Function

ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "gdsCRFDataElement", "CRFData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Public Function gdsCRFElement(ByVal vClinicalTrialId As Long, _
                              ByVal vVersionId As Integer, _
                              ByVal vCRFPageId As Long, _
                              ByVal vCRFElementId As Integer) As ADODB.Recordset
'---------------------------------------------------------------------
'Retrieve details of a CRF element
'Creates a ReadOnly recordset
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    sSQL = "SELECT CRFElement.* " _
        & "  FROM   CRFElement  " _
        & "  WHERE  CRFElement.ClinicalTrialId          = " & vClinicalTrialId _
        & "  AND    CRFElement.VersionId                = " & vVersionId _
        & "  AND    CRFElement.CRFPageId                = " & vCRFPageId _
        & "  AND    CRFElementId                        = " & vCRFElementId
    
    Set gdsCRFElement = New ADODB.Recordset
    gdsCRFElement.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

    Exit Function

ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsCRFElement", "CRFData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Private Function mnNextCRFPageId(ByVal vClinicalTrialId As Long, _
                                 ByVal vVersionId As Integer) As Integer
'---------------------------------------------------------------------
'Returns the next available unique id for a CRF page
'---------------------------------------------------------------------
Dim rsTmp As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler
    
    sSQL = "SELECT MAX(CRFPageId) as MaxCRFPageId FROM CRFPage " _
        & "  WHERE  ClinicalTrialId          = " & vClinicalTrialId _
        & "  AND    VersionId        = " & vVersionId
    
    Set rsTmp = New ADODB.Recordset
    rsTmp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If IsNull(rsTmp!MaxCRFPageId) Then     'if no current CRF pages in this trial
        mnNextCRFPageId = gnFIRST_ID
    Else
        mnNextCRFPageId = rsTmp!MaxCRFPageId + gnID_INCREMENT
    End If
    
    rsTmp.Close
    Set rsTmp = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "mnNextCRFPageId", "CRFData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Public Function gdsCRFPageDataItemsYorder(ByVal vClinicalTrialId As Long, _
                                    ByVal vVersionId As Integer, _
                                    ByVal vCRFPageId As Long) As ADODB.Recordset
'---------------------------------------------------------------------
'Creates a ReadOnly recordset
'---------------------------------------------------------------------
'The file CRFElement contains an empty field called PrintOrder. Each record within
'CRFElement pertains to a control and a caption (if it is of the type that requires
'a caption). When printing the forms it is neccessary to have a Y coordinate for each
'control that is the lesser of Y and CaptionY. This routine takes the required data
'into rsTemp and then reads through rsTemp setting the value of PrintOrder.
'the resulting recordset is then returned to the caller as gdsCRFPageDataItemsYorder
'
'Re-Written by Mo Morris 11/11/99, because the following line
'   rsTemp.Sort = "PrintOrder"
'would not work under ADO
'MIGHT WORK WHEN ALL OF DAO IS REMOVED
'
'Mo 29/4/2004   Bug 2259. Check for empty captions added. When assessing CaptionY and
'               (control)Y ignore CaptionY when the Caption is empty.
'---------------------------------------------------------------------

Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim nY As Long
Dim nCaptionY As Long
Dim sCaption As String
On Error GoTo ErrHandler

    'Mo 23/5/2003, Filtering on vClinicalTrialId, vVersionId and vCRFPageId added. They should
    'have been here in 1999 (naughty me). Filtering on non-RQG question elements added (OwnerQGroupId = 0)
    'Mo 29/4/2004, Bug 2259, Caption added to SELECT part of SQL
    sSQL = "SELECT Caption, CaptionY, Y, ControlType, PrintOrder FROM CRFElement" _
        & " WHERE CRFElement.ClinicalTrialId = " & vClinicalTrialId _
        & " AND CRFElement.VersionId = " & vVersionId _
        & " AND CRFElement.CRFPageId = " & vCRFPageId _
        & " AND CRFElement.OwnerQGroupId = 0"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText

'   ATN 4/2/2000
'   Need to use VB variables to update the recordset, rather than recordset-to-recordset update
    Do While Not rsTemp.EOF
        ' NCJ 31/10/00 - If Line, Comment or Picture use Y
        ' (old studies may have non-zero CaptionY)
        If rsTemp![CaptionY] = 0 Or rsTemp![ControlType] > gnVISUAL_ELEMENT Then
            nY = rsTemp![Y]
            rsTemp![PrintOrder] = nY
        Else
            'Mo 29/4/2004, Bug 2259
            'Check that the Caption is not empty
            If IsNull(rsTemp![Caption]) Then
                sCaption = ""
            Else
                sCaption = Trim(rsTemp![Caption])
            End If
            If Len(sCaption) = 0 Then
                'its an empty caption, so ignore CaptionY and use Y
                nY = rsTemp![Y]
                rsTemp![PrintOrder] = nY
            Else
                If rsTemp![CaptionY] < rsTemp![Y] Then
                    nCaptionY = rsTemp![CaptionY]
                    rsTemp![PrintOrder] = nCaptionY
                Else
                    nY = rsTemp![Y]
                    rsTemp![PrintOrder] = nY
                End If
            End If
        End If
        rsTemp.Update
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing

'   ATN 21/12/99
'   Oracle doesn't like LEFT JOIN.
    If goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
        sSQL = "SELECT CRFElement.*, DataItemCode, DataItemName,DataType, DataItemLength, DataItemFormat, " _
            & "  UnitOfMeasurement, Derivation, DataItemHelpText " _
            & "  FROM   CRFElement , DataItem " _
            & "  WHERE  DataItem.ClinicalTrialId      (+) = CRFElement.ClinicalTrialId " _
            & "  AND    DataItem.VersionId            (+) = CRFElement.VersionId " _
            & "  AND    DataItem.DataItemId           (+) = CRFElement.DataItemId " _
            & "  AND    CRFElement.ClinicalTrialId          = " & vClinicalTrialId _
            & "  AND    CRFElement.VersionId        = " & vVersionId _
            & "  AND    CRFElement.CRFPageId                = " & vCRFPageId _
            & "  AND    CRFElement.OwnerQGroupId = 0" _
            & "  ORDER by PrintOrder"
    Else
        sSQL = "SELECT CRFElement.*, DataItemCode, DataItemName,DataType, DataItemLength, DataItemFormat, " _
            & "  UnitOfMeasurement, Derivation, DataItemHelpText " _
            & "  FROM   CRFElement LEFT JOIN DataItem " _
            & "  ON     CRFElement.ClinicalTrialId      = DataItem.ClinicalTrialId " _
            & "  AND    CRFElement.VersionId    = DataItem.VersionId " _
            & "  AND    CRFElement.DataItemId           = DataItem.DataItemId " _
            & "  WHERE  CRFElement.ClinicalTrialId          = " & vClinicalTrialId _
            & "  AND    CRFElement.VersionId        = " & vVersionId _
            & "  AND    CRFElement.CRFPageId                = " & vCRFPageId _
            & "  AND    CRFElement.OwnerQGroupId = 0" _
            & "  ORDER by PrintOrder"
    End If
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
        
    'sort the recordset on the field PageOrder and return it
    'rsTemp.Sort = "PrintOrder"
    Set gdsCRFPageDataItemsYorder = New ADODB.Recordset
    Set gdsCRFPageDataItemsYorder = rsTemp

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsCRFPageDataItemsYorder", "CRFData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

