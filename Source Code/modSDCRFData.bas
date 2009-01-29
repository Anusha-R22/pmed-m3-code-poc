Attribute VB_Name = "modSDCRFData"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999-2004. All Rights Reserved
'   File:       modSDCRFData.bas
'   Author      Paul Norris, 24/09/99
'   Purpose:    All common TrialData functions for the StudyDefinition project
'               are in this module.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
' Revisions:
'   NCJ 15 Oct 99 - Get stored integer GridSize in gnInsertCRFElement
'   Mo Morris   11/11/99    DAO to ADO conversion
'   Mo Morris   16/11/99    DAO to ADO conversion
'   NCJ 3 Dec 99 - Sort out single quotes in SQL strings
'   NCJ 13 Dec 99 - Ids to Long
'                   SR 2375 - Include default colour for new CRFElements
'   NCJ 9 Feb 00 - Deal with quotes in CRFPageDate expression
'   NCJ 22 Feb 00 - Combined some UpdateCRFPage calls
'   TA 14/09/2000 sql for clinical test date expression in UpdateCRFElementProperties
'   TA 27/07/2001 - "Local" field in CRFElement table changed to "LocalFlag" because of JET4 probs
'   ZA 09/08/01 - added date prompt field for eForm
' NCJ 29 Nov 01 - Made changes to gnInsertCRFElement
' ASH 13/12/2001- added routine DataItemCodeUsedInStudy
'MACRO 3.0
' NCJ 7 Dec 01 - Added new routines to handle database operations on Question Groups & EForm Groups
' NCJ/REM 10 Dec 01 - Testing and debugging new routines
' NCJ 13/12/01 - Added ShowStatusFlag to UpdateCRFElementProperties
' NCJ 3 Jan 02 - Added DataItemCodeUsedInStudy routine from 2.2; bug fix to DBInsertCRFGroupMemberElement
' ZA 19/07/2002 - Added font and colour properties of caption in gnInsertCRFElement
' NCJ 15 Aug 02 - Sorted out null captions in gnInsertCRFElement
' ZA 22/08/2002 - Use null instead of 0 when inserting new caption font details
' TA 27/08/2002: Added ElementUse column value when inserting into CRFElement table
'                   in routines DBInsertCRFGroupMemberElement and gnInsertCRFElement
'ZA 09/09/2002 - Hide/show status icons for a question or group based on the menu
'ASH 4/11/2002 - Added EformWidth to UpdateCRFPageDetails
' NCJ 22 Nov 02 - Make sure ShowFlag = 0 for new Group elements
' NCJ 29 Jun 04 - Added ReplaceQuotes around Caption in DBInsertCRFGroupMemberElement
'----------------------------------------------------------------------------------------'

Option Explicit

'-------------------------------------------------------------------
Public Sub gDeleteCRFElement(ByVal vClinicalTrialId As Long, _
                            ByVal vVersionId As Integer, _
                            ByVal vCRFPageId As Long, _
                            ByVal vCRFElementId As Integer, _
                            ByVal vDataItemId As Long)
'-------------------------------------------------------------------
' Delete a CRF Element from the "CRFElement" table
'-------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    'Begin transaction
    TransBegin
    
    sSQL = "DELETE FROM CRFElement " _
        & "  WHERE  ClinicalTrialId          = " & vClinicalTrialId _
        & "  AND    VersionId                = " & vVersionId _
        & "  AND    CRFPageId                = " & vCRFPageId _
        & "  AND    CRFElementId             = " & vCRFElementId
                        
    MacroADODBConnection.Execute sSQL
    
    TransCommit
          
Exit Sub
ErrHandler:
    'RollBack transaction
    TransRollBack
    
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gDeleteCRFElement", "modSDCRFData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'---------------------------------------------------------------------
Public Function gdsCRFPageList(ByVal vClinicalTrialId As Long, _
                               ByVal vVersionId As Integer, _
                               Optional bFormAlphabeticOrder As Boolean) As ADODB.Recordset
'---------------------------------------------------------------------
' Retrieve all CRF pages in a trial
' PN change 21 - add bFormAlphabeticOrder parameter
'---------------------------------------------------------------------

Dim sSQL As String

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM CRFPage " _
        & " WHERE ClinicalTrialId =  " & vClinicalTrialId _
        & " AND VersionId =  " & vVersionId
    
    ' PN change 21
    If Not bFormAlphabeticOrder Then
        sSQL = sSQL & "  ORDER BY CRFPageOrder"
    Else
        sSQL = sSQL & "  ORDER BY CRFTitle"
    End If
    
    Set gdsCRFPageList = New ADODB.Recordset
    gdsCRFPageList.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
           
    Exit Function

ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsCRFPageList", "modSDCRFData")
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
Public Function gnInsertCRFElement(ByVal lClinicalTrialId As Long, _
                                   ByVal nVersionId As Integer, _
                                   ByVal lCRFPageId As Long, _
                                   ByVal lDataItemId As Long, _
                                   ByVal lGroupId As Long, _
                                   ByVal sglX As Single, _
                                   ByVal sglY As Single, _
                                   ByVal sglCaptionX As Single, _
                                   ByVal sglCaptionY As Single, _
                                   ByVal sCaption As String, _
                                   Optional nControlType As Variant) As Integer
'---------------------------------------------------------------------
' Insert a "top level" CRF element onto a CRF page (i.e. not a group member)
' If the data item is 0, it's a Visual item
' For questions, the ControlType is calculated
' Changed by Mo Morris 2/10/99 - Partial DAO to ADO conversion
' NCJ 3 Dec 99 - Deal with single quotes in vCaption
' NCJ 29/30 Nov 01 - Takes GroupID, sglCaptionX and sglCaptionY too, but NOT vForm
' NCJ 15 Aug 02 - Changed Caption from Variant to String
' TA 27/08/2002: Added ElementUse column value when inserting into CRFElement table
' ZA 09/09/2002 - Hide/show status icon based on Hide Icons menu
'---------------------------------------------------------------------
Dim sSQL As String
Dim nNewCRFElementId As Integer
Dim nFieldOrder As Integer
Dim enShowStatusFlag As eStatusFlag
Dim enRFCDefault As eRFCDefault

    On Error GoTo ErrHandler
    
    'Begin transaction
    TransBegin
    
    
    nNewCRFElementId = mnNextCRFElementId(lClinicalTrialId, nVersionId, lCRFPageId)
    
    If lDataItemId = gnZERO And lGroupId = gnZERO Then
        ' Not a question or a group
        nFieldOrder = gnZERO
        
    Else                            '   data item or group
        nFieldOrder = mnNextFieldOrder(lClinicalTrialId, nVersionId, lCRFPageId)
        
        'ZA 06/09/2002 - set the status flag property for "top level" CRF element
        ' NCJ 22 Nov 02 - Never have Status flags for groups
        If frmMenu.mnuHideIcons.Checked Or lGroupId > 0 Then
            enShowStatusFlag = eStatusFlag.Hide
        Else
            enShowStatusFlag = eStatusFlag.Show
        End If
        
        If lGroupId = 0 Then
            ' Not a group so RFC is according to current setting
            enRFCDefault = DefaultRFCOption
        Else
            ' It's a group so RFC always off
            enRFCDefault = eRFCDefault.RFCDefaultOff
        End If
        
    End If
    
    If lDataItemId > 0 Then ' a question
        ' Get the most appropriate control type
        nControlType = GetControlType(lClinicalTrialId, nVersionId, lDataItemId)
    End If
    
    ' Use single quotes and ReplaceQuotes round vCaption - NCJ 3 Dec 99
    If sCaption = "" Then
        sCaption = "NULL"
    Else
        sCaption = "'" & ReplaceQuotes(sCaption) & "'"
    End If
    
    ' NCJ 7 Jan 00 - Store colour as 0 instead so as to use current default
    ' Mo Morris 30/8/01 Db Audit (FontBold, FontItalic, FontSize, Mandatory, RequireComment with defaults of 0 added)
    ' NCJ 29 Nov 01 - Added extra QuestionGroup fields
    ' ZA 19/07/2002 - Added caption font properties
    sSQL = "INSERT INTO CRFElement (ClinicalTrialId, VersionId, " _
        & "CRFPageId, CRFElementId, DataItemId, " _
        & "X, Y, CaptionX, CaptionY, Caption, " _
        & "ControlType, FieldOrder, Optional, Hidden, LocalFlag, " _
        & "FontColour, FontBold, FontItalic, FontSize, " _
        & "Mandatory, RequireComment, " _
        & "OwnerQGroupId, QGroupId, QGroupFieldOrder, ShowStatusFlag, " _
        & "CaptionFontBold, CaptionFontItalic, CaptionFontSize, CaptionFontColour," _
        & "ElementUse)" _
        & " VALUES (" & lClinicalTrialId & "," & nVersionId & ", " _
        & lCRFPageId & "," & nNewCRFElementId & "," & lDataItemId & ", " _
        & sglX & "," & sglY & "," & sglCaptionX & "," & sglCaptionY & ", " & sCaption & ", " _
        & nControlType & "," & nFieldOrder & ",0,0,0," _
        & "0,0,0,0, " _
        & "0," & enRFCDefault & ", " _
        & "0, " & lGroupId & ", 0," & enShowStatusFlag & ", " _
        & "null,null,null,null," & eElementUse.User & ")"
                        
    MacroADODBConnection.Execute sSQL
    
    'End transaction
    TransCommit
    
    gnInsertCRFElement = nNewCRFElementId
           
    Exit Function

ErrHandler:
    'RollBack transaction
    TransRollBack
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "gnInsertCRFElement", "modSDCRFData")
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
Public Sub DBInsertCRFGroupMemberElements(ByVal lClinicalTrialId As Long, _
                                   ByVal nVersionId As Integer, _
                                   ByVal lDataItemId As Long, _
                                   ByVal lGroupId As Long, _
                                   ByVal nQGroupFieldOrder As Integer)
'---------------------------------------------------------------------
' Insert new group member CRF elements for this DataItem
' onto every CRF page which uses this group
'REM 06/03/02 - Assigned the recordset to an array
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsCRFPages As ADODB.Recordset
Dim nErr As Integer
Dim sErrDesc As String
Dim i As Integer
Dim vData As Variant

    On Error GoTo Errlabel
    
    ' Find all the CRFPages which use this group
    sSQL = "SELECT CRFPageID from CRFElement " _
            & " WHERE ClinicalTrialId = " & lClinicalTrialId _
            & " AND  VersionId    = " & nVersionId _
            & " AND QGroupId = " & lGroupId
    Set rsCRFPages = New ADODB.Recordset
    rsCRFPages.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    'REM 06/03/02 - assign recordset to an array
    If rsCRFPages.RecordCount > 0 Then

        vData = rsCRFPages.GetRows
        rsCRFPages.Close
        Set rsCRFPages = Nothing
    
        'looping thru array to perform delete
        For i = 0 To UBound(vData, 2)
            Call DBInsertCRFGroupMemberElement(lClinicalTrialId, _
                                       nVersionId, _
                                       CLng(vData(0, i)), _
                                       lDataItemId, _
                                       lGroupId, _
                                       0, _
                                       nQGroupFieldOrder)
        Next
    
    Else
        rsCRFPages.Close
        Set rsCRFPages = Nothing
    End If
    
    'Commented out by REM 06/03/02 as changed to looping through an array
'    Do While Not rsCRFPages.EOF
'        Call DBInsertCRFGroupMemberElement(lClinicalTrialId, _
'                                   nVersionId, _
'                                   rsCRFPages!CRFPageId, _
'                                   lDataItemId, _
'                                   lGroupId, _
'                                   0, _
'                                   nQGroupFieldOrder)
'        rsCRFPages.MoveNext
'    Loop
'    rsCRFPages.Close
'    Set rsCRFPages = Nothing
            
Exit Sub
Errlabel:
    nErr = Err.Number
    sErrDesc = Err.Description & "|modSDCRFData.DBInsertCRFGroupMemberElements"
    'RollBack transaction
    TransRollBack
    Err.Raise nErr, , sErrDesc

End Sub

'---------------------------------------------------------------------
Public Function DBInsertCRFGroupMemberElement(ByVal lClinicalTrialId As Long, _
                                   ByVal nVersionId As Integer, _
                                   ByVal lCRFPageId As Long, _
                                   ByVal lDataItemId As Long, _
                                   ByVal lOwnerGroupID As Long, _
                                   ByVal nFieldOrder As Integer, _
                                   ByVal nQGroupFieldOrder As Integer) As Integer
'---------------------------------------------------------------------
' Insert a group member CRF element onto a CRF page
' Assume it's a question
' The the ControlType and Caption are calculated
' If nFieldOrder = 0, we figure out what it should be
' TA 27/08/2002: Added ElementUse column value when inserting into CRFElement table
' ZA 09/09/2002 - Hide/show status icon based on Hide RQG icons menu
' NCJ 29 Jun 04 - Added ReplaceQuotes around Caption
'---------------------------------------------------------------------
Dim sSQL As String
Dim nNewCRFElementId As Integer
Dim sCaption As String
Dim nControlType As Integer
Dim nOwnerFieldOrder As Integer
Dim rsTemp As ADODB.Recordset
Dim nErr As Integer
Dim sErrDesc As String
Dim enShowStatusFlag As eStatusFlag

    On Error GoTo Errlabel
    
    'Begin transaction
    TransBegin
    
    ' Do we need to find out the FieldOrder?
    If nFieldOrder = 0 Then
        ' Find the FieldOrder of the owning group element
        sSQL = "SELECT FieldOrder from CRFElement " _
            & " WHERE ClinicalTrialId = " & lClinicalTrialId _
            & " AND  VersionId    = " & nVersionId _
            & " AND  CRFPageId    = " & lCRFPageId _
            & " AND QGroupId = " & lOwnerGroupID
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        nOwnerFieldOrder = rsTemp!FieldOrder
        rsTemp.Close
        Set rsTemp = Nothing
    Else
        ' Use the one we've been given
        nOwnerFieldOrder = nFieldOrder
    End If
    
    nNewCRFElementId = mnNextCRFElementId(lClinicalTrialId, nVersionId, lCRFPageId)
        
    nControlType = GetControlType(lClinicalTrialId, nVersionId, lDataItemId)
    
    sCaption = DataItemNameFromId(lClinicalTrialId, lDataItemId)
    'ZA 09/09/2002 - check if this question will have it status flag shown
    If frmMenu.mnuHideRQGIcons.Checked Then
        enShowStatusFlag = eStatusFlag.Hide
    Else
        enShowStatusFlag = eStatusFlag.Show
    End If
    
    ' ZA 29/07/2002 - Added caption font properties
    ' NCJ 29 Jun 04 - Added ReplaceQuotes around Caption
    sSQL = "INSERT INTO CRFElement (ClinicalTrialId, VersionId, " _
        & "CRFPageId, CRFElementId, DataItemId, " _
        & "X, Y, CaptionX, CaptionY, Caption, " _
        & "ControlType, FieldOrder, Optional, Hidden, LocalFlag, " _
        & "FontColour, FontBold, FontItalic, FontSize, " _
        & "Mandatory, RequireComment, " _
        & "OwnerQGroupId, QGroupId, QGroupFieldOrder, ShowStatusFlag, " _
        & "CaptionFontBold, CaptionFontItalic, CaptionFontSize, CaptionFontColour, ElementUse)" _
        & "VALUES (" _
        & lClinicalTrialId & "," & nVersionId & ", " _
        & lCRFPageId & "," & nNewCRFElementId & "," & lDataItemId & ", " _
        & "0,0,0,0, '" & ReplaceQuotes(sCaption) & "', " _
        & nControlType & "," & nOwnerFieldOrder & ",0,0,0," _
        & "0,0,0,0, " _
        & "0," & DefaultRFCOption & ", " _
        & lOwnerGroupID & ", 0," & nQGroupFieldOrder & ", " & enShowStatusFlag & ", " _
        & "null,null,null,null," & eElementUse.User & ")"
                        
    MacroADODBConnection.Execute sSQL
    
    'End transaction
    TransCommit
    
    DBInsertCRFGroupMemberElement = nNewCRFElementId
           
Exit Function
Errlabel:
    nErr = Err.Number
    sErrDesc = Err.Description & "|modSDCRFData.DBInsertCRFGroupMemberElement"
    'RollBack transaction
    TransRollBack
    Err.Raise nErr, , sErrDesc
    
End Function

'---------------------------------------------------------------------
Public Sub DBDeleteQuestionGroup(lStudyId As Long, nVersionId As Integer, lQGroupId As Long)
'---------------------------------------------------------------------
' REM 11/12/01
' Deletes a Question Group from all 4 tables that hold question group information
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo Errlabel
    
    'Delete from QGroup table
    sSQL = "DELETE FROM QGroup" & _
           " WHERE ClinicalTrialID = " & lStudyId & _
           " AND VersionID = " & nVersionId & _
           " AND QGroupID = " & lQGroupId
    MacroADODBConnection.Execute sSQL

    'Delete from QGroupQuestion table
    sSQL = "DELETE FROM QGroupQuestion" & _
           " WHERE ClinicalTrialID = " & lStudyId & _
           " AND VersionID = " & nVersionId & _
           " AND QGroupID = " & lQGroupId
    MacroADODBConnection.Execute sSQL

    'Delete QGroup from EFormQGroup table
    sSQL = "DELETE FROM EFormQGroup" & _
            " WHERE ClinicalTrialID = " & lStudyId & _
            " AND VersionID = " & nVersionId & _
            " AND QGroupID = " & lQGroupId
    MacroADODBConnection.Execute sSQL

    'Delete QGroup from CRFElemet table
    sSQL = "DELETE FROM CRFElement" & _
            " WHERE ClinicalTrialID = " & lStudyId & _
            " AND VersionID = " & nVersionId & _
            " AND (QGroupID = " & lQGroupId & _
            " OR OwnerQGroupID = " & lQGroupId & ")"
    MacroADODBConnection.Execute sSQL

Exit Sub
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|modSDCRFData.DBDeleteQuestionGroup"
End Sub

'---------------------------------------------------------------------
Private Function GetControlType(ByVal lClinicalTrialId As Long, _
                                ByVal nVersionId As Integer, _
                                ByVal lDataItemId As Long) As Integer
'---------------------------------------------------------------------
' Get the most appropriate Control type for this data item
'---------------------------------------------------------------------
Dim rsDataItem As ADODB.Recordset
Dim rsDataValues As ADODB.Recordset
Dim nControlType As Integer

    Set rsDataItem = New ADODB.Recordset
    ' Get the data item info
    Set rsDataItem = gdsDataItem(lClinicalTrialId, nVersionId, lDataItemId)

    ' Work out the best ControlType for the question
    If rsDataItem!DataType = DataType.Category Then    'category
        Set rsDataValues = New ADODB.Recordset
        Set rsDataValues = gdsDataValues(lClinicalTrialId, nVersionId, rsDataItem!DataItemId)
        If rsDataValues.EOF Then
            nControlType = gn_TEXT_BOX
        Else
            rsDataValues.MoveLast
            If rsDataValues.RecordCount <= gnUseOptionButton Then
           ' If rsDataValues.RecordCount > gnRECOMMENDED_MAXIMUM_NUMBER_OF_OPTION_BUTTONS Then
                nControlType = gn_OPTION_BUTTONS
            Else
                nControlType = gn_POPUP_LIST
            End If
        End If
        rsDataValues.Close
        Set rsDataValues = Nothing
    ElseIf rsDataItem!DataType = DataType.Multimedia Then    'multimedia
        nControlType = gn_ATTACHMENT
    Else
        nControlType = gn_TEXT_BOX
    End If
    
    rsDataItem.Close
    Set rsDataItem = Nothing
    
    GetControlType = nControlType

End Function

'---------------------------------------------------------------------
Public Sub DBReorderCRFGroupElements(ByVal lClinicalTrialId As Long, _
                                   ByVal nVersionId As Integer, _
                                   ByVal lQGroupId As Long)
'---------------------------------------------------------------------
' Set the QGroupFieldOrder field for each CRFElement group member
' to be the same as the QOrder field from its group definition
' within this study
'-----------------------------------------------------------------------
Dim sSQL As String
Dim sSQLValue As String
Dim rsQOrders As ADODB.Recordset

    On Error GoTo Errlabel
    
    ' Get the DataItemIds and their group orders
    sSQL = "SELECT DataItemId, QOrder FROM QGroupQuestion " _
            & " WHERE ClinicalTrialId = " & lClinicalTrialId _
            & " AND   VersionId       = " & nVersionId _
            & " AND   QGroupId   = " & lQGroupId
    Set rsQOrders = New ADODB.Recordset
    rsQOrders.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    ' For each one, update any entries in the CRFElement table
    Do While Not rsQOrders.EOF
        sSQL = "UPDATE CRFElement " _
            & " SET QGroupFieldOrder = " & rsQOrders!QOrder _
            & " WHERE ClinicalTrialId = " & lClinicalTrialId _
            & " AND   VersionId       = " & nVersionId _
            & " AND   OwnerQGroupId   = " & lQGroupId _
            & " AND   DataItemId   = " & rsQOrders!DataItemId
            
        MacroADODBConnection.Execute sSQL
        rsQOrders.MoveNext
    Loop
    
    rsQOrders.Close
    Set rsQOrders = Nothing
    
Exit Sub
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|modSDCRFData.DBReorderCRFGroupElement"

End Sub

'---------------------------------------------------------------------
Public Sub DBDeleteCRFGroupElements(ByVal lClinicalTrialId As Long, _
                                   ByVal nVersionId As Integer, _
                                   ByVal lQGroupId As Long, _
                                   ByVal lDataItemId As Long)
'---------------------------------------------------------------------
' Delete all the CRFElement group members
' that use this lDataItemId in this GroupId
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo Errlabel
    
    sSQL = "DELETE FROM CRFElement " _
            & " WHERE ClinicalTrialId = " & lClinicalTrialId _
            & " AND   VersionId       = " & nVersionId _
            & " AND   OwnerQGroupId   = " & lQGroupId _
            & " AND   DataItemId   = " & lDataItemId
    
    MacroADODBConnection.Execute sSQL
    
Exit Sub
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|modSDCRFData.DBDeleteCRFGroupElements"

End Sub

'---------------------------------------------------------------------
Public Function mnNextCRFPageOrder(ByVal vClinicalTrialId As Long, _
                                 ByVal vVersionId As Integer) As Integer
'---------------------------------------------------------------------
'Returns the next available unique id for a CRF page
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim sSQL As String

On Error GoTo ErrHandler

    sSQL = "SELECT MAX(CRFPageOrder) as MaxCRFPageOrder FROM CRFPage " _
        & "  WHERE ClinicalTrialId = " & vClinicalTrialId _
        & "  AND   VersionId       = " & vVersionId
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If IsNull(rsTemp!MaxCRFPageOrder) Then     'if no current CRF pages in this trial
        mnNextCRFPageOrder = gnFIRST_ID
    Else
        mnNextCRFPageOrder = rsTemp!MaxCRFPageOrder + gnID_INCREMENT
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
       
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "mnNextCRFPageOrder", "modSDCRFData")
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
Public Function mnNextFieldOrder(ByVal vClinicalTrialId As Long, _
                                  ByVal vVersionId As Integer, _
                                  ByVal vCRFPageId As Long) As Integer
'---------------------------------------------------------------------
'Returns the next available field order number for a CRF page
'---------------------------------------------------------------------

Dim rsTemp As ADODB.Recordset
Dim sSQL As String

On Error GoTo ErrHandler

    sSQL = "SELECT MAX(FieldOrder) as MaxFieldOrder FROM CRFElement " _
        & "  WHERE  ClinicalTrialId          = " & vClinicalTrialId _
        & "  AND    VersionId                = " & vVersionId _
        & "  AND    CRFPageId                = " & vCRFPageId
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If IsNull(rsTemp!MaxFieldOrder) Then     'if no CRF elements on this page
        mnNextFieldOrder = gnFIRST_ID
    Else
        mnNextFieldOrder = rsTemp!MaxFieldOrder + gnID_INCREMENT
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
       
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "mnNextFieldOrder", "modSDCRFData")
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
Public Sub UpdateCRFElementProperties(ByVal vClinicalTrialId As Long, _
                                        ByVal vVersionId As Integer, _
                                        ByVal vCRFPageId As Long, _
                                        ByVal vCRFElementId As Integer, _
                                        ByVal vOptional As Integer, _
                                        ByVal vHidden As Integer, _
                                        ByVal vLocal As Integer, _
                                        ByVal vSkipCondition As String, _
                                        ByVal vMandatory As Integer, _
                                        ByVal vRequireComment As Integer, _
                                        ByVal vAuthorisationLevel As String, _
                                        ByVal sLabTestDate As String, _
                                        ByVal nShowFlag As Integer, _
                                        ByVal enElementUse As eElementUse, _
                                        ByVal nDisplayLength As Integer, _
                                        ByVal sDescription As String)
'---------------------------------------------------------------------
' Update CRFElement properties
' NCJ 13 Dec 01 - Added ShowFlag property
' ZA 21/08/2002 - Added ElementUse property
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    'Begin transaction
    TransBegin
    
    ' NCJ 13/12/01 - Added ShowStatusFlag
    sSQL = "UPDATE CRFElement " _
            & " SET Optional = " & vOptional _
            & " , Hidden = " & vHidden _
            & " , LocalFlag = " & vLocal _
            & " , ShowStatusFlag = " & nShowFlag _
            & " , ElementUse = " & enElementUse _
            & " , SkipCondition = "
    
    If vSkipCondition > "" Then
        sSQL = sSQL & "'" & ReplaceQuotes(vSkipCondition) & "'"
    Else
        sSQL = sSQL & "NULL"
    End If
    
   
    sSQL = sSQL & ", Mandatory=" & vMandatory
    sSQL = sSQL & ", RequireComment=" & vRequireComment
    
    ' PN change 23 new RoleCode field - 06/09/99
    sSQL = sSQL & ", RoleCode='" & vAuthorisationLevel & "'"
    
    
    'TA 14/09/2000 sql for clinical test date expression
    sSQL = sSQL & " , ClinicalTestDateExpr = "
    If sLabTestDate <> "" Then
        sSQL = sSQL & "'" & ReplaceQuotes(sLabTestDate) & "'"
    Else
        sSQL = sSQL & "NULL"
    End If
    
    'ZA 13/09/2002 - display length for a data item
    sSQL = sSQL & " , DisplayLength = " & SQL_ValueToStringValue(ConvertToNull(nDisplayLength, vbInteger))
    
    'TA 18/06/2003 - description of a CRFElement
    sSQL = sSQL & " , DESCRIPTION = " & SQL_ValueToStringValue(ConvertToNull(sDescription, vbString))
    
    
    sSQL = sSQL & "  WHERE  ClinicalTrialId          = " & vClinicalTrialId _
                & "  AND    VersionId                = " & vVersionId _
                & "  AND    CRFPageId                = " & vCRFPageId _
                & "  AND    CRFElementId             = " & vCRFElementId
    
    MacroADODBConnection.Execute sSQL
         
    'UpdateProformaSkip vCRFPageId, vCRFElementId, vSkipCondition
    
    'End transaction
    TransCommit
       
    Exit Sub

ErrHandler:
    'RollBack transaction
    TransRollBack
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                            "UpdateCRFElementProperties", "modSDCRFData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'---------------------------------------------------------------------
Public Sub UpdateCRFPageTitle(ByVal vClinicalTrialId As Long, _
                              ByVal vVersionId As Integer, _
                              ByVal vCRFPageId As Long, _
                              ByVal vCRFTitle As String)
'---------------------------------------------------------------------
' NCJ 22 Feb 00 - THIS ROUTINE NO LONGER USED
' (See UpdateCRFPageDetails instead)
'---------------------------------------------------------------------
'Dim sSQL As String
'
'    On Error GoTo ErrHandler
'
'    sSQL = "UPDATE CRFPage SET " _
'            & " CRFTitle = '" & ReplaceQuotes(vCRFTitle) & "'" _
'            & "  WHERE ClinicalTrialId = " & vClinicalTrialId _
'            & "  AND   VersionId       = " & vVersionId _
'            & "  AND   CRFPageId       = " & vCRFPageId
'
'    MacroADODBConnection.Execute sSQL
'
'Exit Sub
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpdateCRFPageTitle", "modSDCRFData")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            End
'   End Select
    
End Sub

'---------------------------------------------------------------------
Public Sub UpdateCRFPageDetails(ByVal lClinicalTrialId As Long, _
                                ByVal nVersionId As Integer, _
                                ByVal lCRFPageId As Long, _
                                ByVal sCRFTitle As String, _
                                ByVal sCRFPageLabel As String, _
                                ByVal nLocalCRFPageLabel As Integer, _
                                ByVal sFormDateExpr As String, _
                                ByVal nSequentialEntry As Integer, _
                                ByVal nHideIfInactive As Integer, _
                                ByVal nDisplayNumbers As Integer, _
                                ByVal neFormDatePrompt As Integer, _
                                ByVal lEformWidth As Long)
'---------------------------------------------------------------------
' NCJ 3 Dec 99 - nSequentialEntry changed from String to Integer
' SDM 06/01/00 SR2392    Added HideIfInactive
' NCJ 22 Feb 00 Added CRFTitle, CRFPageLabel, LocalCRFPageLabel,
'       FormDateExpr and DisplayNumbers
' ZA 09/08/01, added eFormDatePrompt
' ASH 4/11/2002 - Added EformWidth
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    ' NCJ 3 Dec 99 - Removed quotes round nSequentialEntry
    sSQL = "UPDATE CRFPage SET " _
            & " CRFTitle = '" & ReplaceQuotes(sCRFTitle) & "', " _
            & " CRFPageLabel = '" & ReplaceQuotes(sCRFPageLabel) & "', " _
            & " LocalCRFPageLabel = " & nLocalCRFPageLabel & ", " _
            & " CRFPageDateLabel = '" & ReplaceQuotes(sFormDateExpr) & "', " _
            & " SequentialEntry = " & nSequentialEntry & ", " _
            & " HideIfInactive = " & nHideIfInactive & ", " _
            & " DisplayNumbers = " & nDisplayNumbers & ", " _
            & " eFormDatePrompt = " & neFormDatePrompt & ", " _
            & " EformWidth = " & lEformWidth _
            & "  WHERE ClinicalTrialId = " & lClinicalTrialId _
            & "  AND   VersionId       = " & nVersionId _
            & "  AND   CRFPageId       = " & lCRFPageId
    
    MacroADODBConnection.Execute sSQL
       
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpdateCRFPageDetails", "modSDCRFData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
    
End Sub

'---------------------------------------------------------------------
Public Sub UpdateCRFFieldOrder(ByVal lClinicalTrialId As Long, _
                               ByVal nVersionId As Integer, _
                               ByVal lCRFPageId As Long, _
                               ByVal nCRFElementID As Integer, _
                               ByVal nFieldOrder As Integer, _
                               ByVal lQGroupId As Long)
'---------------------------------------------------------------------
' NCJ 14 May 03 - Added in QGroup consideration if lQGroupId > 0
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    sSQL = "UPDATE CRFElement SET " _
        & "  FieldOrder = " & nFieldOrder _
        & "  WHERE ClinicalTrialId = " & lClinicalTrialId _
        & "  AND   VersionId       = " & nVersionId _
        & "  AND   CRFPageId       = " & lCRFPageId _
        & "  AND   CRFElementId    = " & nCRFElementID
    
    MacroADODBConnection.Execute sSQL
    
    If lQGroupId > 0 Then
        ' NCJ 14 May 03 - Also update this group's items
        sSQL = "UPDATE CRFElement SET " _
            & "  FieldOrder = " & nFieldOrder _
            & "  WHERE ClinicalTrialId = " & lClinicalTrialId _
            & "  AND   VersionId       = " & nVersionId _
            & "  AND   CRFPageId       = " & lCRFPageId _
            & "  AND   OwnerQGroupId   = " & lQGroupId
        
        MacroADODBConnection.Execute sSQL
    End If
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modSDCRFData.UpdateCRFFieldOrder"
    
End Sub

'-----------------------------------------------------------------------------------------------------
Public Function DataItemCodeUsedInStudy(lClinicalTrialId As Long, sDataItemCode As String) As Boolean
'-----------------------------------------------------------------------------------------------------
' 13/12/2001 ASH
'checks if question has been used in current study
'-----------------------------------------------------------------------------------------------------
Dim rsDataItem As ADODB.Recordset
Dim rsCRFdataItem As ADODB.Recordset
Dim sSQL As String
Dim nCount As Integer

    On Error GoTo ErrorHandler
        
    sSQL = "Select DataItemID from DataItem where ClinicaltrialId = " & lClinicalTrialId _
    & " AND " & GetSQLStringEquals("DataItemCode", sDataItemCode)
    
    'check to see if selected question code exist in DataItem table
    Set rsDataItem = New ADODB.Recordset
    rsDataItem.CursorLocation = adUseClient
    rsDataItem.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic
    nCount = rsDataItem.RecordCount
    If nCount <= 0 Then
        DataItemCodeUsedInStudy = False
    Else
        'check to see if selected question code exist in CRFElement table
        sSQL = "Select * from CRFElement where ClinicaltrialId = " & lClinicalTrialId _
        & " AND DataItemID = " & rsDataItem.Fields(0).Value
    
        Set rsCRFdataItem = New ADODB.Recordset
        rsCRFdataItem.CursorLocation = adUseClient
        rsCRFdataItem.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic
        nCount = rsCRFdataItem.RecordCount
            If nCount <= 0 Then
                DataItemCodeUsedInStudy = False
            Else
                DataItemCodeUsedInStudy = True
            End If
    End If
                
Exit Function
ErrorHandler:
    Err.Raise Err.Number, , Err.Description & "|modSDCRFData.DataItemCodeUsedInStudy"
    
End Function

'-----------------------------------------------------------------------------------------------------
Public Function DefaultRFCOption() As eRFCDefault
'-----------------------------------------------------------------------------------------------------
'ZA 29/08/2002 - checks if the study as default RFC set for all questions
'ZA 10/09/2002 - get this value from Reason for change menu under options
'-----------------------------------------------------------------------------------------------------
    
    If frmMenu.mnuDefaultRFC.Checked = True Then
        DefaultRFCOption = eRFCDefault.RFCDefaultOn
    Else
        DefaultRFCOption = eRFCDefault.RFCDefaultOff
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, , Err.Description & "|modSDCRFData.DefaultRFCOption"
End Function
