Attribute VB_Name = "modBatchDataEntry"
'----------------------------------------------------------------------------------------'
' File:         modBatchDataEntry
' Copyright:    InferMed Ltd. 2000-2004. All Rights Reserved
' Author:       Mo Morris, July 2002
' Purpose:      Contains variable declarations and facilities required by the Batch Data Entry Module
'----------------------------------------------------------------------------------------'
'   Mo Morris December 2002
'
'   Code  added to handle MACRO_BD's (MACRO Batch Data Entry Module) Command Line
'   facilities. There are 2 different Command line calls:
'
'       Batch Import - Importing a named file of Batch response entries into MACRO's
'       Response Buffer (the BatchResponseData table)
'
'       Batch Upload - Upload the current contents of MACRO's Response Buffer (the
'       BatchResponseData table)
'
'   Note that when a Batch Import is successfully completed Batch Upload will
'   be called.
'
'   Mo Morris 28/1/2008 Bug 2979, RoleCode added to command line login
'   The syntax for a Batch Import Commamd Line call is:-
'
'       MACRO_BD /BI/UserName/Password/DatabaseCode/RoleCode/PathAndNameOfFile
'
'   The syntax for a Batch Upload Commamd Line call is:-
'
'       MACRO_BD /BU/UserName/Password/DatabaseCode/RoleCode
'
'   If the Command Line Parameters are erroneous a log file is created with
'   the name BDCLyyyymmddhhmmss.log and the reason for the failure is written to it.
'   BDCL stands for Batch Data Command Line.
'
'----------------------------------------------------------------------------------------'
'   Revisions:
' Mo Morris 13/1/2003    BDCommandLineLoginOK and CreateBDCLLogFile moved to basMainMACROModule
' NCJ 6 May 03 - Added UserNameFull and UserRole to LoadSubject
' NCJ 29 Jan 04 - Must deal with single quotes occurring in sMessage in SetUploadMessage
' TA 20/04/2004: remove null on the response value because code can't handle nulls
' NCJ 28 Jun 04 - Get a lab def for a lab test question, and do registration automatically
' NCJ 14 Jul 04 - Do case-insensitive matching in ValidCategoryResponse
' Mo 21/6/2005 bug 2538, Prevent site codes containing upper case characters getting into the database.
' NCJ 21 Jun 06 - Bug 2718 - Enabled Registration in SaveFormResponses
' Mo 13/2/2007  - Bug 2877 - User Code Correction to a Batch Data Entry Message
' Mo 14/2/2007  - Bug 2876 - Correct Batch Data Entry's handling of multiple roles when checking a users
'                 ability to change data (UserCanChangeData) and authorise questions (CheckAuthorisation)
' NCJ 27 Sept 07 - Bugs 2935 & 2937 - Sorting out User name and User name full for derived & authorisation questions
' Mo 5/2/2008   Bug 3010. Minor changes around the Batch Upload (/BU) calls to MACRO_BD.
'               New subs UnlockBatchUpload & LockBatchUpload together with function BatchUploadRunning added.
'               UploadBatchResponses now calls BatchUploadRunning, LockBatchUpload & UnlockBatchUpload.
'               Enumeration eBatchUploadStatus added.
' Mo 14/3/2008  Bug 3010. The Lock that indicates that Batch Data Entry Upload is running has been changed
'               from
'                   table BatchUploadLock(LockStatus) being 0 (not running) or 1 (running)
'               to
'                   table MACROUserSetting(UserName,UserSetting,SettingValue) existing with values ('BUL','Batch Upload Lock',1)
'               LockBatchUpload and UnlockBatchUpload have been changed, BatchUploadRunning has been removed.
'               Enumeration eBatchUploadStatus aremoved.
' Mo 26/6/2008  WO-080002 - Bug 3042 - Unobtainable status Changes
'       New global variable:-
'           Public gnSelUnobtainable As Integer
'       Changed constant:-
'           Public Const gnMINFORMWIDTH As Integer = 15400
'       Changed subroutines:-
'           ValidBatchResponseLine
'           ImportBatchResponseLine
'           UploadBatchResponses
'----------------------------------------------------------------------------------------'

Option Explicit

Public glSelTrialId As Long
Public gsSelSite As String
Public glSelPersonId As Long
Public gsSelLabel As String
Public glSelDataItemId As Long
Public glSelRepeatNumber As Long
Public glSelCRFPageId As Long
Public glSelEFormCycle As Long
Public gdblSelEFormDate As Double
Public glSelVisitId As Long
Public glSelVisitCycle As Long
Public gdblSelVisitDate As Double
Public gsSelResponse As String
Public gsSelCatCode As String
'Mo 26/6/2008 - WO-080002
Public gnSelUnobtainable As Integer

Public goArezzo As Arezzo_DM

'Mo 26/6/2008 - WO-080002, gnMINFORMWIDTH increased from 14400 to 15400
Public Const gnMINFORMWIDTH As Integer = 15400
Public Const gnMINFORMHEIGHT As Integer = 7500

Public Const gnMINFORMWIDTHERRORSFORM As Integer = 8000
Public Const gnMINFORMHEIGHTERRORSFORM As Integer = 5000

Public Const gnMINFORMWIDTHGENERATOR As Integer = 6000
Public Const gnMINFORMHEIGHTGENERATOR  As Integer = 6000

Public gbEditTakingPlace As Boolean

Private moStudyDef As StudyDefRO
Private mlImportCount As Long

' Store whether we've already set up an eForm's lab
Private mbFormLabSet As Boolean

'---------------------------------------------------------------------
Public Function ValidBatchResponseFile(ByVal sBatchResponseFile As String) As Boolean
'---------------------------------------------------------------------
' This function validates the contents of a Batch Response File.
' NCJ 29 Jun 04 - Some code removed to separate routine ValidBatchResponseLine
'---------------------------------------------------------------------
Dim nIOFileNumber As Integer
Dim sBRFLine As String
Dim lLineCount As Long
Dim sErrorMessages As String
Dim bErrorsExist As Boolean

    On Error GoTo ErrHandler

    Call HourglassOn
    
    'open the Batch Response File
    nIOFileNumber = FreeFile
    Open sBatchResponseFile For Input As #nIOFileNumber
    
    sErrorMessages = ""
    lLineCount = 0
    bErrorsExist = False
    
    'Read the Batch Response File line by line
    Do While Not EOF(nIOFileNumber)
        lLineCount = lLineCount + 1
        frmMenu.txtProgress.Text = "Validating line no. " & lLineCount
        frmMenu.txtProgress.Refresh
        Line Input #nIOFileNumber, sBRFLine
        
        ' NCJ 29 Jun 04 - Call new line validation routine, and only for non-blank lines
        If Trim(sBRFLine) > "" Then
            'Set the overall Errors Exist flag if line validation fails
            If Not ValidBatchResponseLine(sBRFLine, lLineCount, sErrorMessages) Then
                bErrorsExist = True
            End If
        End If
    Loop
    
    'close the input file
    Close #nIOFileNumber
    
    Call HourglassOff
    
    'If errors have occurred display them all
    If bErrorsExist Then
        'Don't display Batch Response File Errors Form if running in command line mode
        If Not UCase(Left(Command, 3)) = "/BI" Then
            frmMenu.txtProgress.Text = "Validation Failed"
            frmMenu.txtProgress.Refresh
            frmBRFErrors.txtErrors = sErrorMessages
            frmBRFErrors.Show vbModal
        End If
        ValidBatchResponseFile = False
    Else
        mlImportCount = lLineCount
        ValidBatchResponseFile = True
    End If
    
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ValidBatchResponseFile", "modBatchDataEntry")
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
Private Function ValidBatchResponseLine(ByVal sBRFLine As String, _
                                        ByVal lLineCount As Long, _
                                        ByRef sErrors As String) As Boolean
'---------------------------------------------------------------------
' NCJ 29 Jun 04 - This routine created from code in ValidBatchResponseFile
' This function validates the contents of a single line from the Batch Response File.
' Returns TRUE if no errors were found, or FALSE otherwise
' Error messages are added to the sErrors string
'
' Each line should have the following makeup:-
'
'Element No.    Name                    Type        Details
'   0           ClinicalTrialName       Text 15     Mandatory
'   1           Site                    Text 8      Mandatory
'   2           PersonId                Long        Optional with SubjectLabel
'   3           SubjectLabel            Text 50     Optional with PersonId
'   4           VisitCode               Text 15     Mandatory
'   5           VisitCycleNumber        Integer     Optional with VisitCycleDate
'   6           VisitCycleDate          Text 10     Optional with VisitCyclenumber
'   7           CRFPageCode             Text 15     Mandatory
'   8           CRFPageCycleNumber      Integer     Optional with CRFPageCycleDate
'   9           CRFPageCycelDate        Text 10     Optional with CRFPageCycleNumber
'   10          DataItemCode            Text 15     Mandatory
'   11          RepeatNumber            Integer     Mandatory
'   12          Response                Text 255    Mandatory (NULL response is now valid)
'   13          Unobtainable            Integer     Optional (this element can be missing)
'   14          UserName                Text 20     Optional (this element can be missing)
'
'Changed Mo 7/5/2003, a NULL response is now a valid Response
'Mo 30/6/2008 - WO-080002
'Unobtainable has been inserted as an optional element between Response and UserName.
'Note that Unobtainable is an optional field.
'If a line contains 13 elements (nos. 0 to 12) it will be deemed to have no Unobtainable and no UserName.
'If a line contains 14 elements (nos. 0 to 13) the contents of the last element will be inspected.
'If its blank, 0 or 1 it will be deemed to be an Unobtainable field.
'If its not blank, 0 or 1 it will be deemed to be a UserName and will be validated as such.
'If a line contains 15 element (nos. 0 to 14) it will be deemed to have Unobtainable (element no. 13)
'and UserName (element no. 14)
'---------------------------------------------------------------------
Dim asBRFArray() As String
Dim nNumberOfElements As Integer
Dim bNoErrorsOnThisLine As Boolean
Dim lClinicalTrialId As Long
Dim lVisitId As Long
Dim lCRFPageId As Long
Dim lDataItemId As Long

    On Error GoTo ErrLabel

    bNoErrorsOnThisLine = True
    
    lClinicalTrialId = 0
    lVisitId = 0
    lCRFPageId = 0
    lDataItemId = 0
    asBRFArray = SplitCSV(sBRFLine)
    nNumberOfElements = UBound(asBRFArray)
    'Mo 30/6/2008 - WO-080002
    'check that each line contains 13, 14 or 15 elements
    'Note that nNumberOfElements is always 1 less than the actual number of elements
    'e.g. "one,two,three,four,five" contains 5 elements, but the Ubound would be 4.
    If (nNumberOfElements <> 12) And (nNumberOfElements <> 13) And (nNumberOfElements <> 14) Then
        Call AddLineErrMsg(sErrors, lLineCount, " contains the wrong number of elements.")
        bNoErrorsOnThisLine = False
    End If
    If bNoErrorsOnThisLine Then
        'Validate element 0, ClinicalTrialName
        'If ClinicalTrialName does not exist lClinicalTrialId will be set to 0
        If Not ClinicalTrialNameExists(asBRFArray(0), lClinicalTrialId) Then
            Call AddLineErrMsg(sErrors, lLineCount, " contains the unknown Study Name " & asBRFArray(0) & ".")
            bNoErrorsOnThisLine = False
        End If
        'Validate element 1, Site
        If lClinicalTrialId <> 0 Then
            If Not SiteExists(asBRFArray(1), lClinicalTrialId) Then
                Call AddLineErrMsg(sErrors, lLineCount, " contains the invalid Site code " & asBRFArray(1) & ".")
                bNoErrorsOnThisLine = False
            End If
        End If
        'Check element 2 (PersonId) and 3 (SubjectLabel) are not both blank
        If (asBRFArray(2) = "") And (asBRFArray(3) = "") Then
            Call AddLineErrMsg(sErrors, lLineCount, " contains neither a Subject Id nor a Subject Label" & ".")
            bNoErrorsOnThisLine = False
        End If
        'Check element 2 (PersonId) and 3 (SubjectLabel) do not both exist
        If (asBRFArray(2) <> "") And (asBRFArray(3) <> "") Then
            Call AddLineErrMsg(sErrors, lLineCount, " contains both a Subject Id and a Subject Label" & ".")
            bNoErrorsOnThisLine = False
        End If
        'Validate element 2, PersonId
        If (asBRFArray(2) <> "") Then
            If ((Not gblnValidString(asBRFArray(2), valNumeric) Or Len(asBRFArray(2)) > 9)) Then
                Call AddLineErrMsg(sErrors, lLineCount, " contains the invalid Subject Id " & asBRFArray(2) & ".")
                bNoErrorsOnThisLine = False
            End If
        End If
        'Validate element 3, SubjectLabel
        If (asBRFArray(3) <> "") Then
            If Not gblnValidString(asBRFArray(3), valOnlySingleQuotes) Then
                Call AddLineErrMsg(sErrors, lLineCount, " contains an invalid Subject Label that " & gsCANNOT_CONTAIN_INVALID_CHARS)
                bNoErrorsOnThisLine = False
            End If
        End If
        'Validate element 4, VisitCode
        If lClinicalTrialId <> 0 Then
            If Not VisitCodeExists(asBRFArray(4), lClinicalTrialId, lVisitId) Then
                Call AddLineErrMsg(sErrors, lLineCount, " contains the unknown Visit Code " & asBRFArray(4) & ".")
                bNoErrorsOnThisLine = False
            End If
        End If
        'Check element 5 (VisitCycleNumber) and 6 (VisitCycleDate) are not both blank
        If (asBRFArray(5) = "") And (asBRFArray(6) = "") Then
            Call AddLineErrMsg(sErrors, lLineCount, " contains neither a Visit Cycle Number nor a Visit Cycle Date" & ".")
            bNoErrorsOnThisLine = False
        End If
        'Check element 5 (VisitCycleNumber) and 6 (VisitCycleDate) do not both exist
        If (asBRFArray(5) <> "") And (asBRFArray(6) <> "") Then
            Call AddLineErrMsg(sErrors, lLineCount, " contains both a Visit Cycle Number and a Visit Cycle Date" & ".")
            bNoErrorsOnThisLine = False
        End If
        'Validate element 5, VisitCycleNumber
        If (asBRFArray(5) <> "") Then
            If ((Not gblnValidString(asBRFArray(5), valNumeric) Or Len(asBRFArray(5)) > 9)) Then
                Call AddLineErrMsg(sErrors, lLineCount, " contains the invalid Visit Cycle Number " & asBRFArray(5) & ".")
                bNoErrorsOnThisLine = False
            End If
        End If
        'Validate element 6, VisitCycleDate
        If (asBRFArray(6) <> "") Then
            If ValidateDate(asBRFArray(6), gdblSelEFormDate) > "" Then
                Call AddLineErrMsg(sErrors, lLineCount, " contains the invalid Visit Cycle Date " & asBRFArray(6) & ".")
                bNoErrorsOnThisLine = False
            End If
        End If
        'Validate element 7, CRFPageCode
        If (lClinicalTrialId <> 0) And (lVisitId <> 0) Then
            If Not eFormCodeExists(asBRFArray(7), lClinicalTrialId, lVisitId, lCRFPageId) Then
                Call AddLineErrMsg(sErrors, lLineCount, " contains the unknown eForm Code " & asBRFArray(7) & ".")
                bNoErrorsOnThisLine = False
            End If
        End If
        'Check element 8 (CRFPageCycleNumber) and 9 (CRFPageCycleDate) are not both blank
        If (asBRFArray(8) = "") And (asBRFArray(9) = "") Then
            Call AddLineErrMsg(sErrors, lLineCount, " contains neither an eForm Cycle Number nor a valid eForm Cycle Date" & ".")
            bNoErrorsOnThisLine = False
        End If
        'Check element 8 (CRFPageCycleNumber) and 9 (CRFPageCycleDate) do not both exist
        If (asBRFArray(8) <> "") And (asBRFArray(9) <> "") Then
            Call AddLineErrMsg(sErrors, lLineCount, " contains both an eForm Cycle Number and an invalid eForm Cycle Date" & ".")
            bNoErrorsOnThisLine = False
        End If
        'Validate element 8, CRFPageCycleNumber
        If (asBRFArray(8) <> "") Then
            If ((Not gblnValidString(asBRFArray(8), valNumeric) Or Len(asBRFArray(8)) > 9)) Then
                Call AddLineErrMsg(sErrors, lLineCount, " contains the invalid eForm Cycle Number " & asBRFArray(8) & ".")
                bNoErrorsOnThisLine = False
            End If
        End If
        'Validate element 9, CRFPageCycleDate
        If (asBRFArray(9) <> "") Then
            If ValidateDate(asBRFArray(9), gdblSelEFormDate) > "" Then
                Call AddLineErrMsg(sErrors, lLineCount, " contains the invalid eForm Cycle Date " & asBRFArray(9) & ".")
                bNoErrorsOnThisLine = False
            End If
        End If
        'Validate element 10, DataItemCode
        If (lClinicalTrialId <> 0) And (lCRFPageId <> 0) Then
            If Not DataItemCodeExists(asBRFArray(10), lClinicalTrialId, lCRFPageId, lDataItemId) Then
                Call AddLineErrMsg(sErrors, lLineCount, " contains the unknown Question Code " & asBRFArray(10) & ".")
                bNoErrorsOnThisLine = False
            End If
        End If
        'Validate element 11, RepeatNumber
        If ((Not gblnValidString(asBRFArray(11), valNumeric) Or Len(asBRFArray(11)) > 9)) Then
            Call AddLineErrMsg(sErrors, lLineCount, " contains the invalid Repeat Number " & asBRFArray(11) & ".")
            bNoErrorsOnThisLine = False
        End If
        'Validate element 12, Response
        If lDataItemId <> 0 Then
            'Check for a non empty string, an empty string represents a NULL response and is valid
            If asBRFArray(12) <> "" Then
                'Check DataItemType for a category question
                If DataTypeFromId(lClinicalTrialId, lDataItemId) = DataType.Category Then
                    'Validate the Response as a category code
                    If Not ValidCategoryResponse(asBRFArray(12), lClinicalTrialId, lDataItemId) Then
                        Call AddLineErrMsg(sErrors, lLineCount, " contains the invalid Category Response  " & asBRFArray(12) & ".")
                        bNoErrorsOnThisLine = False
                    End If
                Else
                    'Validate the Response as free text
                    If Not gblnValidString(asBRFArray(12), valOnlySingleQuotes) Then
                        Call AddLineErrMsg(sErrors, lLineCount, " contains an invalid Response that " & gsCANNOT_CONTAIN_INVALID_CHARS)
                        bNoErrorsOnThisLine = False
                    End If
                End If
            End If
        End If
        'Mo 30/6/2008 - WO-080002
        'The following section is now based on the number of elements.
        
        'If nNumberOfElements = 12 (nos. 0 to 12) then
        '   Unobtainable and UserName do not exist, no further validation required
        
        'If nNumberOfElements = 13 (nos. 0 to 13) then
        '   element 13 is either Unobtainable or UserName
        '   If it is blank, 0 or 1 then it will be treated as Unobtainable with no UserName
        '   If it is not blank, 0 or 1 it will be validated as UserName
        If nNumberOfElements = 13 Then
            If asBRFArray(13) <> "" And asBRFArray(13) <> "0" And asBRFArray(13) <> "1" Then
                'A UserName exists, validate it
                If Not UserNameExists(asBRFArray(13)) Then
                    Call AddLineErrMsg(sErrors, lLineCount, " contains the unknown User Name " _
                        & asBRFArray(13) & ".")
                    bNoErrorsOnThisLine = False
                Else
                    'Check UserName is enabled
                    If Not UserNameEnabled(asBRFArray(13)) Then
                        Call AddLineErrMsg(sErrors, lLineCount, " contains the disabled User Name " _
                            & asBRFArray(13) & ".")
                        bNoErrorsOnThisLine = False
                    Else
                        'Check UserName has permissions for currently opened database
                        If Not UserNameHasDBRights(asBRFArray(13)) Then
                            Call AddLineErrMsg(sErrors, lLineCount, " contains the User Name " _
                                & asBRFArray(13) & ", which does not have access rights for the currently open database" & ".")
                            bNoErrorsOnThisLine = False
                        End If
                    End If
                End If
            Else
                'if Unobtainable is set check that it is accompanied by a blank response
                If asBRFArray(13) = "1" And (asBRFArray(12) <> "") Then
                    Call AddLineErrMsg(sErrors, lLineCount, " Unobtainable with a non-blank response of " _
                        & asBRFArray(12) & " not allowed" & ".")
                    bNoErrorsOnThisLine = False
                End If
            End If
        End If
        
        'If nNumberOfElements = 14 (nos. 0 to 14) then
        '   element 13 will be validated as Unobtainable (blank, 0 or 1)
        '   element 14 will be validated as UserName
        If nNumberOfElements = 14 Then
            If asBRFArray(13) <> "" And asBRFArray(13) <> "0" And asBRFArray(13) <> "1" Then
                Call AddLineErrMsg(sErrors, lLineCount, " contains the invalid Unobtainable setting of " _
                    & asBRFArray(13) & ".")
                bNoErrorsOnThisLine = False
            End If
            'if Unobtainable is set check that it is accompanied by a blank response
            If asBRFArray(13) = "1" And (asBRFArray(12) <> "") Then
                Call AddLineErrMsg(sErrors, lLineCount, " Unobtainable with a non-blank response of " _
                    & asBRFArray(12) & " not allowed" & ".")
                bNoErrorsOnThisLine = False
            End If
            If asBRFArray(14) <> "" Then
                'A UserName exists, validate it
                If Not UserNameExists(asBRFArray(14)) Then
                    Call AddLineErrMsg(sErrors, lLineCount, " contains the unknown User Name " _
                        & asBRFArray(14) & ".")
                    bNoErrorsOnThisLine = False
                Else
                    'Check UserName is enabled
                    If Not UserNameEnabled(asBRFArray(14)) Then
                        Call AddLineErrMsg(sErrors, lLineCount, " contains the disabled User Name " _
                            & asBRFArray(14) & ".")
                        bNoErrorsOnThisLine = False
                    Else
                        'Check UserName has permissions for currently opened database
                        If Not UserNameHasDBRights(asBRFArray(14)) Then
                            Call AddLineErrMsg(sErrors, lLineCount, " contains the User Name " & _
                                asBRFArray(14) & ", which does not have access rights for the currently open database" & ".")
                            bNoErrorsOnThisLine = False
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    ' Return the result for this line
    ValidBatchResponseLine = bNoErrorsOnThisLine

    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modBatchDataEntry.ValidBatchResponseLine"

End Function

'---------------------------------------------------------------------
Private Sub AddLineErrMsg(ByRef sErrors As String, ByVal lLineCount As Long, ByVal sMSG As String)
'---------------------------------------------------------------------
' Add a line error message to the error message string in sErrors
'---------------------------------------------------------------------

    sErrors = sErrors & "Line no. " & lLineCount & sMSG & vbNewLine

End Sub

'---------------------------------------------------------------------
Private Function ClinicalTrialNameExists(ByVal sClinicalTrialName As String, _
                                        ByRef lClinicalTrialId As Long) As Boolean
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bExists As Boolean

    On Error GoTo ErrHandler

    sSQL = "SELECT ClinicalTrialId, ClinicalTrialName  FROM ClinicalTrial " _
        & "WHERE ClinicalTrialName = '" & sClinicalTrialName & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        bExists = False
        lClinicalTrialId = 0
    Else
        bExists = True
        lClinicalTrialId = rsTemp!ClinicalTrialId
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    ClinicalTrialNameExists = bExists
        
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ClinicalTrialNameExists", "modBatchDataEntry")
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
Private Function SiteExists(ByVal sSite As String, _
                            ByVal lClinicalTrialId As Long) As Boolean
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bExists As Boolean

    On Error GoTo ErrHandler

    'Mo 21/6/2005 bug 2538, Prevent site codes containing upper case characters
    'getting into the database. Check the site codes for lowercase characters
    If sSite <> LCase(sSite) Then
        bExists = False
        SiteExists = bExists
        Exit Function
    End If
    
    sSQL = "SELECT TrialSite FROM TrialSite " _
        & "WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND TrialSite = '" & sSite & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        bExists = False
    Else
        bExists = True
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    SiteExists = bExists
        
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "SiteExists", "modBatchDataEntry")
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
Private Function VisitCodeExists(ByVal sVisitCode As String, _
                                ByVal lClinicalTrialId As Long, _
                                ByRef lVisitId As Long) As Boolean
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bExists As Boolean

    On Error GoTo ErrHandler

    sSQL = "SELECT VisitCode, VisitId FROM StudyVisit " _
        & "WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND VisitCode = '" & sVisitCode & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        bExists = False
        lVisitId = 0
    Else
        bExists = True
        lVisitId = rsTemp!VisitId
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    VisitCodeExists = bExists
        
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "VisitCodeExists", "modBatchDataEntry")
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
Private Function eFormCodeExists(ByVal sCRFPageCode As String, _
                                ByVal lClinicalTrialId As Long, _
                                ByVal lVisitId As Long, _
                                ByRef lCRFPageId As Long) As Boolean
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bExists As Boolean

    On Error GoTo ErrHandler

    sSQL = "SELECT CRFPage.CRFPageCode, CRFPage.CRFPageId FROM StudyVisitCRFPage, CRFPage " _
        & "WHERE StudyVisitCRFPage.ClinicalTrialId = CRFPage.ClinicalTrialId" _
        & " AND StudyVisitCRFPage.CRFPageId = CRFPage.CRFPageId" _
        & " AND StudyVisitCRFPage.ClinicalTrialId = " & lClinicalTrialId _
        & " AND StudyVisitCRFPage.VisitId = " & lVisitId _
        & " AND CRFPage.CRFPageCode = '" & sCRFPageCode & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        bExists = False
        lCRFPageId = 0
    Else
        bExists = True
        lCRFPageId = rsTemp!CRFPageId
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    eFormCodeExists = bExists
        
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "eFormCodeExists", "modBatchDataEntry")
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
Private Function DataItemCodeExists(ByVal sDataItemCode As String, _
                                    ByVal lClinicalTrialId As Long, _
                                    ByVal lCRFPageId As Long, _
                                    ByRef lDataItemId As Long) As Boolean
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bExists As Boolean

    On Error GoTo ErrHandler

    sSQL = "SELECT DataItem.DataItemCode, DataItem.DataItemId FROM DataItem, CRFElement " _
        & "WHERE DataItem.ClinicalTrialId = CRFElement.ClinicalTrialId " _
        & "AND DataItem.DataItemId = CRFElement.DataItemId " _
        & " AND DataItem.ClinicalTrialId = " & lClinicalTrialId _
        & " AND CRFElement.CRFPageId = " & lCRFPageId _
        & " AND DataItem.DataItemCode = '" & sDataItemCode & "'"

    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        bExists = False
        lDataItemId = 0
    Else
        bExists = True
        lDataItemId = rsTemp!DataItemId
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    DataItemCodeExists = bExists
        
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "DataItemCodeExists", "modBatchDataEntry")
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
Public Function ValidateDate(ByVal sInDate As String, _
                            ByRef dblDate As Double) As String
'---------------------------------------------------------------------
' Validate text date value
' Returns empty string if all OK, otherwise returns error message
' dblDate is set to the date as a double (if valid)
' Assume sDate > ""
'---------------------------------------------------------------------
Dim sMSG As String
Dim sArezzoDate As String
Dim sDate As String

    On Error GoTo ErrHandler

    sMSG = ""
    dblDate = 0
    ' Read date using current default date format
    sDate = goArezzo.ReadValidDate(sInDate, "dd/mm/yyyy", sArezzoDate)
    If sDate = "" Then
        ' empty string, therefore invalid
        sMSG = sInDate & " is not recognised as a valid date." & vbCrLf
        sMSG = sMSG & "Please enter the date in the format dd/mm/yyyy"
    Else
        ' It was a valid date - check it's reasonable
        dblDate = goArezzo.ArezzoDateToDouble(sArezzoDate)
        If dblDate <= 0 Then
            ' It's too far in the past
            sMSG = sDate & " is not accepted as a valid date." & vbCrLf
            sMSG = sMSG & "The date must not be before 1900"
            dblDate = 0
        ElseIf dblDate > CDbl(Now) Then
            ' It's in the future
            sMSG = sDate & " is not accepted as a valid date." & vbCrLf
            sMSG = sMSG & "The date must not be in the future."
            dblDate = 0
        Else
            ' It was OK
        End If
    End If
        
    ValidateDate = sMSG

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ValidateDate", "modBatchDataEntry")
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
Private Function ValidCategoryResponse(ByVal sResponse As String, _
                                        ByVal lClinicalTrialId As Long, _
                                        ByVal lDataItemId As Long) As Boolean
'---------------------------------------------------------------------
'Note that either a ValueCode or an ItemValue is an acceptable response
'to a category question
' NCJ 14 July 04 - Do non-case-sensitive matching
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bFound As Boolean
Dim sLCResponse As String

    On Error GoTo ErrHandler

    'initialize the bFound
    bFound = False
    
    sLCResponse = LCase(sResponse)
    
    sSQL = "SELECT ValueCode, ItemValue FROM ValueData " _
        & "WHERE ClinicalTrialId = " & lClinicalTrialId & " " _
        & "AND DataItemId = " & lDataItemId & " " _
        & "ORDER BY ValueOrder"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Do Until rsTemp.EOF Or bFound
        ' NCJ 14 Jul 04 - Case-insensitive comparison
        If LCase(rsTemp!ValueCode) = sLCResponse Then
            bFound = True
        End If
        If LCase(rsTemp!ItemValue) = sLCResponse Then
            bFound = True
        End If
        rsTemp.MoveNext
    Loop
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    ValidCategoryResponse = bFound

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ValidCategoryResponse", "modBatchDataEntry")
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
Private Function UserNameExists(ByVal sUserName As String) As Boolean
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bExists As Boolean

    On Error GoTo ErrHandler

    sSQL = "SELECT UserName FROM MACROUser " _
        & "WHERE UserName = '" & sUserName & "'"

    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        bExists = False
    Else
        bExists = True
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    UserNameExists = bExists
        
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UserNameExists", "modBatchDataEntry")
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
Private Function UserNameEnabled(ByVal sUserName As String) As Boolean
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bEnabled As Boolean

    On Error GoTo ErrHandler

    sSQL = "SELECT Enabled FROM MACROUser " _
        & "WHERE UserName = '" & sUserName & "'"

    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsTemp!Enabled = 0 Then
        bEnabled = False
    Else
        bEnabled = True
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    UserNameEnabled = bEnabled
        
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UserNameEnabled", "modBatchDataEntry")
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
Private Function UserNameHasDBRights(ByVal sUserName As String) As Boolean
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bExists As Boolean

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM UserDatabase " _
        & "WHERE UserName = '" & sUserName & "'" _
        & " AND DatabaseCode = '" & goUser.DatabaseCode & "'"

    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        bExists = False
    Else
        bExists = True
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    UserNameHasDBRights = bExists
        
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UserNameHasDBRights", "modBatchDataEntry")
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
Public Sub ImportBatchResponseFile(ByVal sBatchResponseFile As String)
'---------------------------------------------------------------------
'This sub imports the contents of a Batch Response File that has previously
'been validated by ValidBatchResponseFile.
'The sub reads the import file a line at a time and places the contents
'in table BatchResponseData and the Response Buffer (lvwBuffer)
'---------------------------------------------------------------------
Dim nIOFileNumber As Integer
Dim sBRFLine As String
Dim itmX As MSComctlLib.ListItem
Dim lLineCount As Long
Dim i As Integer

    On Error GoTo ErrHandler

    Call HourglassOn
    
    'open the Batch Response File
    nIOFileNumber = FreeFile
    Open sBatchResponseFile For Input As #nIOFileNumber
    
    lLineCount = 0
    'Read the Batch Response File line by line
    Do While Not EOF(nIOFileNumber)
        Line Input #nIOFileNumber, sBRFLine
        lLineCount = lLineCount + 1
        frmMenu.txtProgress.Text = "Importing " & lLineCount & " of " & mlImportCount
        frmMenu.txtProgress.Refresh
        ' NCJ 29 Jun 04 - Call new routine, and only if line isn't blank
        If Trim(sBRFLine) > "" Then
            Call ImportBatchResponseLine(sBRFLine, lLineCount)
        End If
    Loop
    
    'close the input file
    Close #nIOFileNumber

    'Set the Max Column Widths
    'Mo 30/6/2008 - WO-080002
    For i = 2 To 13
        Call lvw_SetColWidth(frmMenu.lvwBuffer, i, LVSCW_AUTOSIZE_USEHEADER)
    Next i
    
    frmMenu.txtProgress.Text = "Import completed"
    frmMenu.txtProgress.Refresh
    
    Call HourglassOff

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ImportBatchResponseFile", "modBatchDataEntry")
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
Private Sub ImportBatchResponseLine(ByVal sBRFLine As String, ByVal lLineCount As Long)
'---------------------------------------------------------------------
' NCJ 29 Jun 04 - Created from code copied from ImportBatchResponseFile
'This sub imports a line of a Batch Response File that has previously
'been validated by ValidBatchResponseFile.
'The sub places the contents
'in table BatchResponseData and the Response Buffer (lvwBuffer)
'Mo 30/6/2008 - WO-080002
'---------------------------------------------------------------------
Dim sSQL As String
Dim asBRFArray() As String
Dim nNumberOfElements As Integer
Dim lNextBatchId As Long
Dim lClinicalTrialId As Long
Dim lPersonId As Long
Dim sSubjectLabel As String
Dim lVisitId As Long
Dim lVisitCycleNumber As Long
Dim dblVisitCycleDate As Double
Dim lCRFPageId As Long
Dim lCRFPageCycleNumber As Long
Dim dblCRFPageCycleDate As Double
Dim lDataItemId As Long
Dim sUserName As String
Dim itmX As MSComctlLib.ListItem
Dim sVisitCycle As String
Dim seFormCycle As String
Dim sResponse As String
Dim lDataType As Long
Dim lRepeatNumber As Long
'Mo 30/6/2008 - WO-080002
Dim nUnobtainable As Integer

    On Error GoTo ErrLabel

    asBRFArray = SplitCSV(sBRFLine)
    nNumberOfElements = UBound(asBRFArray)
    
    'Get next available BatchResponseId
    lNextBatchId = GetNextBatchId
    
    'Prepare the fields for inserting into table BatchResponseData
    lClinicalTrialId = TrialIdFromName(asBRFArray(0))
    'Perpare PersonId and SubjectLabel
    If asBRFArray(2) = "" Then
        lPersonId = 0
    Else
        lPersonId = CLng(asBRFArray(2))
    End If
    If asBRFArray(3) = "" Then
        sSubjectLabel = "Null"
    Else
        sSubjectLabel = "'" & asBRFArray(3) & "'"
    End If
    lVisitId = VisitIdFromCode(lClinicalTrialId, asBRFArray(4))
    lCRFPageId = CRFPageIdFromCode(lClinicalTrialId, asBRFArray(7))
    lDataItemId = DataItemIdFromCode(lClinicalTrialId, asBRFArray(10))
    
    'Prepare VisitCycleNumber and VisitCycleDate
    If asBRFArray(5) = "" Then
        Call ValidateDate(asBRFArray(6), dblVisitCycleDate)
        lVisitCycleNumber = 0
    Else
        lVisitCycleNumber = CLng(asBRFArray(5))
        dblVisitCycleDate = 0
    End If
    'Prepare CRFPageCycleNumber and CRFPageCycleDate
    If asBRFArray(8) = "" Then
        Call ValidateDate(asBRFArray(9), dblCRFPageCycleDate)
        lCRFPageCycleNumber = 0
    Else
        lCRFPageCycleNumber = CLng(asBRFArray(8))
        dblCRFPageCycleDate = 0
    End If
    'Prepare RepeatNumber
    lRepeatNumber = asBRFArray(11)
    'if its a numeric question run ConvertLocalNumToStandard over the response
    lDataType = DataTypeFromId(lClinicalTrialId, lDataItemId)
    Select Case lDataType
    Case DataType.IntegerData, DataType.LabTest, DataType.Real
        sResponse = ConvertLocalNumToStandard(asBRFArray(12))
    Case Else
        sResponse = asBRFArray(12)
    End Select
    
    'Mo 30/6/2008 - WO-080002
    'Handle Unobtainable and UserName fields
    If nNumberOfElements = 12 Then
        'Unobtainable not present so set it to 0
        nUnobtainable = 0
        'UserName not present so set to currently logged in user
        sUserName = goUser.UserName
    End If
    If nNumberOfElements = 13 Then
        'element 13 is either Unobtainable or UserName
        'If it is blank, 0 or 1 then it is treated as Unobtainable with no UserName
        If asBRFArray(13) = "" Or asBRFArray(13) = "0" Or asBRFArray(13) = "1" Then
            If asBRFArray(13) = "1" Then
                nUnobtainable = 1
            Else
                'note that an Unobtainable of blank is set to 0
                nUnobtainable = 0
            End If
            'no UserName supplied so set to currently logged in user
            sUserName = goUser.UserName
        Else
            'element 13 must be a UserName
            sUserName = asBRFArray(13)
            'Unobtainable not present so set it to 0
            nUnobtainable = 0
        End If
    End If
    If nNumberOfElements = 14 Then
        'treat element asBRFArray(13) as Unobtainable
        'valid Unobtainable values blank, 0 or 1
        If asBRFArray(13) = "1" Then
            nUnobtainable = 1
        Else
            'note that an Unobtainable of blank is set to 0
            nUnobtainable = 0
        End If
        'treat element asBRFArray(14) as UserName, it might still be empty
        If asBRFArray(14) <> "" Then
            sUserName = asBRFArray(14)
        Else
            sUserName = goUser.UserName
        End If
    End If
    
    'Add new entry to table BatchResponseData
    'Mo 30/6/2008 - WO-080002, Unobtainable added to sql
    sSQL = "INSERT INTO BatchResponseData (BatchResponseId, ClinicalTrialId, Site, PersonId, SubjectLabel, " _
        & "VisitId, VisitCycleNumber, VisitCycleDate, CRFPageID, CRFPageCycleNumber, CRFPageCycleDate, " _
        & "DataItemId, RepeatNumber, Response, UserName, Unobtainable) " _
        & "VALUES (" & lNextBatchId & "," & lClinicalTrialId & ",'" & asBRFArray(1) & "'," & lPersonId & "," & sSubjectLabel & "," _
        & lVisitId & "," & lVisitCycleNumber & "," & dblVisitCycleDate & "," & lCRFPageId & "," & lCRFPageCycleNumber & "," & dblCRFPageCycleDate & "," _
        & lDataItemId & "," & lRepeatNumber & ",'" & ReplaceQuotes(sResponse) & "','" & sUserName & "'," & nUnobtainable & ")"
    MacroADODBConnection.Execute sSQL
    
    'Increment txtBufferCount
    frmMenu.txtBufferCount.Text = frmMenu.txtBufferCount.Text + 1
    frmMenu.txtBufferCount.Refresh
    
    If frmMenu.chkDisplay.Value = 1 Then
        'Add new entry to Response Buffer (lvwBuffer)
        Set itmX = frmMenu.lvwBuffer.ListItems.Add(, , lNextBatchId)
        itmX.SubItems(1) = asBRFArray(0)
        itmX.SubItems(2) = asBRFArray(1)
        'Decide between PersonId or Subject Label
        If asBRFArray(2) = "" Then
            itmX.SubItems(3) = asBRFArray(3)
            'Mo 30/6/2008 - WO-080002
            itmX.SubItems(14) = 0
        Else
            itmX.SubItems(3) = asBRFArray(2)
            'Mo 30/6/2008 - WO-080002
            itmX.SubItems(14) = 1
        End If
        itmX.SubItems(4) = asBRFArray(4)
        'Decide between Visit Cycle Number or visit Cycle Date
        If asBRFArray(5) = "" Then
            sVisitCycle = asBRFArray(6)
            'Mo 30/6/2008 - WO-080002
            itmX.SubItems(15) = 0
        Else
            sVisitCycle = asBRFArray(5)
            'Mo 30/6/2008 - WO-080002
            itmX.SubItems(15) = 1
        End If
        itmX.SubItems(5) = sVisitCycle
        itmX.SubItems(6) = asBRFArray(7)
        'Decide between eForm Cycle Number or eForm Cycle Date
        If asBRFArray(8) = "" Then
            seFormCycle = asBRFArray(9)
            'Mo 30/6/2008 - WO-080002
            itmX.SubItems(16) = 0
        Else
            seFormCycle = asBRFArray(8)
            'Mo 30/6/2008 - WO-080002
            itmX.SubItems(16) = 1
        End If
        itmX.SubItems(7) = seFormCycle
        itmX.SubItems(8) = asBRFArray(10)
        itmX.SubItems(9) = lRepeatNumber
        itmX.SubItems(10) = sResponse
        'Mo 30/6/2008 - WO-080002
        If nUnobtainable = 1 Then
            itmX.SubItems(11) = 1
        End If
        itmX.SubItems(12) = sUserName
        
        'Make sure last entry is visible
        frmMenu.lvwBuffer.ListItems(itmX.Index).EnsureVisible
    End If

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modBatchDataEntry.ImportBatchResponseLine"

End Sub

'---------------------------------------------------------------------
Public Function GetNextBatchId() As Long
'---------------------------------------------------------------------
' Returns next available BatchResponsId for table BatchResponseData
'---------------------------------------------------------------------
Dim rsNextBatchId As ADODB.Recordset
Dim sSQL As String
Dim lNewBatchId As Long
    
    On Error GoTo ErrLabel
        
    sSQL = " Select MAX(BatchResponseId) as MaxBatchId FROM BatchResponseData"
    Set rsNextBatchId = New ADODB.Recordset
    rsNextBatchId.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    If IsNull(rsNextBatchId!MaxBatchId) Then
        lNewBatchId = 1
    Else
        lNewBatchId = rsNextBatchId!MaxBatchId + 1
    End If
    rsNextBatchId.Close
    Set rsNextBatchId = Nothing
    
    GetNextBatchId = lNewBatchId
        
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modBatchDataEntry.GetNextBatchId"

End Function

'---------------------------------------------------------------------
Public Sub UploadBatchResponses()
'---------------------------------------------------------------------
'revisions
'TA 20/04/2004: remove null on the response value becasue code can't handle nulls
' NCJ 30 Jun 04 - Use module-level StudyDef
' NCJ 27 Sept 07 - Bugs 2935 and 2937 - Sort out user name/user name full
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsBRD  As ADODB.Recordset
Dim bStudyExists As Boolean
Dim bSubjectloaded As Boolean
Dim lClinicalTrialId As Long
Dim lPrevClinicalTrialId As Long
Dim sSiteIdLabel As String
Dim sPrevSiteIdLabel As String
Dim lPersonId As Long
Dim bOKToContinue As Boolean
Dim oQuestion As eFormElementRO
Dim oSubject As StudySubject
Dim oVisitInstance As VisitInstance
Dim oeFormRO As eFormRO
Dim oeFormInstance As EFormInstance
Dim oResponse As Response
Dim sLoadMessage As String
Dim bChanged As Boolean
Dim nStatus As Integer
Dim sErrMsg As String
Dim sRFC As String
Dim bUserHasRights As Boolean
Dim sUserName As String
Dim sUserMessage As String
Dim lBatchId As Long
Dim sThisFormDetails As String
Dim sFormLoadDetails As String
Dim lTotal As Long
Dim lCount As Long
'Mo 14/2/2007 Bug 2876
'Dim sRoleCode As String
Dim nLoadResponsesResult As eLoadResponsesResult
Dim sEFILockToken As String
Dim sVEFILockToken As String
Dim sLockErrMsg As String
Dim alBatchIdsForDeletion() As Long
Dim nBatchDeleteNumber As Integer
Dim sSite As String
Dim sUserNameFull As String     ' NCJ 27 Sept 07
Dim oCodedTermHistory As MACROCCBS30.CodedTermHistory
Dim oTimeZone As MACROTimeZoneBS30.Timezone

    On Error GoTo ErrHandler
    
    If Not LockBatchUpload Then Exit Sub

    Call HourglassOn

    'Get all the response from the BatchResponseData table and sort it
    sSQL = "SELECT * FROM BatchResponseData, CRFElement " _
        & "WHERE BatchResponseData.ClinicalTrialId = CRFElement.ClinicalTrialId " _
        & "AND BatchResponseData.CRFPageId = CRFElement.CRFPageId " _
        & "AND BatchResponseData.DataItemId = CRFElement.DataItemId " _
        & "ORDER BY BatchResponseData.ClinicalTrialId, Site, PersonId, SubjectLabel, " _
        & "VisitId, VisitCycleNumber, VisitCycleDate, " _
        & "BatchResponseData.CRFPageId, CRFPageCycleNumber, CRFPageCycleDate, " _
        & "RepeatNumber, CRFElement.FieldOrder, CRFElement.QGroupFieldOrder"
    Set rsBRD = New ADODB.Recordset
    rsBRD.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    lTotal = rsBRD.RecordCount
    
    Set oTimeZone = New MACROTimeZoneBS30.Timezone
    
    lCount = 0
    lPrevClinicalTrialId = 0
    sPrevSiteIdLabel = ""
    sFormLoadDetails = ""
    sUserName = ""
    sUserNameFull = ""
    mbFormLabSet = False
    Do While Not rsBRD.EOF
        lCount = lCount + 1
        frmMenu.txtProgress.Text = "Uploading " & lCount & " of " & lTotal
        frmMenu.txtProgress.Refresh
        ' NCJ 28 Jun 04 - Store site
        sSite = rsBRD!Site
        lClinicalTrialId = rsBRD!ClinicalTrialId
        'The variables sThisFormDetails and sFormLoadDetails control the Loading and Saving of a eForms responses
        sThisFormDetails = lClinicalTrialId & "|" & sSite & "|" & rsBRD!PersonId & "|" & rsBRD!SubjectLabel & "|" _
            & rsBRD!VisitId & "|" & rsBRD!VisitCycleNumber & "|" & rsBRD!VisitCycleDate & "|" _
            & rsBRD!CRFPageId & "|" & rsBRD!CRFPageCycleNumber & "|" & rsBRD!CRFPageCycleDate
        'Check to see if current Batch Response is from the last eForm for which LoadResponses has occured
        'If not SaveResponse for the last eForm
        If (sThisFormDetails <> sFormLoadDetails) And (sFormLoadDetails <> "") Then
            sFormLoadDetails = ""
            'Save the previously loaded eForm's responses
            Call SaveFormResponses(oSubject, oeFormInstance, alBatchIdsForDeletion, sUserName)
        End If
        lBatchId = rsBRD!BatchResponseId
        'Check That the UserName has required permissions
        ' NCJ 27 Sept 07 - Issue 2937 - Also get UserNameFull
        If rsBRD!UserName <> sUserName Then
            If UserCanChangeData(rsBRD!UserName, sUserNameFull, sUserMessage) Then
                bUserHasRights = True
            Else
                bUserHasRights = False
            End If
            sUserName = rsBRD!UserName
        End If
        If Not bUserHasRights Then
            Call SetUploadMessage(lBatchId, sUserMessage)
        Else
            'Check for a change of study
            If lClinicalTrialId <> lPrevClinicalTrialId Then
                'check that the study still exists in the current database
                If TrialNameFromId(lClinicalTrialId) = "" Then
                    bStudyExists = False
                Else
                    'The study exists, Load it
                    bStudyExists = True
                    ' NCJ 30 Jun 04 - Create the study def object here
                    Set moStudyDef = New StudyDefRO
                    moStudyDef.Load gsADOConnectString, lClinicalTrialId, 1, goArezzo
                End If
                lPrevClinicalTrialId = lClinicalTrialId
                'Clear sPrevSiteIdLabel when the study has changed
                sPrevSiteIdLabel = ""
            End If
            If Not bStudyExists Then
                Call SetUploadMessage(lBatchId, "Unknown study")
            Else
                'Validate the PersonId/Subject Label by calling CheckPersonIdLabel, which returns 0 if both are invalid
                lPersonId = CheckPersonIdLabel(lClinicalTrialId, sSite, rsBRD!PersonId, rsBRD!SubjectLabel, lBatchId)
                'Proceed if we have a valid/existing PersonId
                If lPersonId > 0 Then
                    sSiteIdLabel = sSite & lPersonId
                    'Check for a different subject
                    If sSiteIdLabel <> sPrevSiteIdLabel Then
                        'Load the Subject
                        ' NCJ 6 May 03 - Added UserNameFull and UserRole
                        ' NCJ 27 Sept 07 - Changed sUserNameFull
                        ' NB UserRole now doesn't match!!!
                        Set oSubject = moStudyDef.LoadSubject(sSite, lPersonId, sUserName, Read_Write, _
                                        sUserNameFull, goUser.UserRole, True)
                        'Check that the subject has loaded correctly
                        If moStudyDef.Subject.CouldNotLoad Then
                            sLoadMessage = moStudyDef.Subject.CouldNotLoadReason
                            bSubjectloaded = False
                            sPrevSiteIdLabel = ""
                        Else
                            bSubjectloaded = True
                            sPrevSiteIdLabel = sSiteIdLabel
                        End If
                    End If
                    If Not bSubjectloaded Then
                        Call SetUploadMessage(lBatchId, "Subject could not be loaded - " & sLoadMessage)
                    Else
                        bOKToContinue = True
                        'Check the Visit, eForm, Visit/eForm and Question, This call creates oQuestion
                        bOKToContinue = CheckVisiteFormQuestion(rsBRD!VisitId, rsBRD!CRFPageId, rsBRD!DataItemId, oQuestion, lBatchId)
                        If bOKToContinue Then
                            'Check that the TrialSubject entry is not Locked or Frozen
                            If oSubject.LockStatus <> eLockStatus.lsUnlocked Then
                                Call SetUploadMessage(lBatchId, "Subject is Locked or Frozen")
                                bOKToContinue = False
                            End If
                        End If
                        If bOKToContinue Then
                            'Check the VisitCycleNumber or the VisitCycleDate, This call creates oVisitInstance
                            bOKToContinue = CheckVisitCycleNumberDate(rsBRD!VisitId, rsBRD!VisitCycleNumber, rsBRD!VisitCycleDate, oSubject, oVisitInstance, lBatchId)
                        End If
                        If bOKToContinue Then
                            'Check that the Visit is not locked or Frozen
                            If oVisitInstance.LockStatus <> eLockStatus.lsUnlocked Then
                                Call SetUploadMessage(lBatchId, "Visit is Locked or Frozen")
                                bOKToContinue = False
                            End If
                        End If
                        If bOKToContinue Then
                            'Check the CRFPageCycleNumber or the CRFPageCycleDate, This call creates oeFormInstance
                            bOKToContinue = CheckeFormCycleNumberDate(rsBRD!CRFPageId, rsBRD!CRFPageCycleNumber, rsBRD!CRFPageCycleDate, oSubject, oVisitInstance, oeFormRO, oeFormInstance, lBatchId)
                        End If
                        If bOKToContinue Then
                            'Check that the eForm is not locked or Frozen
                            If oeFormInstance.LockStatus <> eLockStatus.lsUnlocked Then
                                Call SetUploadMessage(lBatchId, "eForm is Locked or Frozen")
                                bOKToContinue = False
                            End If
                        End If
                        If bOKToContinue Then
                            'If its a new eForm that is being processed then sFormLoadDetails will be blank
                            If sFormLoadDetails = "" Then
                                'Load the Responses for the required eForm
                                nLoadResponsesResult = oSubject.LoadResponses(oeFormInstance, sLockErrMsg, sEFILockToken, sVEFILockToken)
                                If nLoadResponsesResult <> eLoadResponsesResult.lrrReadWrite Then
                                    Call SetUploadMessage(lBatchId, "eForm could not be loaded (" & sLockErrMsg & ")")
                                    bOKToContinue = False
                                Else
                                    ' NCJ 30 Jun 04 - Remember we haven't done the lab yet
                                    mbFormLabSet = False
                                    'Having Opened an eForm call RefreshSkipsAndDerivations
                                    ' NCJ 27 Sept 07 - Bug 2935 - Don't pass user name to RefreshSkips
'                                    Call oeFormInstance.RefreshSkipsAndDerivations(OpeningEForm, rsBRD!UserName)
                                    Call oeFormInstance.RefreshSkipsAndDerivations(OpeningEForm, "")
                                    'Having Opened a new eForm reset the Delete Controls
                                    nBatchDeleteNumber = 0
                                    ReDim alBatchIdsForDeletion(nBatchDeleteNumber)
                                    'store the details of the Loaded Form
                                    sFormLoadDetails = lClinicalTrialId & "|" & sSite & "|" & rsBRD!PersonId & "|" & rsBRD!SubjectLabel & "|" _
                                        & rsBRD!VisitId & "|" & rsBRD!VisitCycleNumber & "|" & rsBRD!VisitCycleDate & "|" _
                                        & rsBRD!CRFPageId & "|" & rsBRD!CRFPageCycleNumber & "|" & rsBRD!CRFPageCycleDate
                                End If
                            End If
                        End If
                        If bOKToContinue Then
                            'Create a Response for the current entry
                            Set oResponse = oeFormInstance.Responses.ResponseByElement(oQuestion, rsBRD!RepeatNumber)
                            'Check that oResponse is not Nothing
                            If oResponse Is Nothing Then
                                Call SetUploadMessage(lBatchId, "Question details wrong (check the repeat number)")
                                bOKToContinue = False
                            End If
                        End If
                        If bOKToContinue Then
                            'Check that the Question is not Locked or Frozen
                            If oResponse.LockStatus <> eLockStatus.lsUnlocked Then
                                Call SetUploadMessage(lBatchId, "Question is Locked or Frozen")
                                bOKToContinue = False
                            End If
                        End If
                        If bOKToContinue Then
                            'Check that the Question is enterable
                            bOKToContinue = CheckQuestionEnterable(oResponse, oQuestion, oeFormInstance, oSubject, lBatchId)
                        End If
                        If bOKToContinue Then
                            'Check the need for Authorisation
                            bOKToContinue = CheckAuthorisation(sUserName, oQuestion, lBatchId)
                        End If
                        If bOKToContinue Then
                            'Validate the new response
                            ' NCJ 28 Jun 04 - Set up a lab if necessary
                            If oQuestion.DataType = DataType.LabTest Then
                                Call SetUpLab(sSite, oeFormInstance)
                            End If
                            Select Case oQuestion.DataType
                            'TA 20/04/2004: remove null on the response value becasue code can't handle nulls
                            Case DataType.IntegerData, DataType.LabTest, DataType.Real
                                'if its a numeric question run ConvertStandardToLocalNum over the response
                                nStatus = oResponse.ValidateValue(ConvertStandardToLocalNum(RemoveNull(rsBRD!Response)), sErrMsg, bChanged)
                            Case Else
                                nStatus = oResponse.ValidateValue(RemoveNull(rsBRD!Response), sErrMsg, bChanged)
                            End Select
                            If nStatus = eStatus.InvalidData Then
                                Call SetUploadMessage(lBatchId, "Invalid value - " & sErrMsg)
                                oResponse.RejectValue
                            Else
                                'Check for a changed response
                                If bChanged Then
                                    sRFC = ""
                                    'Check for the need for an automatic ReasonForChange message
                                    ' NCJ 30 Jun 04 - Also check for Lab Result RFC
                                    If oResponse.RequiresValueRFC Then
                                        sRFC = "*** Response changed by Batch Data Entry"
                                    ElseIf oQuestion.DataType = DataType.LabTest Then
                                        ' NB RequiresLabResultRFC takes dummy arguments
                                        If oResponse.RequiresLabResultRFC(0, 0) Then
                                            ' NR/CTC has changed
                                            sRFC = "*** NR Status or CTC Grade changed by Batch Data Entry"
                                        End If
                                    End If
                                    'Confirm the valid response
                                    ' NCJ 27 Sept 07 - Bug 2937 - Only pass user name (and user name full) for authorisation questions
                                    If oResponse.Element.Authorisation > "" Then
                                        Call oResponse.ConfirmValue("", sRFC, sUserName, sUserNameFull)
                                    Else
                                        Call oResponse.ConfirmValue("", sRFC, "")
                                    End If
                        
                                    'Check for the need to generate the next row of a RQG
                                    If Not oQuestion.OwnerQGroup Is Nothing Then
                                        Call oeFormInstance.QGroupInstanceById(oQuestion.OwnerQGroup.QGroupID).CreateNewRow
                                    End If
                                    ' NCJ 27 Sept 07 - Bug 2935 - Don't pass user name here
'                                    Call oeFormInstance.RefreshSkipsAndDerivations(ChangingResponse, rsBRD!UserName, oResponse)
                                    Call oeFormInstance.RefreshSkipsAndDerivations(ChangingResponse, "", oResponse)
                                    'Mo 26/10/2005 COD0020
                                    If gbClinicalCoding Then
                                        If oResponse.Element.DataType = eDataType.Thesaurus Then
                                            If oResponse.CodingStatus <> eCodingStatus.csNotCoded Then
                                                 'load the coded term
                                                Set oCodedTermHistory = New MACROCCBS30.CodedTermHistory
                                                Call oCodedTermHistory.InitAuto(goUser.CurrentDBConString, CLng(lClinicalTrialId), CStr(sSite), _
                                                    CLng(lPersonId), CLng(oResponse.ResponseId), CInt(oResponse.RepeatNumber))
                                                'set the new status
                                                Call oCodedTermHistory.SetStatus(CInt(eCodingStatus.csNotCoded), goUser.UserName, goUser.UserNameFull, _
                                                    oResponse.Value, CDbl(oResponse.TimeStamp), CInt(oTimeZone.TimezoneOffset))
                                                'save the changed value
                                                Call oCodedTermHistory.Save(goUser.CurrentDBConString, CLng(oVisitInstance.Visit.VisitId), _
                                                    CInt(oVisitInstance.CycleNo), CLng(oeFormInstance.eForm.EFormId), _
                                                    CInt(oeFormInstance.CycleNo))
                                                Set oCodedTermHistory = Nothing
                                            End If
                                        End If
                                    End If
                                End If
                                'Mo 30/6/2008 - WO-080002
                                'Check for Unobtainable status being set
                                If rsBRD!Unobtainable = 1 Then
                                    If ((oResponse.Status = eStatus.Missing) Or (oResponse.Status = eStatus.Requested)) Then
                                        'Check for RequiresStatusRFC
                                        If oResponse.RequiresStatusRFC(eStatus.Unobtainable) Then
                                            Call oResponse.SetStatus(eStatus.Unobtainable, "*** Status changed to Unobtainable by Batch Data Entry", "")
                                        Else
                                            Call oResponse.SetStatus(eStatus.Unobtainable, "", "")
                                        End If
                                    End If
                                End If
                                'Not that if bChanged is not set it means that an identical response
                                'has been presented. It is better to delete such entries
                                'Mark the successfully uploaded Batch Response entry for deletion
                                nBatchDeleteNumber = nBatchDeleteNumber + 1
                                ReDim Preserve alBatchIdsForDeletion(nBatchDeleteNumber)
                                alBatchIdsForDeletion(nBatchDeleteNumber) = lBatchId
                            End If
                        End If
                    End If  'that started If Not bSubjectloaded
                End If  'that started - If lPersonId > 0
            End If  'that started - If Not bStudyExists
        End If  'that strted - If Not bUserHasrights
        rsBRD.MoveNext
    Loop
    
    'Save the last loaded eform
    If sFormLoadDetails <> "" Then
        'Save the previously loaded eForm's responses
        Call SaveFormResponses(oSubject, oeFormInstance, alBatchIdsForDeletion, sUserName)
    End If
    
    frmMenu.txtProgress.Text = "Upload completed"
    frmMenu.txtProgress.Refresh
    
    'Setting objects to nothing
    If Not oSubject Is Nothing Then
        Set oSubject = Nothing
    End If
    
    If Not moStudyDef Is Nothing Then
        moStudyDef.Terminate
    End If
    
    Set moStudyDef = Nothing
    Set oTimeZone = Nothing
    
    Call HourglassOff
    
    Call UnlockBatchUpload

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UploadBatchResponses", "modBatchDataEntry")
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
Private Sub SetUploadMessage(ByVal lBatchResponseId As Long, _
                            ByVal sMessage As String)
'---------------------------------------------------------------------
' NCJ 29 Jan 04 - Must deal with single quotes occurring in sMessage
' NCJ 30 Jun 04 - Timestamp and " - " added to front of message, and "Unable to upload" to end
'---------------------------------------------------------------------
Dim sSQL As String
Dim sText As String

    On Error GoTo ErrLabel
    
    sText = Format(Now, "dd/mm/yyyy hh:mm:ss") & " - " & sMessage & " - unable to upload"

    ' NCJ 29 Jan 04 - Added ReplaceQuotes around sMessage
    sSQL = "UPDATE BatchResponseData " _
        & "SET UploadMessage = '" & ReplaceQuotes(sText) & "'" _
        & " WHERE BatchResponseId = " & lBatchResponseId
    MacroADODBConnection.Execute sSQL


Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modBatchDataEntry.SetUploadMessage"
End Sub

'---------------------------------------------------------------------
Private Function PersonIdExists(ByVal lClinicalTrialId, _
                                ByVal sSite As String, _
                                ByVal lPersonId As Long) As Boolean
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bExists As Boolean

    On Error GoTo ErrHandler

    sSQL = "SELECT PersonId FROM TrialSubject " _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND TrialSite = '" & sSite & "'" _
        & " AND PersonId = " & lPersonId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        bExists = False
    Else
        bExists = True
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    PersonIdExists = bExists
        
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "PersonIdExists", "modBatchDataEntry")
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
Public Function LockSubject(ByVal sUser As String, _
                            ByVal lStudyId As Long, _
                            ByVal sSite As String, _
                            ByVal lSubjectId As Long, _
                            ByRef sMessage As String) As String
'---------------------------------------------------------------------
'Lock a subject. Based on MACRO_DM's modDataEntry.LockSubject
'Returns a token if lock is successful or empty string if not.
'If the Subject can not be locked sMessage is set to the reason.
' NCJ 1 Jul 04 - Return more meaningful error messages
'---------------------------------------------------------------------
Dim sToken As String
Dim sLockDetails As String

    On Error GoTo ErrLabel
    
    sToken = MACROLOCKBS30.LockSubject(gsADOConnectString, sUser, lStudyId, sSite, lSubjectId)
    Select Case sToken
    Case MACROLOCKBS30.DBLocked.dblStudy
        sLockDetails = MACROLOCKBS30.LockDetailsStudy(gsADOConnectString, lStudyId)
        If sLockDetails = "" Then
            sMessage = "This study is currently being edited by another user."
        Else
            sMessage = "This study is currently being edited by " & Split(sLockDetails, "|")(0) & "."
        End If
        sToken = ""
    Case MACROLOCKBS30.DBLocked.dblSubject
        sLockDetails = MACROLOCKBS30.LockDetailsSubject(gsADOConnectString, lStudyId, sSite, lSubjectId)
        If sLockDetails = "" Then
            sMessage = "This subject is currently being edited by another user."
        Else
            sMessage = "This subject is currently being edited by " & Split(sLockDetails, "|")(0) & "."
        End If
        sToken = ""
    Case MACROLOCKBS30.DBLocked.dblEFormInstance
        ' An eForm is in use, but we don't know which one, so give a generic message
        sMessage = "This subject is currently being edited by another user."
        sToken = ""
    End Select
    
    LockSubject = sToken
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modBatchDataEntry.LockSubject"

End Function

'---------------------------------------------------------------------
Public Sub UnlockSubject(ByVal lStudyId As Long, _
                        ByVal sSite As String, _
                        ByVal lSubjectId As Long, _
                        ByVal sToken As String)
'---------------------------------------------------------------------
'Unlock a subject. Based on MACRO_DM's modDataEntry.UnlockSubject
'---------------------------------------------------------------------

    On Error GoTo ErrLabel

    If sToken <> "" Then
        'if no gsStudyToken then UnlockSubject is being called without a corresponding LockSubject being called first
        MACROLOCKBS30.UnlockSubject gsADOConnectString, sToken, lStudyId, sSite, lSubjectId
        'always set this to empty string for same reason as above
        sToken = ""
    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modBatchDataEntry.UnlockSubject"

End Sub

'---------------------------------------------------------------------
Private Function UserCanChangeData(ByVal sUserName As String, _
                                    ByRef sUserNameFull As String, _
                                    ByRef sMessage As String) As Boolean
'---------------------------------------------------------------------
'Returns True if user specified in sUserName has permission "F5003" - gsFnChangeData.
'Otherwise it returns false
'---------------------------------------------------------------------
' Mo 14/2/2007 Bug 2876
'   sRoleCode removed as a ByRef parameter.
'   UserCanChangeData now checks all Roles that a user might have
' NCJ 27 Sept 07 - Bug 2937 - Also returns UserNameFull
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsRoles As ADODB.Recordset
Dim rsTemp As ADODB.Recordset
Dim bOKToContinue As Boolean
Dim sRoleCode As String
Dim bFound As Boolean

    On Error GoTo ErrHandler
    
    bOKToContinue = True
    sMessage = ""
    
    'Check UserName is known
    ' NCJ 27 Sept 07 - Retrieve UserNameFull too
    sSQL = "SELECT UserName, USERNAMEFULL FROM MACROUser " _
        & "WHERE UserName = '" & sUserName & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        sMessage = "Unknown UserName"
        bOKToContinue = False
    End If
    
    If bOKToContinue Then
        ' NCJ 27 Sept 07 - Remember user name full
        sUserNameFull = rsTemp!UserNameFull
        'Check UserName
        sSQL = "SELECT RoleCode FROM UserRole " _
            & "WHERE UserName = '" & sUserName & "'"
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
        If rsTemp.RecordCount = 0 Then
            sMessage = "UserName does not have permissions for this database"
            bOKToContinue = False
        End If
    End If
    
    'initialize bFound
    bFound = False
    
    If bOKToContinue Then
        'Loop through rsTemp checking to see if at least one of the roles contains permission "F5003"
        ' NCJ 27 Sept 07 - NB We *should* be checking the study/site, too!
        Do Until rsTemp.EOF Or bFound
            'Check that the Users RoleCode contains permision "F5003"
            sSQL = "SELECT * FROM RoleFunction " _
                & "WHERE RoleCode = '" & rsTemp!rolecode & "'" _
                & " AND FunctionCode = 'F5003'"
            Set rsRoles = New ADODB.Recordset
            rsRoles.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
            
            If rsRoles.RecordCount = 0 Then
                sMessage = "UserName does not have permissions to change data"
                bOKToContinue = False
            Else
                bFound = True
            End If
            rsTemp.MoveNext
        Loop
        rsRoles.Close
        Set rsRoles = Nothing
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    If bOKToContinue Then
        UserCanChangeData = True
    Else
        UserCanChangeData = False
    End If

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UserCanChangeData", "modBatchDataEntry")
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
Private Function CheckPersonIdLabel(ByVal lClinicalTrialId As Long, _
                                        ByVal sSite As String, _
                                        ByVal lSubjectId As Long, _
                                        ByVal sSubjectLabel As Variant, _
                                        ByVal lBatchId As Long) As Long
'---------------------------------------------------------------------
Dim lPersonId As Long

    On Error GoTo ErrHandler

    'Check for a SubjectLabel instead of a PersonId and retrieve the PersonId (if the SubjectLabel is unique)
    'When a SubjectLabel is supplied the PersonId is set to 0.
    If lSubjectId = 0 Then
        'Check the supplied SubjectLabel
        lPersonId = IdFromTrialSiteSubjectLabel(lClinicalTrialId, sSite, sSubjectLabel)
        If lPersonId = 0 Then
            Call SetUploadMessage(lBatchId, "Unknown Subject Label")
        End If
        If lPersonId = -1 Then
            Call SetUploadMessage(lBatchId, "Subject Label is not unique")
        End If
    Else
        'Check the supplied PersonId
        lPersonId = lSubjectId
        If Not PersonIdExists(lClinicalTrialId, sSite, lPersonId) Then
            Call SetUploadMessage(lBatchId, "Subject Id does not exist")
            lPersonId = 0
        End If
    End If
    
    CheckPersonIdLabel = lPersonId

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CheckPersonIdLabel", "modBatchDataEntry")
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
Private Function CheckVisiteFormQuestion(ByVal lVisitId As Long, _
                                            ByVal lCRFPageId As Long, _
                                            ByVal lDataItemId As Long, _
                                            ByRef oQuestion As eFormElementRO, _
                                            ByVal lBatchId As Long) As Boolean
'---------------------------------------------------------------------
Dim oVisit As VisitRO
Dim oForm As eFormRO
Dim oVisitForm As VisitEFormRO
Dim bFound As Boolean

    'On Error GoTo ErrHandler
    On Error Resume Next
    'Check the VisitId is still valid for this Study
    Set oVisit = moStudyDef.VisitById(lVisitId)
    'If oVisit Is Nothing Then
    If Err.Number <> 0 Then
        Err.Clear
        Call SetUploadMessage(lBatchId, "Invalid visit")
        CheckVisiteFormQuestion = False
        Exit Function
    End If

    'Check the CRFPageId is still valid for this Study
    Set oForm = moStudyDef.eFormById(lCRFPageId)
    'If oForm Is Nothing Then
    If Err.Number <> 0 Then
        Err.Clear
        Call SetUploadMessage(lBatchId, "Invalid eForm")
        CheckVisiteFormQuestion = False
        Exit Function
    End If

    'Check the eForm occurs in the Visit
    Set oVisitForm = oVisit.VisitEFormByEForm(oForm)
    'If oVisitForm Is Nothing Then
    If Err.Number <> 0 Then
        Err.Clear
        Call SetUploadMessage(lBatchId, "Invalid visit/eForm combination")
        CheckVisiteFormQuestion = False
        Exit Function
    End If
    
    On Error GoTo ErrHandler

    'Load the current form's elements, prior to checking the DataItemId is one of its elements
    Call moStudyDef.LoadElements(oForm)
    bFound = False
    For Each oQuestion In oForm.EFormElements
        If oQuestion.QuestionId = lDataItemId Then
            bFound = True
            Exit For
        End If
    Next oQuestion
    If Not bFound Then
        Call SetUploadMessage(lBatchId, "Invalid question")
        CheckVisiteFormQuestion = False
        Exit Function
    End If
    
    'everything is valid
    CheckVisiteFormQuestion = True

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CheckVisiteFormQuestion", "modBatchDataEntry")
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
Private Function CheckVisitCycleNumberDate(ByVal lVisitId As Long, _
                                            ByVal nVisitCycleNumber As Integer, _
                                            ByVal dblVisitCycleDate As Double, _
                                            ByRef oSubject As StudySubject, _
                                            ByRef oVisitInstance As VisitInstance, _
                                            ByVal lBatchId As Long) As Boolean
'---------------------------------------------------------------------
Dim bFound As Boolean

    On Error GoTo ErrHandler

    For Each oVisitInstance In oSubject.VisitInstancesById(lVisitId)
        bFound = False
        If nVisitCycleNumber = 0 Then
            'Look for a match on VisitCycelDate
            If oVisitInstance.VisitDate = dblVisitCycleDate Then
                bFound = True
                Exit For
            End If
        Else
            'look for a match VisitCycelNumber
            If oVisitInstance.CycleNo = nVisitCycleNumber Then
                bFound = True
                Exit For
            End If
        End If
    Next oVisitInstance
    
    If Not bFound Then
        If nVisitCycleNumber = 0 Then
            Call SetUploadMessage(lBatchId, "Visit date does not exist")
        Else
            Call SetUploadMessage(lBatchId, "Visit cycle number does not exist")
        End If
        CheckVisitCycleNumberDate = False
    Else
        CheckVisitCycleNumberDate = True
    End If
    
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CheckVisitCycleNumberDate", "modBatchDataEntry")
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
Private Function CheckeFormCycleNumberDate(ByVal lCRFPageId As Long, _
                                            ByVal nCRFPageCycleNumber As Integer, _
                                            ByVal dblCRFPageCycleDate As Double, _
                                            ByRef oSubject As StudySubject, _
                                            ByRef oVisitInstance As VisitInstance, _
                                            ByRef oeFormRO As eFormRO, _
                                            ByRef oeFormInstance As EFormInstance, _
                                            ByVal lBatchId As Long) As Boolean
'---------------------------------------------------------------------
Dim bFound As Boolean

    On Error GoTo ErrHandler

    Set oeFormRO = moStudyDef.eFormById(lCRFPageId)
    For Each oeFormInstance In oVisitInstance.eFormInstancesByEForm(oeFormRO)
        bFound = False
        If nCRFPageCycleNumber = 0 Then
            'Look for a match on CRFPageCycleDate
            If oeFormInstance.eFormDate = dblCRFPageCycleDate Then
                bFound = True
                Exit For
            End If
        Else
            'look for a match CRFPageCycleNumber
            If oeFormInstance.CycleNo = nCRFPageCycleNumber Then
                bFound = True
                Exit For
            End If
        End If
    Next oeFormInstance
    If Not bFound Then
        If nCRFPageCycleNumber = 0 Then
            Call SetUploadMessage(lBatchId, "eForm date does not exist")
        Else
            Call SetUploadMessage(lBatchId, "eForm cycle number does not exist")
        End If
        CheckeFormCycleNumberDate = False
    Else
        CheckeFormCycleNumberDate = True
    End If

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CheckeFormCycleNumberDate", "modBatchDataEntry")
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
Private Function CheckQuestionEnterable(ByRef oResponse As Response, _
                                        ByRef oQuestion As eFormElementRO, _
                                        ByRef oeFormInstance As EFormInstance, _
                                        ByRef oSubject As StudySubject, _
                                        ByVal lBatchId As Long) As Boolean
'---------------------------------------------------------------------
'The original test for enterable was as follows:-
'
'    If (oResponse.Status = eStatus.NotApplicable) _
'        Or (oQuestion.DerivationExpr > "") _
'        Or (oQuestion.Hidden) _
'        Or (oQuestion.DataType = eDataType.Category And Not oQuestion.ActiveCategories) Then
'
'The Hidden element has been taken out of the logic so that Hidden questions that do not have
'a Derivation can be entered via Batch Data Entry.
'Note that Hidden questions without a Derivation occur as special VTRACK questions
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If (oResponse.Status = eStatus.NotApplicable) _
        Or (oQuestion.DerivationExpr > "") _
        Or (oQuestion.DataType = eDataType.Category And Not oQuestion.ActiveCategories) Then
        Call SetUploadMessage(lBatchId, "Question not enterable")
        CheckQuestionEnterable = False
    Else
        CheckQuestionEnterable = True
    End If

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CheckQuestionEnterable", "modBatchDataEntry")
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
Private Function CheckAuthorisation(ByVal sUserName As String, _
                                    ByRef oQuestion As eFormElementRO, _
                                    ByVal lBatchId As Long) As Boolean
'---------------------------------------------------------------------
' Mo 14/2/2007 Bug 2876
'   sUserName added as a ByVal parameter.
'   sRoleCode removed as a ByVal parameter.
'   CheckAuthorisation now checks all of a Users Roles for a match to the Authorisation Role
'---------------------------------------------------------------------
Dim bOK As Boolean
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bFound As Boolean

    On Error GoTo ErrHandler

    'initialize bOK
    bOK = True
    'initialize bFound
    bFound = True
    
    If oQuestion.Authorisation > "" Then
        bFound = False
        'get all of a users roles
        sSQL = "SELECT RoleCode FROM UserRole " _
            & "WHERE UserName = '" & sUserName & "'"
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
        'Loop through rsTemp checking to see if at least one of the roles matches the Authorisation Role
        Do Until rsTemp.EOF Or bFound
            ' NCJ 27 Sept 07 - Ignore case!
            If LCase(rsTemp!rolecode) = LCase(oQuestion.Authorisation) Then
                bFound = True
            End If
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        Set rsTemp = Nothing
    End If
    
    If Not bFound Then
        Call SetUploadMessage(lBatchId, "This question requires an Authorisation role that you do not have")
        bOK = False
    End If
    
    CheckAuthorisation = bOK

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CheckAuthorisation", "modBatchDataEntry")
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
Public Function QuestionIsRQG() As Boolean
'---------------------------------------------------------------------
'Returns True if a Question is a RQG (Repeating Question Group) Question
'or False if it is not
'If for some reason the question is not found False is returned
'---------------------------------------------------------------------
'Mo 18/3/2003   Filtering on glSelTrialId added to SQL
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bRQG As Boolean

    On Error GoTo ErrHandler
    
    sSQL = "SELECT CRFElement.OwnerQGroupId " _
        & "FROM DataItem, CRFElement " _
        & "WHERE DataItem.ClinicalTrialId = CRFElement.ClinicalTrialId " _
        & "AND DataItem.DataItemId = CRFElement.DataItemId " _
        & "AND CRFElement.ClinicalTrialId = " & glSelTrialId & " " _
        & "AND CRFElement.CRFPageId = " & glSelCRFPageId & " " _
        & "AND DataItem.DataItemId = " & glSelDataItemId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        bRQG = False
    Else
        If rsTemp!OwnerQGroupId > 0 Then
            bRQG = True
        Else
            bRQG = False
        End If
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    QuestionIsRQG = bRQG
    
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "QuestionIsRQG", "modBatchDataEntry")
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
Private Sub SaveFormResponses(ByRef oSubject As StudySubject, _
                            ByRef oeFormInstance As EFormInstance, _
                            ByRef alBatchIdsForDeletion() As Long, _
                            ByVal sUserName As String)
'---------------------------------------------------------------------
'This will make 5 attempts at Saving the current forms Responses before giving up
' NCJ 28 Jun 04 - Do registration after saving form
' NCJ 21 Jun 06 - Bug 2718 - Moved RemoveResponses to AFTER DoRegistration
'---------------------------------------------------------------------
Dim oTimeZone As Timezone
Dim nTimeZoneOffSet As Integer
Dim j As Integer
Dim nSaveResponsesResult As eSaveResponsesResult
Dim sLockErrMsg As String
Dim bSaved As Boolean
Dim sRegReturn As String

    On Error GoTo ErrLabel
     
    Set oTimeZone = New Timezone
    nTimeZoneOffSet = oTimeZone.TimezoneOffset

    bSaved = False
    For j = 1 To 5
        'Save the previously loaded eForm's responses
        nSaveResponsesResult = oSubject.SaveResponses(oeFormInstance, sLockErrMsg, nTimeZoneOffSet)
        If nSaveResponsesResult = eSaveResponsesResult.srrSuccess Then
            'The Save was a Success, so exit the For Loop
            bSaved = True
            Exit For
        ElseIf nSaveResponsesResult = eSaveResponsesResult.srrSubjectReloaded Then
            ' NCJ 27 Sept 07 - Bug 2935 - Don't pass user name to RefreshSkips
'            Call oeFormInstance.RefreshSkipsAndDerivations(ReloadingSubjectData, sUsername)
            Call oeFormInstance.RefreshSkipsAndDerivations(ReloadingSubjectData, "")
        End If
    Next j
    
    'Check that the eForm Responses have been successfully saved
    If bSaved Then
        Call DeleteSavedEntries(alBatchIdsForDeletion)
    Else
        Call CommentUnSavedEntries(alBatchIdsForDeletion)
    End If
    
    ' NCJ 28 Jun 04 - Do registration
    ' (this one call does all the necessaries)
    Call DoRegistration(oSubject, oeFormInstance.eFormTaskId, _
                        goUser.CurrentDBConString, goUser.DatabaseCode, sRegReturn)
    
    ' A call to RemoveResponses will remove the locks put in place by this software
    ' NCJ 21 Jun 06 - Moved this to AFTER DoRegistration
    Call oSubject.RemoveResponses(oeFormInstance, True)

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modBatchDataEntry.SaveFormResponses"
End Sub

'---------------------------------------------------------------------
Private Sub DeleteSavedEntries(ByRef alBatchIdsForDeletion() As Long)
'---------------------------------------------------------------------
Dim i As Integer
Dim sSQL As String

    On Error GoTo ErrLabel
    
    For i = 1 To UBound(alBatchIdsForDeletion)
        sSQL = "DELETE FROM BatchResponseData " _
            & "WHERE BatchResponseId = " & alBatchIdsForDeletion(i)
        MacroADODBConnection.Execute sSQL
    Next i

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modBatchDataEntry.DeleteSavedEntries"
    
End Sub

'---------------------------------------------------------------------
Private Sub CommentUnSavedEntries(ByRef alBatchIdsForDeletion() As Long)
'---------------------------------------------------------------------
Dim i As Integer

    On Error GoTo ErrLabel
    
    For i = 1 To UBound(alBatchIdsForDeletion)
        Call SetUploadMessage(alBatchIdsForDeletion(i), "Locks prevented responses being saved")
    Next i

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modBatchDataEntry.CommentUnSavedEntries"

End Sub

'---------------------------------------------------------------------
Private Sub SetUpLab(sSite As String, oEFI As EFormInstance)
'---------------------------------------------------------------------
' NCJ 28 Jun 04
' Set up a lab for this eForm if there isn't one already
' Take the first lab definition we find (if any)
'---------------------------------------------------------------------
Dim oLabs As clsLabs
Dim oLab As clsLab
Dim bGotLab As Boolean

    On Error GoTo ErrLabel
    
    ' If we've already done it, don't bother again
    If mbFormLabSet Then Exit Sub
    
    bGotLab = False
    
    ' See if there are any labs
    Set oLabs = New clsLabs
    Call oLabs.Load(sSite)
    If oLabs.Count > 0 Then
        ' There are labs available!
        If oEFI.LabCode > "" And oEFI.LabExists Then
            ' See if they're OK with what they've got
            On Error Resume Next
            Set oLab = oLabs.Item(oEFI.LabCode)
            
            On Error GoTo ErrLabel
            
            If Not oLab Is Nothing Then
                ' It still exists so we're OK
                bGotLab = True
            End If
        End If
        If Not bGotLab Then
            ' Take the first lab
            Set oLab = oLabs.Item(1)
            oEFI.LabCode = oLab.Code
        End If
    Else
        ' No labs available for this site
        oEFI.LabCode = ""
    End If
    
    ' We've done it now (mbFormLabSet will get reset when the next eForm is loaded)
    mbFormLabSet = True
    
    Set oLab = Nothing
    Set oLabs = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modBatchDataEntry.SetUpLab"

End Sub

'---------------------------------------------------------------------
Public Sub TidyUpBDE()
'---------------------------------------------------------------------
' NCJ 30 Jun 04 - Make sure we tidy up after ourselves
'---------------------------------------------------------------------

    ' Ignore errors here
    On Error Resume Next
    
    If Not moStudyDef Is Nothing Then
        Call moStudyDef.Terminate
        Set moStudyDef = Nothing
    End If
    
End Sub

'---------------------------------------------------------------------
Public Function GetStudyStatusId(ByVal lStudyId As Long) As Integer
'---------------------------------------------------------------------
' Returns a study's StatusId
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim sSQL As String
Dim nStatusId As Integer

    On Error GoTo ErrLabel
    
    sSQL = "SELECT StatusId FROM ClinicalTrial WHERE ClinicalTrialId = " & lStudyId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        nStatusId = 0
    Else
        nStatusId = rsTemp!statusId
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    GetStudyStatusId = nStatusId

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modBatchDataEntry.GetStudyStatusId"

End Function

'---------------------------------------------------------------------
Public Function LockBatchUpload() As Boolean
'---------------------------------------------------------------------
Dim sSQL As String
    
    On Error Resume Next

    sSQL = "INSERT INTO MACROUserSettings VALUES('BUL','Batch Upload Lock',1)"
    MacroADODBConnection.Execute sSQL

    If Err.Number Then
        Call DialogInformation("Batch Data Entry Upload is currently Locked. Try again later." & vbNewLine _
            & "If Upload remains locked, unlock it by clicking 'UnLock Batch Data Entry Upload' from the File menu.", "Batch Data Entry Upload")
        LockBatchUpload = False
    Else
        LockBatchUpload = True
    End If

End Function

'---------------------------------------------------------------------
Public Sub UnlockBatchUpload()
'---------------------------------------------------------------------
Dim sSQL As String

    sSQL = "DELETE FROM MACROUserSettings " _
        & "WHERE UserName = 'BUL' " _
        & "AND UserSetting = 'Batch Upload Lock'"
    MacroADODBConnection.Execute sSQL
        
End Sub
