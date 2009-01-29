Attribute VB_Name = "modLISAUtils"

'------------------------------------------------------------------
' File: modLISAUtils.bas
' Copyright InferMed Ltd 2003 All Rights Reserved
' Author: Nicky Johns, August 2003
' Purpose: Supporting routines for the MACRO/LISA interface
'------------------------------------------------------------------
' REVISIONS
' NCJ 14 Aug 2003 - Initial development
' NCJ 27 Aug 03 - Option to load a subject as read-only when retrieving data
' NCJ 21 Jan 04 - Added CreateSubject
' FEB 04 - LISA PHASE II
' NCJ 14 Apr 04 - Check for non-existent visit in LastVisitInstance
' NCJ 20 May 04 - Added error message for Invalid User
' NCJ 19 Jul 04 - Moved FormKey here
'------------------------------------------------------------------

Option Explicit

' The name of the "settings" file containing eForm codes, question codes etc.
Const msLISASettingsFILENAME = "LISADataCodes.txt"

' The XML tags
Public Const gsTAG_VISIT As String = "Visit"
Public Const gsTAG_EFORM As String = "Eform"
Public Const gsTAG_QUESTION As String = "Question"

' The XML attributes
Public Const gsATTR_CODE As String = "Code"
Public Const gsATTR_CYCLE As String = "Cycle"
Public Const gsATTR_VALUE As String = "Value"
Public Const gsATTR_CATCODE As String = "CatCode"

Public Const gsATTR_STUDY As String = "Study"
Public Const gsATTR_SITE As String = "Site"
Public Const gsATTR_LABEL As String = "Label"

' The initial XML for an empty subject
Public Const gsXML_VERSION_HEADER As String = "<?xml version=""1.0""?>"
Public Const gsXML_EMPTY_SUBJ As String = "<MACROSubject> </MACROSubject>"

Public Const gsXMLTAG0_INPUTERR = "<MACROInputErrors>"
Public Const gsXMLTAG1_INPUTERR = "</MACROInputErrors>"

' "Unexpected error" return code
Public Const glERROR_RESULT As Long = -1
' Return code for invalid user
Public Const glINVALID_USER_ERR = -2
' Return code for invalid lock tokens
Public Const glINVALID_LOCKS = 1

Public Const gsINVALID_USER_MSG = "Invalid MACRO user"

'---------------------------------------------------------------------
Public Function LoadSubject(oUser As MACROUser, _
                        ByVal sStudyName As String, _
                        ByRef sSite As String, _
                        ByVal sSubjLabel As String, _
                        ByRef sErrMsg As String, _
                        Optional bReadOnly As Boolean = False) As StudySubject
'---------------------------------------------------------------------
' Load the subject as specified
' Returns Nothing if subject load not successful
' Load as Read-Only (with no AREZZO) if bReadOnly = TRUE
'---------------------------------------------------------------------
Dim sTempPath As String
Dim oSubject As StudySubject
Dim oArezzo As Arezzo_DM
Dim oStudyDef As StudyDefRO
Dim lStudyId As Long
Dim lSubjId As Long
Dim enUpdateMode As eUIUpdateMode

    On Error GoTo ErrLabel
    
    Set LoadSubject = Nothing

Debug.Print Timer & " Starting LoadSubject"

    ' Get settings file (TRUE means look one level up because we're in a DLL)
    Call InitialiseSettingsFile(True)
    sTempPath = GetMACROSetting("Temp", App.Path & "\..\Temp\")

    ' Get the study ID and Subject ID (needed for loading)
    If GetIDsFromNames(oUser, sStudyName, sSite, sSubjLabel, lStudyId, lSubjId) Then
        If Not bReadOnly Then
            enUpdateMode = eUIUpdateMode.Read_Write
        Else
            enUpdateMode = eUIUpdateMode.Read_Only
        End If
        Set oArezzo = InitNewArezzo(sTempPath, oUser.CurrentDBConString, lStudyId)
        Set oStudyDef = LoadStudyDef(oUser.CurrentDBConString, lStudyId, oArezzo, sErrMsg)
        If Not oStudyDef Is Nothing Then
            Call oStudyDef.LoadSubject(sSite, lSubjId, oUser.UserName, enUpdateMode, _
                                        oUser.UserNameFull, oUser.UserRole)
            If oStudyDef.Subject.CouldNotLoad Then
                ' Give up
                sErrMsg = "Unable to open subject: " & oStudyDef.Subject.CouldNotLoadReason
            Else
                ' Successfully loaded
                Set LoadSubject = oStudyDef.Subject
            End If
        End If
    Else
        ' Couldn't recognise the subject
        sErrMsg = "Subject does not exist"
    End If
    
    Set oSubject = Nothing
    Set oArezzo = Nothing
    Set oStudyDef = Nothing

Debug.Print Timer & " Done LoadSubject"

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modLISAUtils.LoadSubject"

End Function

'---------------------------------------------------------------------
Public Function InitNewArezzo(ByVal sTempPath As String, _
                            ByVal sDBConString As String, _
                            ByVal lStudyId As Long) As Arezzo_DM
'---------------------------------------------------------------------
' Create and initialise a new instance of AREZZO for the given study
'---------------------------------------------------------------------
Dim oArezzoMemory As clsAREZZOMemory
Dim oArezzo As Arezzo_DM

    On Error GoTo ErrLabel
    
    'Create and initialise a new Arezzo instance
    Set oArezzo = New Arezzo_DM

    Set oArezzoMemory = New clsAREZZOMemory
    Call oArezzoMemory.Load(lStudyId, sDBConString)
    Call oArezzo.Init(sTempPath, oArezzoMemory.AREZZOSwitches)
    Set oArezzoMemory = Nothing

    Set InitNewArezzo = oArezzo
    Set oArezzo = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modLISAUtils.InitNewArezzo"

End Function


'---------------------------------------------------------------------
Private Function LoadStudyDef(ByVal sDBConString As String, _
                            ByVal lStudyId As Long, _
                            oArezzo As Arezzo_DM, _
                            ByRef sMsg As String) As StudyDefRO
'---------------------------------------------------------------------
' Load and return the study def
' Returns Nothing if study not successfully loaded, with sMsg a suitable message
'---------------------------------------------------------------------
Dim sErrMsg As String
Dim bLoadedOK As Boolean
Dim oStudyDef As StudyDefRO

    On Error GoTo ErrLabel
    
    Set oStudyDef = New StudyDefRO
    sErrMsg = oStudyDef.Load(sDBConString, lStudyId, 1, oArezzo)
    If sErrMsg > "" Then
        ' Give up
        sMsg = "Unable to load study: " & sErrMsg
        Set oStudyDef = Nothing
    End If
    
    Set LoadStudyDef = oStudyDef
    Set oStudyDef = Nothing

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modLISAUtils.LoadStudyDef"

End Function

'--------------------------------------------------------------------
Public Function GetIDsFromNames(oUser As MACROUser, _
                        ByVal sStudyName As String, _
                        ByRef sSite As String, _
                        ByVal sSubjLabel As String, _
                        ByRef lStudyId As Long, _
                        ByRef lSubjId As Long) As Boolean
'--------------------------------------------------------------------
' Get the study ID and subject ID of the specified subject
' sSite can be "", with sSubjLabel the subject label
' or if Site > "", sSubjLabel can be either label or numeric ID
' Returns FALSE if no subject found
'--------------------------------------------------------------------
Dim vSubjects As Variant
    
    On Error GoTo ErrLabel
    
    If sSite > "" Then
        ' We have a site
        If IsNumeric(sSubjLabel) Then
            ' Treat a numeric subject label as a subject ID
            vSubjects = oUser.DataLists.GetSubjectList(, sStudyName, sSite, Val(sSubjLabel))
        Else
            vSubjects = oUser.DataLists.GetSubjectList(sSubjLabel, sStudyName, sSite)
        End If
    Else
        ' We don't have a site - assume the label's OK
        vSubjects = oUser.DataLists.GetSubjectList(sSubjLabel, sStudyName)
    End If
    
    If IsNull(vSubjects) Then
        ' Subject does not exist
        lStudyId = 0
        lSubjId = 0
        GetIDsFromNames = False
'        AddErrorMsg eDataInputError.SubjectNotExist, "Subject does not exist"
    Else
        ' Assume subject is in row 0 - Pick off IDs
        lStudyId = vSubjects(eSubjectListCols.StudyId, 0)
        lSubjId = vSubjects(eSubjectListCols.SubjectId, 0)
        sSite = vSubjects(eSubjectListCols.Site, 0)
        GetIDsFromNames = True
    End If

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modLISAUtils.GetIDsFromNames"

End Function

'---------------------------------------------------
Public Function SetXMLRequest(ByVal sXMLData As String, _
                    oXMLDoc As MSXML2.DOMDocument, _
                    ByRef sStudyName As String, _
                    ByRef sSite As String, _
                    ByRef sSubjLabel As String) As Boolean
'---------------------------------------------------
' Initialise oXMLDoc and load in the XML Data string
' Return the Study, Site and Subject details
' Return FALSE if any of the values are missing or if it's invalid XML
'---------------------------------------------------

    On Error GoTo ErrLabel
    
    SetXMLRequest = False
    
    sStudyName = ""
    sSite = ""
    sSubjLabel = ""
    
    Set oXMLDoc = New MSXML2.DOMDocument
    If oXMLDoc.loadXML(sXMLData) Then
        With oXMLDoc.documentElement
            sStudyName = .getAttribute(gsATTR_STUDY)
'            sSite = .getAttribute(gsATTR_SITE)
            sSubjLabel = .getAttribute(gsATTR_LABEL)
        End With
'        SetXMLRequest = ((sStudyName > "") And (sSite > "") And (sSubjLabel > ""))
        SetXMLRequest = ((sStudyName > "") And (sSubjLabel > ""))
    End If

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modLISAUtils.SetXMLRequest"

End Function

'---------------------------------------------------
Public Sub EFormsQuestionsList(ByRef sEFormList As String, _
                        ByRef sQuestionList As String)
'---------------------------------------------------
' NCJ 24 Feb 04
' Get the list of eForms and questions from the LISA Settings file
' Each is returned as string containing comma-separated list of single-quoted names
' (this format is assumed in the Settings file)
'---------------------------------------------------
Dim oLISASetting As IMEDSettings

Const sEFORMS = "SQLEForms"
Const sQUS = "SQLQuestions"

    On Error GoTo ErrLabel
    
    Set oLISASetting = New IMEDSettings
    
    Call oLISASetting.Init(App.Path & "\" & msLISASettingsFILENAME)
    sEFormList = oLISASetting.GetKeyValue(sEFORMS, "")
    sQuestionList = oLISASetting.GetKeyValue(sQUS, "")

    Set oLISASetting = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modLISAUtils.EFormsQuestionsList"

End Sub

'---------------------------------------------------
Public Sub GetLISALockList(ByRef sLISAVisit As String, _
                        ByRef sVEFList As String, _
                        ByRef sCurForms As String)
'---------------------------------------------------
' NCJ 18 Mar 04
' Get the list of visits and eForms for locking from the LISA Settings file
' Each is returned as string in a known format
' (this format is assumed in the Settings file)
' NCJ 19 Jul 04 - Use new sCUR_SQLFORMS setting
'---------------------------------------------------
Dim oLISASetting As IMEDSettings

' The names of the settings we want
Const sLISA_VISIT = "LISAVisit"
Const sLOCK_FORMS = "LockForms"
Const sCUR_FORMS = "CurrentLISAForms"
Const sCUR_SQLFORMS = "CurrentLISASQLForms"

    On Error GoTo ErrLabel
    
    Set oLISASetting = New IMEDSettings
    
    Call oLISASetting.Init(App.Path & "\" & msLISASettingsFILENAME)
    sLISAVisit = oLISASetting.GetKeyValue(sLISA_VISIT, "")
    sVEFList = oLISASetting.GetKeyValue(sLOCK_FORMS, "")
'    sCurForms = oLISASetting.GetKeyValue(sCUR_FORMS, "")
    sCurForms = oLISASetting.GetKeyValue(sCUR_SQLFORMS, "")

    Set oLISASetting = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modLISAUtils.GetLISALockList"

End Sub

'----------------------------------------------------
Public Function LastVisitInstance(oSubject As StudySubject, lVisitId As Long) As VisitInstance
'----------------------------------------------------
' Return VisitInstance with latest Cycle no.
' Returns Nothing if there are no visit instances with this VisitId
'----------------------------------------------------
Dim oVI As VisitInstance
Dim oLatestVI As VisitInstance
Dim colVIs As Collection

    On Error GoTo ErrLabel
    
    Set oLatestVI = Nothing
    
    Set colVIs = oSubject.VisitInstancesById(lVisitId)
    ' Check there is one
    If colVIs.Count > 0 Then
        Set oLatestVI = colVIs(1)       ' Take the first one
        For Each oVI In colVIs
            If oVI.CycleNo > oLatestVI.CycleNo Then
                Set oLatestVI = oVI
            End If
        Next
    End If
    
    Set LastVisitInstance = oLatestVI
    
    Set oLatestVI = Nothing
    Set oVI = Nothing
    Set colVIs = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modLISAUtils.LastVisitInstance"

End Function

'----------------------------------------------------------------------
Public Function FormKey(ByVal lEFITaskId As Long) As String
'----------------------------------------------------------------------
' Get a unique "key" based on this eForm taskid
'----------------------------------------------------------------------

    FormKey = "K" & lEFITaskId
    
End Function


