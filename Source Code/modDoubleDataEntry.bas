Attribute VB_Name = "modDoubleDataEntry"
'----------------------------------------------------------------------------------------'
' File:         modDoubleDataEntry
' Copyright:    InferMed Ltd. 2000-2006. All Rights Reserved
' Author:       Mo Morris, September 2006
' Purpose:      Contains variable declarations and facilities required by the Double Data Entry Module
'----------------------------------------------------------------------------------------'
'   Revisions:
'
'----------------------------------------------------------------------------------------'
Option Explicit

Public glSelTrialId As Long
Public gsSelTrialName As String
Public gsSelSite As String
Public gsSelSubject As String
Public glSelPersonId As Long
Public glSelVisitId As Long
Public gsSelVisitCode As String
Public glSelVisitCycleNumber As Long
Public glSelCRFPageId As Long
Public glFirstSelCRFPageId As Long
Public gsSelCRFPageCode As String
Public glSelCRFPageCycleNumber As Long
Public glSelRepeatNumber As Long
Public gnTextBoxHeight As Integer
Public gn50CharWidth As Integer
Public gnLabelIndex As Integer
Public gnTextBoxIndex As Integer
Public gnCommandIndex As Integer
Public gnCatCodeIndex As Integer

Public gnRegDDFontSize As Integer
Public gsRegDDColourScheme As String
Public glRegDDLightColour As Long
Public glRegDDMediumColour As Long
Public glRegDDDarKColour As Long

Public gnRegFormLeft As Integer
Public gnRegFormTop As Integer
Public gnRegFormWidth As Integer
Public gnRegFormHeight As Integer
Public gnRegFormWindowState As Integer
Public gnRegFormLeftMax As Integer
Public gnRegFormTopMax As Integer
Public gnRegFormWidthMax As Integer
Public gnRegFormHeightMax As Integer
Public gbRegDisplayCategoryCodes As Boolean

Public glCRFPageIdAfterVisitDate As Long

Public gnPassNumber As Integer

Public Enum ePassNumber
    First = 1
    Second = 2
End Enum

Public Enum eDoubleDataStatus
    Entered = 0
    Verified = 1
End Enum

Public Const mnMINFORMWIDTH As Integer = 10700
Public Const mnMINFORMHEIGHT As Integer = 8000
Public Const mnMINVFORMWIDTH As Integer = 8000
Public Const mnMINVFORMHEIGHT As Integer = 9200

Public Const gnVISUAL_ELEMENT = 16384

Public Enum eBatchUploadStatus
    UnLocked = 0
    Locked = 1
End Enum

'---------------------------------------------------------------------
Public Sub GetDataItemDetails(ByVal lClinicalTrialId As Long, _
                           ByVal lDataItemId As String, _
                           ByRef sDataItemName As String, _
                           ByRef nDataItemType As Integer, _
                           ByRef nDataItemLength As Integer)
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo Errhandler
    
    sSQL = "SELECT DataItemName, DataType, DataItemLength FROM DataItem " _
        & "WHERE ClinicalTrialId = " & lClinicalTrialId & " " _
        & "AND DataItemId = '" & lDataItemId & "'"
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        sDataItemName = ""
        nDataItemType = 0
        nDataItemLength = 0
    Else
        sDataItemName = rsTemp!DataItemName
        nDataItemType = rsTemp!DataType
        nDataItemLength = rsTemp!DataItemLength
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
 
Exit Sub
Errhandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "GetDataItemDetails", "modDoubleDataEntry")
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
Public Sub DDFormCentre(frmForm As Form)
'---------------------------------------------------------------------
'   Centre DDform on screen , ignoring bottom 400 twips for status bar
'---------------------------------------------------------------------

    On Error GoTo Errhandler

    If frmForm.WindowState = vbNormal Then
        gnRegFormTop = (Screen.Height - 400 - frmForm.Height) \ 2
        frmForm.Top = gnRegFormTop
        gnRegFormLeft = (Screen.Width - frmForm.Width) \ 2
        frmForm.Left = gnRegFormLeft
    End If
    
Exit Sub
Errhandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "DDFormCentre", "modDoubleDataEntry")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Public Function QuestionIsDerived(ByVal lClinicalTrialId As Long, _
                                    ByVal lDataItemId As Long) As Boolean
'--------------------------------------------------------------------
'Returns True if a Question is a Derived question
'Returns False for non-Derived questions
'--------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bFunctionOutCome As Boolean

    On Error GoTo Errhandler

    sSQL = "SELECT DataItem.Derivation FROM DataItem " _
        & "WHERE DataItem.ClinicalTrialId = " & lClinicalTrialId _
        & " AND DataItem.DataItemId = " & lDataItemId
        
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount <> 1 Then
        bFunctionOutCome = False
    Else
        If RemoveNull(rsTemp!Derivation) = "" Then
            bFunctionOutCome = False
        Else
            bFunctionOutCome = True
        End If
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    QuestionIsDerived = bFunctionOutCome

Exit Function
Errhandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "QuestionIsDerived", "modDoubleDataEntry")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'--------------------------------------------------------------------
Public Function NexteFormExists(ByVal lClinicalTrialId As Long, _
                                ByVal lVisitId As Long, _
                                ByVal lCurrenteFormId As Long, _
                                ByRef lNexteFormId As Long) As Boolean
'--------------------------------------------------------------------
Dim sSQL As String
Dim rsEForm As ADODB.Recordset
Dim bCurrenteFormReached As Boolean
Dim bNexteFormExists As Boolean

    On Error GoTo Errhandler

    'retrieve all eForms within the selected Visit (apart from visit date eforms)
    sSQL = "SELECT CRFPage.CRFPageId, CRFPage.CRFPageOrder " _
        & "FROM CRFPage, StudyVisitCRFPage " _
        & "WHERE CRFPage.ClinicalTrialId = StudyVisitCRFPage.ClinicalTrialId " _
        & "AND CRFPage.CRFPageId = StudyVisitCRFPage.CRFPageId " _
        & "AND CRFPage.ClinicalTrialId = " & lClinicalTrialId & " " _
        & "AND StudyVisitCRFPage.VisitId = " & lVisitId & " " _
        & "AND StudyVisitCRFPage.eFormUse = 0 " _
        & "ORDER BY CRFPage.CRFPageOrder"
    
    Set rsEForm = New ADODB.Recordset
    rsEForm.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    bNexteFormExists = False
    bCurrenteFormReached = False
    
    Do Until rsEForm.EOF
        If bCurrenteFormReached Then
            'Store Next eFormId and exit
            lNexteFormId = rsEForm!CRFPageId
            bNexteFormExists = True
            Exit Do
        End If
        If rsEForm!CRFPageId = lCurrenteFormId Then
            bCurrenteFormReached = True
        End If
        rsEForm.MoveNext
    Loop

    rsEForm.Close
    Set rsEForm = Nothing
    
    NexteFormExists = bNexteFormExists

Exit Function
Errhandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "NexteFormExists", "modDoubleDataEntry")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'--------------------------------------------------------------------
Public Function PreveFormExists(ByVal lClinicalTrialId As Long, _
                                ByVal lVisitId As Long, _
                                ByVal lCurrenteFormId As Long, _
                                ByRef lPreveFormId As Long) As Boolean
 '--------------------------------------------------------------------
Dim sSQL As String
Dim rsEForm As ADODB.Recordset
Dim bCurrenteFormReached As Boolean
Dim bPreveFormExists As Boolean

    On Error GoTo Errhandler

    'retrieve all eForms within the selected Visit (apart from visit date eforms)
    sSQL = "SELECT CRFPage.CRFPageId, CRFPage.CRFPageOrder " _
        & "FROM CRFPage, StudyVisitCRFPage " _
        & "WHERE CRFPage.ClinicalTrialId = StudyVisitCRFPage.ClinicalTrialId " _
        & "AND CRFPage.CRFPageId = StudyVisitCRFPage.CRFPageId " _
        & "AND CRFPage.ClinicalTrialId = " & lClinicalTrialId & " " _
        & "AND StudyVisitCRFPage.VisitId = " & lVisitId & " " _
        & "AND StudyVisitCRFPage.eFormUse = 0 " _
        & "ORDER BY CRFPage.CRFPageOrder DESC"
    
    Set rsEForm = New ADODB.Recordset
    rsEForm.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    bPreveFormExists = False
    bCurrenteFormReached = False
    
    Do Until rsEForm.EOF
        If bCurrenteFormReached Then
            'Store Prev eFormId and exit
            lPreveFormId = rsEForm!CRFPageId
            bPreveFormExists = True
            Exit Do
        End If
        If rsEForm!CRFPageId = lCurrenteFormId Then
            bCurrenteFormReached = True
        End If
        'Jump out of this loop when glFirstSelCRFPageId is reached
        'This prevents navigation to eForms that appear before
        '(in study order) the first entered eForm
        If rsEForm!CRFPageId = glFirstSelCRFPageId Then
            Exit Do
        End If
        rsEForm.MoveNext
    Loop

    rsEForm.Close
    Set rsEForm = Nothing
    
    PreveFormExists = bPreveFormExists

Exit Function
Errhandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "PreveFormExists", "modDoubleDataEntry")
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
Public Function GetEFormQuestions(ByVal lClinicalTrialId As Long, _
                            ByVal lCRFPageId As Long) As ADODB.Recordset
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo Errhandler

    sSQL = "SELECT CRFElement.DataItemId, CRFElement.OwnerQGroupId, CRFElement.QGroupId  " _
        & "FROM CRFElement " _
        & "WHERE CRFElement.ClinicalTrialId = " & lClinicalTrialId & " " _
        & "AND CRFElement.CRFPageId = " & lCRFPageId & " " _
        & "AND CRFElement.ControlType < " & gnVISUAL_ELEMENT & " " _
        & "ORDER BY CRFElement.FieldOrder, CRFElement.QGroupFieldOrder"

    Set GetEFormQuestions = New ADODB.Recordset
    GetEFormQuestions.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

Exit Function
Errhandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "GetEFormQuestions", "modDoubleDataEntry")
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
Public Sub LaunchBatchDataUpload()
'---------------------------------------------------------------------
Dim sCommandLine As String
Dim sUserName As String
Dim sDatabase As String
Dim sRoleCode As String
Dim i As Integer

    On Error GoTo Errhandler

    'extract the required launch parameters
    sUserName = goUser.UserName
    sDatabase = goUser.DatabaseCode
    sRoleCode = goUser.UserRole
    'Prepare the Batch Data Entry Upload Commad Line
    sCommandLine = App.Path & "\" & "MACRO_BD.exe /BU/" & sUserName & "/" & gsEnteredPassword & "/" & sDatabase & "/" & sRoleCode
    'Launch Batch Data Entry Upload
    i = ExecCmdNoWait(sCommandLine)

Exit Sub
Errhandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "LaunchBatchDataUpload", "modDoubleDataEntry")
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
Public Function CRFPageNameFromId(ByVal lClinicalTrialId As Long, _
                            ByVal lCRFPageId As Long) As String
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo Errhandler

    sSQL = "SELECT CRFTitle FROM CRFPage" _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND CRFPageId = " & lCRFPageId
        
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        CRFPageNameFromId = ""
    Else
        CRFPageNameFromId = rsTemp!CRFTitle
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Function
Errhandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CRFPageNameFromId", "modDoubleDataEntry")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function
