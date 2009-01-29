Attribute VB_Name = "modJScript"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   File:       modJScript.bas
'   Author:     ZA, 18/09/2002
'   Purpose:    Create java script files with meta data
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'Revisions
'   ZA 19/09/2002 - removed Script header/footer and spaces between elements in the string
'   ZA 26/09/2002 - Donot include library study in CreateTrialList function
'   ic 11/11/2002 - changed js file name constants
'   ic 25/11/2002 - rewrote functions using user object instead of direct db access
'   ic 11/03/2003 - added trialid to calls
'   ic 18/06/2003 changed string format for regular expression
'----------------------------------------------------------------------------------------'
Option Explicit

'list of files generated by functions
Private Const msQUESTIONS_FILE = "lQuestions.js"
'Private Const msSTUDIES_FILE = "List_studies.js"
Private Const msVISITS_FILE = "lVisits.js"
Private Const msFORMS_FILE = "lEForms.js"
'Private Const msSITES_LIST = "List_sites.js"
Private Const msUSERS_LIST = "lUsers.js"
' DPH 17/06/2003 - Web version file
Private Const msWEBVERS_FILE = "lWebVersion.js"

Private Const msDELIMITER1 = "|"
Private Const msDELIMITER2 = "`"


'ic 25/11/2002 not used, studies are no longer cached
''-------------------------------------------------------------------------
'Public Sub CreateTrialList()
''-------------------------------------------------------------------------
''Creates a java script file with trial list in a "|" delimited string
''-------------------------------------------------------------------------
'Dim oFSO As Scripting.FileSystemObject
'Dim oStream As TextStream
'Dim sScriptBody As String
'Dim oRS As ADODB.Recordset
'Dim sQuery As String
'Dim sDescription As String
'
'    On Error GoTo ErrLabel
'
'    Set oFSO = New Scripting.FileSystemObject
'    Set oRS = New ADODB.Recordset
'
'    sQuery = "select ClinicalTrialId, ClinicalTrialName, ClinicalTrialDescription from ClinicalTrial " & _
'             "where ClinicalTrialId > 0"
'
'    oRS.Open sQuery, MacroADODBConnection, adOpenKeyset, adLockOptimistic
'
'    'create a new file under script folder and overwrite if there is already one
'    Set oStream = oFSO.CreateTextFile(gsSCRIPT_FOLDER_LOCATION & msSTUDIES_FILE, True)
'
'    sScriptBody = "var lstStudies = " & Chr(34)
'
'    If Not oRS.EOF Then
'        Do While Not oRS.EOF
'
'            With oRS
'            If IsNull(.Fields("clinicalTrialDescription").Value) Then
'                sDescription = ""
'            Else
'                sDescription = Replace(.Fields("clinicalTrialDescription").Value, vbCrLf, "")
'            End If
'
'            sScriptBody = sScriptBody & .Fields("ClinicalTrialId").Value & _
'                                      "`" & .Fields("ClinicalTrialName").Value & _
'                                      "`" & sDescription & "|"
'
'            .MoveNext
'
'            End With
'        Loop
'    End If
'
'    'remove the last pipe "|" from the string
'    sScriptBody = Left(sScriptBody, Len(sScriptBody) - 1)
'
'    sScriptBody = sScriptBody & Chr(34) & ";"
'
'    'write the variable name and value
'    oStream.Write sScriptBody
'
'    oStream.Close
'
'    Set oRS = Nothing
'    Set oStream = Nothing
'    Set oFSO = Nothing
'
'    Exit Sub
'
'ErrLabel:
'    If MACROErrorHandler("modJScript", Err.Number, Err.Description, "CreateTrialList", Err.Source) = Retry Then
'        Resume
'    End If
'
'End Sub

'-------------------------------------------------------------------------
Public Sub CreateQuestionsList(ByRef oUser As MACROUser, Optional ByVal sTrialId As String)
'-------------------------------------------------------------------------
'Creates java script file with question details in a "|" delimited string
'revisions
'   ic 25/11/2002 changed from direct db access to using user object
'   ic 11/02/2003 added trialid to args
'   ic 18/06/2003 changed string format for regular expression
'-------------------------------------------------------------------------
Dim colStudy As Collection
Dim oStudy As Study
Dim sDatabase As String
Dim sPath As String
Dim sFileName As String
Dim nFile As Integer
Dim vData As Variant
Dim nLoop As Integer
Dim sQList As String
Dim sSearchTrial As String
 
    
    On Error GoTo ErrLabel
 
    'get user database
    sDatabase = goUser.Database.DatabaseCode
    If Not IsMissing(sTrialId) Then sSearchTrial = sTrialId
    
    'get study collection
    Set colStudy = oUser.GetAllStudies
    
    For Each oStudy In colStudy
        
        If (sSearchTrial = "") Or (sSearchTrial = oStudy.StudyId) Then
            'get next free file number
            nFile = FreeFile
        
            'create directory path (asp\database\studyid\)
            sPath = gsWEB_HTML_LOCATION & "sites\" & sDatabase & "\" & oStudy.StudyId & "\"
    
            If (FolderExistence(sPath)) Then
                sFileName = msQUESTIONS_FILE
                
                vData = oUser.DataLists.GetQuestionList(oStudy.StudyId)
                
                If Not IsNull(vData) Then
                    'open the file for output
                    Open sPath & sFileName For Output As #nFile
                    
                    'ic 18/06/2003 changed string format for regular expression
                    sQList = "lstQuestions['" & oStudy.StudyId & "']='" & msDELIMITER1 & msDELIMITER2 & "All Questions"
                    For nLoop = LBound(vData, 2) To UBound(vData, 2)
                        sQList = sQList & msDELIMITER1 & vData(eDropDownCol.Id, nLoop) & msDELIMITER2 _
                                        & RemoveHTMLHostileChars(vData(eDropDownCol.Text, nLoop))
                    Next
                    'If (Right(sQList, 1) = msDELIMITER1) Then sQList = Left(sQList, (Len(sQList) - 1))
                    sQList = sQList & "';"
                    
                    'write to the file
                    Print #nFile, sQList
                    
                    'close the file
                    Close #nFile
                End If
            End If
        End If
    Next

    Set oStudy = Nothing
    Set colStudy = Nothing
    Exit Sub
    
    
ErrLabel:
    If MACROErrorHandler("modJScript", Err.Number, Err.Description, "CreateQuestionsList", Err.Source) = Retry Then
        Resume
    End If
End Sub

'-------------------------------------------------------------------------
Public Sub CreateVisitList(ByRef oUser As MACROUser, Optional ByVal sTrialId As String)
'-------------------------------------------------------------------------
'Creates java script file with visit details in a "|" delimited string
'revisions
'   ic 25/11/2002 changed from direct db access to using user object
'   ic 11/02/2003 added trialid to args
'   ic 18/06/2003 changed string format for regular expression
'-------------------------------------------------------------------------
Dim colStudy As Collection
Dim oStudy As Study
Dim sDatabase As String
Dim sPath As String
Dim sFileName As String
Dim nFile As Integer
Dim vData As Variant
Dim nLoop As Integer
Dim sVList As String
Dim sSearchTrial As String
 
 
    On Error GoTo ErrLabel
 
    'get user database
    sDatabase = goUser.Database.DatabaseCode
    If Not IsMissing(sTrialId) Then sSearchTrial = sTrialId
    
    'get study collection
    Set colStudy = oUser.GetAllStudies
    
    For Each oStudy In colStudy
    
        If (sSearchTrial = "") Or (sSearchTrial = oStudy.StudyId) Then
            'get next free file number
            nFile = FreeFile
        
            'create directory path (asp\database\studyid\)
            sPath = gsWEB_HTML_LOCATION & "sites\" & sDatabase & "\" & oStudy.StudyId & "\"
    
            If (FolderExistence(sPath)) Then
                sFileName = msVISITS_FILE
            
                vData = oUser.DataLists.GetVisitList(oStudy.StudyId)
                
                If Not IsNull(vData) Then
                    'open the file for output
                    Open sPath & sFileName For Output As #nFile
                    
                    'ic 18/06/2003 changed string format for regular expression
                    sVList = "lstVisits['" & oStudy.StudyId & "']='" & msDELIMITER1 & msDELIMITER2 & "All Visits"
                    For nLoop = LBound(vData, 2) To UBound(vData, 2)
                        sVList = sVList & msDELIMITER1 & vData(eDropDownCol.Id, nLoop) & msDELIMITER2 _
                                        & RemoveHTMLHostileChars(vData(eDropDownCol.Text, nLoop))
                    Next
                    'If (Right(sVList, 1) = msDELIMITER1) Then sVList = Left(sVList, (Len(sVList) - 1))
                    sVList = sVList & "';"
                    
                    'write to the file
                    Print #nFile, sVList
                    
                    'close the file
                    Close #nFile
                End If
            End If
        End If
    Next

    Set oStudy = Nothing
    Set colStudy = Nothing
    Exit Sub
    
    
ErrLabel:
    If MACROErrorHandler("modJScript", Err.Number, Err.Description, "CreateVisitList", Err.Source) = Retry Then
        Resume
    End If
End Sub


'-------------------------------------------------------------------------
Public Sub CreateEFormsList(ByRef oUser As MACROUser, Optional ByVal sTrialId As String)
'-------------------------------------------------------------------------
'Creates java script file with visit details in a "|" delimited string
'revisions
'   ic 25/11/2002 changed from direct db access to using user object
'   ic 11/02/2003 added trialid to args
'   ic 18/06/2003 changed string format for regular expression
'-------------------------------------------------------------------------
Dim colStudy As Collection
Dim oStudy As Study
Dim sDatabase As String
Dim sPath As String
Dim sFileName As String
Dim nFile As Integer
Dim vData As Variant
Dim nLoop As Integer
Dim sEForm As String
Dim sSearchTrial As String
 
 
    On Error GoTo ErrLabel
 
    'get user database
    sDatabase = goUser.Database.DatabaseCode
    If Not IsMissing(sTrialId) Then sSearchTrial = sTrialId
    
    'get study collection
    Set colStudy = oUser.GetAllStudies
    
    For Each oStudy In colStudy
        
        If (sSearchTrial = "") Or (sSearchTrial = oStudy.StudyId) Then
            'get next free file number
            nFile = FreeFile
        
            'create directory path (asp\database\studyid\)
            sPath = gsWEB_HTML_LOCATION & "sites\" & sDatabase & "\" & oStudy.StudyId & "\"
    
            If (FolderExistence(sPath)) Then
                sFileName = msFORMS_FILE
            
                vData = oUser.DataLists.GetEFormList(oStudy.StudyId)
                
                If Not IsNull(vData) Then
                    'open the file for output
                    Open sPath & sFileName For Output As #nFile
                    
                    'ic 18/06/2003 changed string format for regular expression
                    sEForm = "lstEForms['" & oStudy.StudyId & "']='" & msDELIMITER1 & msDELIMITER2 & "All EForms"
                    For nLoop = LBound(vData, 2) To UBound(vData, 2)
                        sEForm = sEForm & msDELIMITER1 & vData(eDropDownCol.Id, nLoop) & msDELIMITER2 _
                                        & RemoveHTMLHostileChars(vData(eDropDownCol.Text, nLoop))
                    Next
                    'If (Right(sEForm, 1) = msDELIMITER1) Then sEForm = Left(sEForm, (Len(sEForm) - 1))
                    sEForm = sEForm & "';"
                    
                    'write to the file
                    Print #nFile, sEForm
                    
                    'close the file
                    Close #nFile
                End If
            End If
        End If
    Next

    Set oStudy = Nothing
    Set colStudy = Nothing
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("modJScript", Err.Number, Err.Description, "CreateEFormsList", Err.Source) = Retry Then
        Resume
    End If
End Sub

'ic 25/11/2002 not used, sites are no longer cached
'-------------------------------------------------------------------------
Public Sub CreateSitesList()
'-------------------------------------------------------------------------
'Creates java script file with question details in a "|" delimited string
'-------------------------------------------------------------------------
    

'Dim oFSO As Scripting.FileSystemObject
'Dim oStream As TextStream
'Dim sScriptBody As String
'Dim oRS As ADODB.Recordset
'Dim sQuery As String
'Dim sDescription As String
'
'    On Error GoTo ErrLabel
'
'    Set oFSO = New Scripting.FileSystemObject
'    Set oRS = New ADODB.Recordset
'
'    sQuery = "Select Site, SiteDescription from Site"
'
'    oRS.Open sQuery, MacroADODBConnection, adOpenKeyset, adLockOptimistic
'
'    'create a new file under script folder and overwrite if there is already one
'    Set oStream = oFSO.CreateTextFile(gsSCRIPT_FOLDER_LOCATION & msSITES_LIST, True)
'
'
'    sScriptBody = "var lstSites = " & Chr(34)
'
'    Do While Not oRS.EOF
'
'        With oRS
'        If IsNull(.Fields("SiteDescription").Value) Then
'            sDescription = ""
'        Else
'            sDescription = Replace(.Fields("SiteDescription").Value, vbCrLf, "")
'        End If
'
'        sScriptBody = sScriptBody & .Fields("Site").Value & _
'                                  "`" & sDescription & "|"
'
'        .MoveNext
'
'        End With
'    Loop
'
'    'remove the last "|" from the string
'    sScriptBody = Left(sScriptBody, Len(sScriptBody) - 1)
'
'    sScriptBody = sScriptBody & Chr(34) & ";"
'
'    'write the variable name and value
'    oStream.Write sScriptBody
'
'    oStream.Close
'
'    Set oRS = Nothing
'    Set oStream = Nothing
'    Set oFSO = Nothing
'
'    Exit Sub
'
'ErrLabel:
'    If MACROErrorHandler("modJScript", Err.Number, Err.Description, "CreateSitesList", Err.Source) = Retry Then
'        Resume
'    End If
'
End Sub

'-------------------------------------------------------------------------
Public Sub CreateUsersList(ByRef oUser As MACROUser)
'-------------------------------------------------------------------------
'Creates java script file with question details in a "|" delimited string
'revisions
'   ic 25/11/2002 changed from direct db access to using user object
'   DPH 17/06/2003 include to CreateWebVersionFile
'   ic 18/06/2003 changed string format for regular expression
'-------------------------------------------------------------------------
Dim colStudy As Collection
Dim oStudy As Study
Dim sPath As String
Dim sFileName As String
Dim nFile As Integer
Dim vData As Variant
Dim nLoop As Integer
Dim sUser As String


    On Error GoTo ErrLabel
 
    'get study collection
    Set colStudy = oUser.GetAllStudies
    
    For Each oStudy In colStudy
        'get next free file number
        nFile = FreeFile
    
        'create directory path (asp\database\studyid\)
        sPath = gsWEB_HTML_LOCATION & "script\"

        If (FolderExistence(sPath)) Then
            sFileName = msUSERS_LIST
            
            vData = oUser.DataLists.GetUserList()
            
            If Not IsNull(vData) Then
                'open the file for output
                Open sPath & sFileName For Output As #nFile
                
                'ic 18/06/2003 changed string format for regular expression
                sUser = "lstUsers='" & msDELIMITER1 & msDELIMITER2 & "All Users"
                For nLoop = LBound(vData, 2) To UBound(vData, 2)
                    sUser = sUser & msDELIMITER1 & vData(eDropDownCol.Id, nLoop) & msDELIMITER2 _
                                    & RemoveHTMLHostileChars(vData(eDropDownCol.Text, nLoop))
                Next
                'If (Right(sUser, 1) = msDELIMITER1) Then sUser = Left(sUser, (Len(sUser) - 1))
                sUser = sUser & "';"
                
                'write to the file
                Print #nFile, sUser
                
                'close the file
                Close #nFile
            End If
        End If
        
        ' DPH 26/02/2007 -- Remove
        Exit For
    Next

    ' DPH 17/06/2003 - include CreateWebVersionFile
    Call CreateWebVersionFile
    
    Set oStudy = Nothing
    Set colStudy = Nothing
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("modJScript", Err.Number, Err.Description, "CreateUsersList", Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub CreateWebVersionFile()
'--------------------------------------------------------------------------------------------------
' DPH 17/07/2003
' Creates a version javascript file incrementing the web version
' assumes is run on the server for web
'--------------------------------------------------------------------------------------------------
Dim nFile As Integer
Dim sFileName As String
Dim sPath As String
Dim sData As String
Dim lVersion As Long

    On Error GoTo ErrLabel

    ' get version number
    lVersion = CLng(GetMACROSetting(MACRO_SETTING_WEBVERSION, "0"))
    
    ' increment version number
    lVersion = lVersion + 1
    
    ' update version number
    If (SetMACROSetting(MACRO_SETTING_WEBVERSION, CStr(lVersion))) Then
    
        ' create version javascript string
        sData = "function fnWebVersion(){return " & lVersion & ";}" & vbCrLf
        
        'get next free file number
        nFile = FreeFile
    
        'create directory path (asp\database\studyid\)
        sPath = gsWEB_HTML_LOCATION & "script\"
    
        If (FolderExistence(sPath)) Then
            sFileName = msWEBVERS_FILE
    
            'open the file for output
            Open sPath & sFileName For Output As #nFile
    
            'write to the file
            Print #nFile, sData
            
            'close the file
            Close #nFile
    
        End If
        
    End If
    
ErrLabel:
    If MACROErrorHandler("modJScript", Err.Number, Err.Description, "CreateWebVersionFile", Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'--------------------------------------------------------------------------------------------------
Private Function RemoveHTMLHostileChars(ByVal sStr As String) As String
'--------------------------------------------------------------------------------------------------
' ic 10/05/2001
' function accepts a string and replaces characters in the string that interrupt html with
' the html equivelent character code or escape sequence
'--------------------------------------------------------------------------------------------------
Dim sRtn As String
    
    sRtn = Replace(sStr, vbCrLf, "\n")
    sRtn = Replace(sRtn, "'", "\'")

    RemoveHTMLHostileChars = sRtn
End Function
