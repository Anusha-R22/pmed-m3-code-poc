Attribute VB_Name = "basTrialEdit"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       basTrialEdit.bas
'   Author:         Andrew Newbigging, November 1997
'   Purpose:    Common SQL routines for maintaining trial definitions.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
 '  1   Andrew Newbigging       24/11/97
'   2   Andrew Newbigging       27/11/97
'   3   Andrew Newbigging       21/01/98
'   4   Andrew Newbigging        24/02/98
'   5   Joanne Lau              10/06/98
'   6   Joanne Lau              29/07/98
'   7   Andrew Newbigging       11/11/98
'       Following routines moved form this module to TrialData module:
'           InsertTrial
'   8   Andrew Newbigging       23/6/99 SR 1090
'       Modified gsUpdateStudyDefinition to check that a default font name has been
'       selected before trying to update this field
'       Mo Morris   8/11/99
'       DAO to ADO conversion
'   WillC 10/11/99 Added error handlers
'   NCJ 13 Dec 99 - ClinicalTrialIds to Long
'   NCJ 27/10/00 - Added TrialDocumentExists
'------------------------------------------------------------------------------------'
Option Explicit
Option Base 0
Option Compare Binary

'Changed Mo Morris 15/9/99
'Function gnCopyCurrentStudyDefinition taken out. No longer required

'---------------------------------------------------------------------
Public Function gdsTrialDocumentList(ClinicalTrialId As Long, VersionId As Integer) As ADODB.Recordset
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM StudyDocument " _
        & " WHERE ClinicalTrialId = " & ClinicalTrialId _
        & " AND VersionId = " & VersionId
    Set gdsTrialDocumentList = New ADODB.Recordset
    gdsTrialDocumentList.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                            "gdsTrialDocumentList", "TrialEdit.bas")
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
Public Function gdsTrialDetails(ClinicalTrialId As Long) As ADODB.Recordset
'---------------------------------------------------------------------
'Creates a Writable recordset.
'---------------------------------------------------------------------
Dim sSQL As String

On Error GoTo ErrHandler
    
    sSQL = "SELECT * FROM ClinicalTrial " _
        & " WHERE ClinicalTrialId = " & ClinicalTrialId
    Set gdsTrialDetails = New ADODB.Recordset
    gdsTrialDetails.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
        

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                            "gdsTrialDetails", "TrialEdit.bas")
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
Public Sub gdsAddTrialDocument(ClinicalTrialId As Long, _
                                VersionId As Integer, _
                                sDocumentPath As String)
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim sSQL As String
Dim tmpX As Integer

On Error GoTo ErrHandler


    sSQL = "SELECT max(DocumentId) as MaxDocumentId FROM StudyDocument " _
        & "WHERE ClinicalTrialId = " & ClinicalTrialId _
        & " AND VersionId = " & VersionId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
                       
    If IsNull(rsTemp!MaxDocumentId) Then
        tmpX = 1
    Else
        tmpX = rsTemp!MaxDocumentId + 1
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
'   ATN 16/12/99
'   Removed quotes so SQL will work on SQL Server
'   WillC 13/3/00 SR3202 Added the call to replaceQuotes to handle a Doc path with a ' ie (Will's.doc)
     
         sSQL = "INSERT INTO StudyDocument (ClinicalTrialId,VersionId,DocumentId," _
        & "DocumentPath ) " _
        & "VALUES( " & ClinicalTrialId & "," & VersionId _
        & "," & tmpX & ",'" & ReplaceQuotes(sDocumentPath) & "' )"

    MacroADODBConnection.Execute sSQL
        

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                            "gdsAddTrialDocument", "TrialEdit.bas")
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
Public Sub gdsUpdateStudyDefinition(ClinicalTrialId As Long, VersionId As Integer, _
            DefaultFontColour As Long, DefaultCRFPageColour As Long, DefaultFontName As String, DefaultFontBold As Integer, _
            DefaultFontItalic As Integer, DefaultFontSize As Single)
'---------------------------------------------------------------------
Dim msSQL As String

On Error GoTo ErrHandler

    msSQL = "UPDATE StudyDefinition  " _
            & " SET StudyDefinition.DefaultFontColour = " & DefaultFontColour _
            & ", StudyDefinition.DefaultCRFPageColour = " & DefaultCRFPageColour
            
    '   ATN 23/6/99 SR 1090
    '   Check that a default font has been chosen before updating this field
    If DefaultFontName > "" Then
        msSQL = msSQL _
            & ", StudyDefinition.DefaultFontName = '" & DefaultFontName & "'"
    End If
    
    msSQL = msSQL _
            & ", StudyDefinition.DefaultFontBold = " & DefaultFontBold _
            & ",  StudyDefinition.DefaultFontItalic = " & DefaultFontItalic
    '   Check that a default font size has been chosen before updating this field
    If DefaultFontSize > 0 Then
    msSQL = msSQL _
            & ", StudyDefinition.DefaultFontSize = " & DefaultFontSize
    End If
    
    msSQL = msSQL _
            & " WHERE ClinicalTrialId = " & ClinicalTrialId _
            & " AND VersionId = " & VersionId
            
    MacroADODBConnection.Execute msSQL
        
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                            "gdsUpdateStudyDefinition", "TrialEdit.bas")
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
Public Function TrialDocumentExists(ClinicalTrialId As Long, VersionId As Integer, _
            sDocument As String) As Boolean
'---------------------------------------------------------------------
' Returns TRUE if trial document already exists
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTrialDocs As ADODB.Recordset

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM StudyDocument " _
        & " WHERE ClinicalTrialId = " & ClinicalTrialId _
        & " AND VersionId = " & VersionId
    If goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
        sSQL = sSQL & " AND NLS_LOWER(DocumentPath) = '" & ReplaceQuotes(lCase(sDocument)) & "'"
    Else
        sSQL = sSQL & " AND DocumentPath = '" & ReplaceQuotes(sDocument) & "'"
    End If
    
    Set rsTrialDocs = New ADODB.Recordset
    rsTrialDocs.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    TrialDocumentExists = (rsTrialDocs.RecordCount > 0)
    
    rsTrialDocs.Close
    Set rsTrialDocs = Nothing
    
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                            "gdsTrialDocumentList", "TrialEdit")
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
Public Sub gdsDeleteTrialDocument(ClinicalTrialId As Long, VersionId As Integer, _
            DocumentPath As String)
'---------------------------------------------------------------------
Dim msSQL As String

On Error GoTo ErrHandler

'WillC 14/3/00 Added ReplaceQuotes to handle eg Will'sDocument SR's 3234 & 3235
    msSQL = "DELETE FROM StudyDocument " _
            & "WHERE ClinicalTrialId = " & ClinicalTrialId _
            & " AND VersionId = " & VersionId _
            & " AND DocumentPath = '" & ReplaceQuotes(DocumentPath) & "'"
                        
    MacroADODBConnection.Execute msSQL

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                            "gdsDeleteTrialDocument", "TrialEdit.bas")
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
Public Function gnNewTrialPhaseId() As Integer
'---------------------------------------------------------------------
' Returns the next available id for a trial phase.
' If the table is empty, returns 1.
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim sSQL As String

On Error GoTo ErrHandler
     
    'Retrieve maximum + 1
    sSQL = " SELECT   (max(PhaseId) + 1) as NewTrialPhaseId " _
        & "FROM     TrialPhase "
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If IsNull(rsTemp!NewTrialPhaseId) Then    'if no records exist
        gnNewTrialPhaseId = 1
    Else                                'else if records exist
        gnNewTrialPhaseId = rsTemp!NewTrialPhaseId
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                            "gnNewTrialPhaseId", "TrialEdit.bas")
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
Public Function gnNewStandardFormatId() As Integer
'---------------------------------------------------------------------
' Returns the next available id for a Standard Data Format.
' If the table is empty, returns 1.
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim sSQL As String
      
On Error GoTo ErrHandler
      
    'Retrieve maximum + 1
    sSQL = " SELECT   (max(StandardDataFormatId) + 1) as NewStandardFormatId " _
        & "FROM     StandardDataFormat "
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If IsNull(rsTemp!NewStandardFormatId) Then    'if no records exist
        gnNewStandardFormatId = 1
    Else                                'else if records exist
        gnNewStandardFormatId = rsTemp!NewStandardFormatId
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                            "gnNewStandardFormatId", "TrialEdit.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
    
End Function

