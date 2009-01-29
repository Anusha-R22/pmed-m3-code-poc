Attribute VB_Name = "modMIMsgImpExp"
'----------------------------------------------------------------------------------------'
' File:         modMIMsgImpExp.bas
' Copyright:    InferMed Ltd. 2003. All Rights Reserved
' Author:       Richard Meinesz September 2003
' Purpose:      Import/Export of MIMessages and LFMessages for the MACRO 3.0 Utilities module
'----------------------------------------------------------------------------------------'
' Revisions:
'   NCJ 28 Oct 03 - Added file header; corrected some message text errors; added sMsg var to ExportMIMsgLFMsgZIP
'----------------------------------------------------------------------------------------'

Option Explicit

Public Const gsEXPORT_MILFMSG_ZIP = "ExportMIMsgLFMsgZIP"

'--------------------------------------------------------------------
Public Function ExportMIMessagesLFMessages(sTrialName As String, sSite As String, sPersonId As String, sZipFileName As String) As String
'--------------------------------------------------------------------
'REM 18/07/03
'Routine to export MIMessages and LFMessages, creates XML files of the data
'--------------------------------------------------------------------
Dim asZIPFileList(2) As String
Dim bValidMACROVersion As Boolean
Dim sMIMsg As String
Dim sLFMsg As String
    
    On Error GoTo ErrLabel
    
    sMIMsg = ExportMIMessage(sTrialName, sSite, sPersonId)
    sLFMsg = ExportLFMessages(sTrialName, sSite, sPersonId)
    
    asZIPFileList(0) = sMIMsg
    asZIPFileList(1) = sLFMsg
    
    If (sMIMsg = "") And (sLFMsg = "") Then ' don't create file header or add to zip file
        Exit Function
    End If
    
    asZIPFileList(2) = FileHeader(sTrialName, sSite, sPersonId)
    ExportMIMessagesLFMessages = ExportMIMsgLFMsgZIP(sZipFileName, asZIPFileList)
    
Exit Function
ErrLabel:
    ExportMIMessagesLFMessages = "MIMessage and LFMessage Export Aborted." & vbCrLf & _
                                 "Error code " & Err.Number & " - " & Err.Description & "."
End Function

'---------------------------------------------------------------------
Public Function ExportMIMsgLFMsgZIP(ByVal sZipFileName As String, asZIPFileList() As String) As String
'---------------------------------------------------------------------
'REM 18/07/03
'Routine adds the MIMsg and LFMsg files to the patient data zip file
'---------------------------------------------------------------------
Dim sZIPfile As String
Dim i As Long
Dim sFile As String
Dim sMsg As String

    On Error GoTo ErrHandler

    gLog gsEXPORT_MILFMSG_ZIP, "Session adding MIMessage and LFMessage files to " & sZipFileName & " starting"
    
    ' Create ZIP file path
    sZIPfile = gsOUT_FOLDER_LOCATION & sZipFileName & ".zip"
    
    ' Make sure folder exists before opening
    If FolderExistence(sZIPfile) Then
            
        ' Error handler for zip process
        On Error GoTo ZipErr
        
        ' Add Files to existing patient ZIP file
        Call ZipFiles(asZIPFileList, sZIPfile)
    
        On Error GoTo ErrHandler
        
        'Kill off the files that have been compacted into a ZIP file
        For i = 0 To UBound(asZIPFileList)
            sFile = asZIPFileList(i)
            If sFile <> "" Then
                Kill sFile
            End If
            gLog gsEXPORT_MILFMSG_ZIP, "Removing " & StripFileNameFromPath(asZIPFileList(i))
        Next i
                
        gLog gsEXPORT_MILFMSG_ZIP, "Session " & sZipFileName & " completed"
        ExportMIMsgLFMsgZIP = ""
    Else
        sMsg = "Session " & sZipFileName & " failed as could not create ZIP file directory " & gsOUT_FOLDER_LOCATION
        gLog gsEXPORT_MILFMSG_ZIP, sMsg
        ExportMIMsgLFMsgZIP = sMsg
'        gLog gsEXPORT_MILFMSG_ZIP, "Session " & sZipFileName & " failed as could not create ZIP file directory " & gsOUT_FOLDER_LOCATION
'        ExportMIMsgLFMsgZIP = "Session " & sZipFileName & " failed as could not create ZIP file directory " & gsOUT_FOLDER_LOCATION
    End If

Exit Function
ZipErr:
    gLog gsEXPORT_MILFMSG_ZIP, "Session " & StripFileNameFromPath(sZipFileName) & " failed. XCeed Error Number " & Err.Number
    ExportMIMsgLFMsgZIP = "Session " & StripFileNameFromPath(sZipFileName) & " failed. XCeed Error Number " & Err.Number
Exit Function
ErrHandler:
    sMsg = "Error while creating MIMsgLFmsg ZIP file." & vbCrLf & _
                          "Error code " & Err.Number & " - " & Err.Description & "."
    gLog gsEXPORT_MILFMSG_ZIP, sMsg
    ExportMIMsgLFMsgZIP = sMsg
End Function


'--------------------------------------------------------------------
Private Function FileHeader(sTrialName As String, sSite As String, sPersonId As String) As String
'--------------------------------------------------------------------
'REM 18/07/03
'Create a file containing all the information about the exported MIMessages and LFMessages
'--------------------------------------------------------------------
Dim sFileHeader As String
Dim sFile As String
Dim sSQL As String
Dim rsVersion As ADODB.Recordset
Dim sMACROVersion As String
Dim sVersion As String

    On Error GoTo ErrLabel
    
    sSQL = "SELECT * FROM MACROControl"
    Set rsVersion = New ADODB.Recordset
    rsVersion.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    
    sMACROVersion = rsVersion!MACROVersion
    sVersion = sMACROVersion & "." & rsVersion!BuildSubVersion
        
    sFileHeader = "File Header:" & vbCrLf & vbCrLf & "TrialName = " & sTrialName & vbCrLf & "Site = " & sSite & vbCrLf _
                & "SubjectId = " & sPersonId & vbCrLf & "MACRO Version = " & sVersion & vbCrLf & "Export Time = " & Now
    
    sFile = gsOUT_FOLDER_LOCATION & "FileHeader.txt"
    StringToFile sFile, sFileHeader
    
    FileHeader = sFile
    
    rsVersion.Close
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modMIMsgImpExp.FileHeader"
End Function

'--------------------------------------------------------------------
Private Function ExportMIMessage(sTrialName As String, sSite As String, sPersonId As String) As String
'--------------------------------------------------------------------
'REM 18/07/03
'Create the Exported MIMessage xml file
'--------------------------------------------------------------------
Dim sSQL As String
Dim rsExpMIMsg As ADODB.Recordset
Dim sFile As String

    On Error GoTo ErrLabel
    
    sSQL = "SELECT * FROM MIMessage " _
         & "WHERE MIMessageTrialName = '" & sTrialName & "'"
    
    If sSite <> "All Sites" Then
        sSQL = sSQL & " AND MIMESSAGESITE = '" & sSite & "'"
    End If
    
    If sPersonId <> "All Subjects" Then
        sSQL = sSQL & " AND MIMESSAGEPERSONID = " & CLng(sPersonId)
    End If
    
    Set rsExpMIMsg = New ADODB.Recordset
    rsExpMIMsg.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    sFile = ""
    'if records are returned then create the xml file
    If rsExpMIMsg.RecordCount <> 0 Then
        sFile = gsOUT_FOLDER_LOCATION & "MIMessages_" & sSite & "_" & sPersonId & ".xml"
        rsExpMIMsg.Save sFile, adPersistXML
    End If
    
    rsExpMIMsg.Close
    
    ExportMIMessage = sFile
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modMIMsgImpExp.ExportMIMessage"
End Function

'--------------------------------------------------------------------
Private Function ExportLFMessages(sTrialName As String, sSite As String, sPersonId As String) As String
'--------------------------------------------------------------------
'REM 18/07/03
'Create the Exported LFMessages xml file
'--------------------------------------------------------------------
Dim sSQL As String
Dim rsExpLFMsg As ADODB.Recordset
Dim sFile As String

    On Error GoTo ErrLabel
    
    sSQL = "SELECT * FROM LFMessage " _
         & "WHERE CLINICALTRIALNAME = '" & sTrialName & "'"
    
    If sSite <> "All Sites" Then
        sSQL = sSQL & " AND TRIALSITE = '" & sSite & "'"
    End If
    
    If sPersonId <> "All Subjects" Then
        sSQL = sSQL & " AND PERSONID = " & CInt(sPersonId)
    End If
    
    Set rsExpLFMsg = New ADODB.Recordset
    rsExpLFMsg.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    sFile = ""
    'If records are returned then create the xml file
    If rsExpLFMsg.RecordCount <> 0 Then
        sFile = gsOUT_FOLDER_LOCATION & "LFMessages_" & sSite & "_" & sPersonId & ".xml"
        rsExpLFMsg.Save sFile, adPersistXML
    End If
    
    rsExpLFMsg.Close
    
    ExportLFMessages = sFile

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modMIMsgImpExp.ExportLFMessages"
End Function

'--------------------------------------------------------------------
Public Function ImportMIMessageLFMessage() As String
'--------------------------------------------------------------------
'REM 24/07/03
'Import the MIMessage and LFMessage data
'--------------------------------------------------------------------
Dim sXMLFile As String
Dim sFileName As String
Dim i As Integer
Dim rs As ADODB.Recordset
Dim vImport As Variant
Dim sFileHeader As String
Dim vFileHeader As Variant
Dim sMACROVersion As String
Dim vMACROVersion As Variant
Dim sImpXML As String

'TODO - check file version, validity of data etc
    On Error GoTo ErrLabel
    
    'check to see if the FileHeader text file exits, if it doesn't do nothing as there were no MIMessages or LFMessages in the import file
    If FolderExistence(gsCAB_EXTRACT_LOCATION & "FileHeader.txt", True) Then
        'get the text string from the file
        sFileHeader = StringFromFile(gsCAB_EXTRACT_LOCATION & "FileHeader.txt")
        'split it on carrage returns
        vFileHeader = Split(sFileHeader, vbCrLf)
        'return the version number row
        sMACROVersion = vFileHeader(5)
        'get the version number after the = sign
        vMACROVersion = Split(sMACROVersion, "=")
        
        sMACROVersion = vMACROVersion(1)
        
        Kill gsCAB_EXTRACT_LOCATION & "FileHeader.txt"
        
        'check to see if the import file version number matches the current MACRO version number
        If Trim(sMACROVersion) = ("3.0." & CURRENT_SUBVERSION) Then
        
            sFileName = "MIMessages"
            For i = 1 To 2
                'check if there is a MIMessage xml file
                sXMLFile = Dir(gsCAB_EXTRACT_LOCATION & sFileName & "*.xml")
                
                'Check that there is a file to extract data from
                If sXMLFile <> "" Then
                    
                    'Create a recordset from XML file
                    Set rs = New ADODB.Recordset
                    rs.Open gsCAB_EXTRACT_LOCATION & sXMLFile
                    
                    'If there are records then import them
                    If rs.RecordCount <> 0 Then
                        
                        vImport = rs.GetRows
                        'if sImpXML is "" then import was successful, otherwise will return error message
                        sImpXML = ImportXML(vImport, sFileName)
                        
                        'remove the file from the CabExtract folder
                        Kill gsCAB_EXTRACT_LOCATION & sXMLFile
                        
                        'if it was not successfully imported then exit sub and return error message
                        If sImpXML <> "" Then
                            ImportMIMessageLFMessage = sImpXML
                            
                            'EXIT FUNCTION IF THERE WAS AN ERROR
                            Exit Function
                        End If
                    
                    End If
                    
                End If
                sFileName = "LFMessages"
            Next
    
            ImportMIMessageLFMessage = ""
        Else
            ImportMIMessageLFMessage = "Import file's MACRO version does not match current MACRO version. Cannot import MIMessages and LFMessages."
        End If
    End If
    
Exit Function
ErrLabel:
    ImportMIMessageLFMessage = "Error occurred while importing MIMessages and LFMessages. Error Description: " & Err.Description & " ,Error Number: " & Err.Number
End Function

'--------------------------------------------------------------------
Private Function ImportXML(vImport As Variant, sFileName As String) As String
'--------------------------------------------------------------------
'REM 16/09/03
'Rotine to import the XML string containing the MIMessage and LFMessage data
'--------------------------------------------------------------------
Dim sSQL As String
Dim i As Integer
Dim j As Integer
Dim lMIMsgId As Long
Dim sMIMsgSite As String
Dim nMIMsgSource As Integer
Dim rsImport As ADODB.Recordset
Dim sLFTrialName As String
Dim lLFTrialId As Long
Dim sLFSite As String
Dim lLFPersonId As Long
Dim nLFSource As Integer
Dim lLFMsgId As Long
    
    On Error GoTo ErrLabel
    
    If sFileName = "MIMessages" Then
        For i = 0 To UBound(vImport, 2)
            'get primary key fields to check for conflict
            lMIMsgId = vImport(0, i)
            sMIMsgSite = vImport(1, i)
            nMIMsgSource = vImport(2, i)
            
            sSQL = "SELECT * FROM MIMessage " _
                & " WHERE MIMESSAGEID = " & lMIMsgId _
                & " AND MIMESSAGESITE = '" & sMIMsgSite & "'" _
                & " AND MIMESSAGESOURCE = " & nMIMsgSource
            Set rsImport = New ADODB.Recordset
            rsImport.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            'if no conflict then insert new row
            If rsImport.RecordCount = 0 Then
                rsImport.AddNew
            End If
                
            For j = 0 To rsImport.Fields.Count - 1
                
                rsImport.Fields(j) = vImport(j, i)
                    
            Next
            rsImport.Update
            
        Next
        
        rsImport.Close
        
    ElseIf sFileName = "LFMessages" Then
            
        For i = 0 To UBound(vImport, 2)
            
            
            sLFTrialName = vImport(0, i)
            lLFTrialId = goUser.Studies.StudyByName(sLFTrialName).StudyId
            sLFSite = vImport(2, i)
            lLFPersonId = vImport(3, i)
            nLFSource = vImport(4, i)
            lLFMsgId = vImport(5, i)
            
            sSQL = "SELECT * FROM LFMessage " _
                & "WHERE CLINICALTRIALNAME = '" & sLFTrialName & "'" _
                & " AND CLINICALTRIALID = " & lLFTrialId _
                & " AND TRIALSITE = '" & sLFSite & "'" _
                & " AND PERSONID = " & lLFPersonId _
                & " AND SOURCE = " & nLFSource _
                & " AND MESSAGEID = " & lLFMsgId
            Set rsImport = New ADODB.Recordset
            rsImport.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockPessimistic, adCmdText

            'if no conflict then insert new row
            If rsImport.RecordCount = 0 Then
                rsImport.AddNew
            End If
                
            For j = 0 To rsImport.Fields.Count - 1
                
                If j = 1 Then 'TrialId needs to be replaced with current DB's trialId
                    rsImport.Fields(j) = lLFTrialId
                Else
                    rsImport.Fields(j) = vImport(j, i)
                End If
            Next
            rsImport.Update
            
        Next
        
        rsImport.Close

    End If

    ImportXML = ""

Exit Function
ErrLabel:
    ImportXML = "Error occurred while importing MIMesssage and LFMessage XML data. Error Description: " & Err.Description & " ,Error Number: " & Err.Number
End Function
