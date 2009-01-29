Attribute VB_Name = "modErrors"
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   File:       modErrors.bas
'   Author:     Zulfiqar Ahmed, September 2001
'   Purpose:    New error Module for Macro 2.2 and above versions
'------------------------------------------------------------------------------------------------'
' REVISIONS
' DPH 18/10/2001 - Check To make sure gsTEMP_PATH folder exists
' NCJ 29 Aug 02 - Made separator private and renamed it
' TA 7/1/03 - code changes to display errors when they occur in the error form
' TA 22/08/2005 - new routine to ouput errors to log file
'------------------------------------------------------------------------------------------------'

Option Explicit

Private Const msLOG_SEPARATOR = "***************************************************************************"

'---------------------------------------------------------------------
Public Function MACROErrorHandler(sObjectName As String, nTrappedErrNum As Long, _
            sTrappedErrDesc As String, sProcName As String, sSource As String) As OnErrorAction
'---------------------------------------------------------------------
' Call this new error handling routine passing it the object name that cause
' the error, error number, error description, procedure name and the source
'---------------------------------------------------------------------
Dim sHTML As String

    Dim sError As String
    Dim sFileName As String
    
    On Error GoTo MemoryErr
        
    If nTrappedErrNum = 0 Then
        MACROErrorHandler = OnErrorAction.Ignore
        Exit Function
    Else
        
        sFileName = Format(Now, "ddmmyy") & ".log"
        sError = ActualError(sTrappedErrDesc)

        Call frmErrors.ProcessErrors(sObjectName, nTrappedErrNum, sError, sProcName, sSource)
        
        ErrorSource (sTrappedErrDesc)
        
        'write errors to a log file
        ' DPH 18/10/2001 - Check To make sure gsTEMP_PATH folder exists
        If FolderExistence(gsTEMP_PATH & "dummy.txt") Then
            WriteLog (gsTEMP_PATH & sFileName)
        End If
        
        frmErrors.Show vbModal
        
        MACROErrorHandler = frmErrors.gOnErrorAction
        'always unload form
        Unload frmErrors
        If MACROErrorHandler = OnErrorAction.QuitMACRO Then
            Call ExitMACRO
            Call MACROEnd
        End If
    End If
    
' SR3685 Trap for out of memory
Exit Function
MemoryErr:
    Select Case Err.Number
        Case 7 ' Out of Memory error
            MsgBox "The application has run out of memory and will now be shut down.", vbCritical, "MACRO"
            Call ExitMACRO
            Call MACROEnd
        Case Else
        
            Screen.MousePointer = vbNormal
            'display more error info
            sHTML = "Secondary error occurred while processing error" & vbCrLf
            sHTML = sHTML & "Original error: " & sTrappedErrDesc & "(" & nTrappedErrNum & ") in " & sObjectName & "." & sProcName & vbCrLf
            sHTML = sHTML & "New error: " & Err.Description & "(" & Err.Number & ") "
            
            DialogError sHTML, "Error report"

            Call ExitMACRO
            Call MACROEnd
    End Select

End Function

'---------------------------------------------------------------------
Public Sub ErrorSource(sSource As String)
'---------------------------------------------------------------------
' This routines accepts a comma delimited error string and then parse
' it to enter the object name and method name to the error details grid
'---------------------------------------------------------------------
Dim iCounter As Integer
Dim sObjectName As String
Dim sMethodName As String
Dim iPos As Integer
Dim sError As String
Dim arrErrorSource() As String
    
    'split the error source string by "|" and add it into an array,
    ' so that we can get each individual error out of it
    arrErrorSource = Split(sSource, "|")
    
    For iCounter = 1 To UBound(arrErrorSource)
        frmErrors.grdErrors.Rows = iCounter + 1
        frmErrors.grdErrors.Row = iCounter
        sError = arrErrorSource(iCounter)
        
        iPos = InStr(1, sError, ".")
        sObjectName = Left(sError, iPos - 1)
        frmErrors.grdErrors.Col = 0
        frmErrors.grdErrors.Text = sObjectName
        
        sMethodName = Right(sError, Len(sError) - iPos)
        frmErrors.grdErrors.Col = 1
        frmErrors.grdErrors.Text = sMethodName
        
    Next iCounter
End Sub

'---------------------------------------------------------------------
Public Function ActualError(sSource As String) As String
'---------------------------------------------------------------------
' Split the long comma delimited error string and get the first element
' from the array because this is the description of the error
'---------------------------------------------------------------------
Dim arrSmall() As String

    arrSmall = Split(sSource, "|")
    
    ActualError = arrSmall(0)
    
End Function

'---------------------------------------------------------------------
Public Function CreateLogFile(sFileName As String, bOverWrite As Boolean) As Boolean
'---------------------------------------------------------------------
'Synopsis: Create a log file for storing errors
'Input   : File name with full path
'Output  : Returns true if successful otherwise returns false
'---------------------------------------------------------------------

Dim objFSO As Scripting.FileSystemObject
Dim objFile As TextStream

    Set objFSO = New Scripting.FileSystemObject
    
    Set objFile = objFSO.CreateTextFile(sFileName, bOverWrite)
    
    objFile.Close
    
    Set objFSO = Nothing
    
End Function

'------------------------------------------------------------------------------'
Public Function AddErrorsToRTB(oForm As Form) As String
'------------------------------------------------------------------------------'
'Add errors from errors grid to the rich text box for printing
'------------------------------------------------------------------------------'
Dim iRow As Integer
Dim iCol As Integer
Dim sStackMessage As String

    sStackMessage = vbCrLf & "Object" & vbTab & "Method" & vbCrLf
    For iRow = 1 To oForm.grdErrors.Rows - 1
        For iCol = 0 To oForm.grdErrors.Cols - 1
            oForm.grdErrors.Row = iRow
            oForm.grdErrors.Col = iCol
            sStackMessage = sStackMessage & oForm.grdErrors.Text & vbTab
        Next iCol
        sStackMessage = sStackMessage & vbCrLf
    Next iRow
    AddErrorsToRTB = sStackMessage
End Function

'------------------------------------------------------------------------------'
Public Function FileAlreadyExists(sFile As String) As Boolean
'------------------------------------------------------------------------------'
' Check if specified file already exists
'------------------------------------------------------------------------------'
Dim objFSO As Scripting.FileSystemObject

    Set objFSO = New Scripting.FileSystemObject
    
    FileAlreadyExists = objFSO.FileExists(sFile)
    
    Set objFSO = Nothing
End Function

'------------------------------------------------------------------------------'
Public Sub WriteLog(sFile As String)
'------------------------------------------------------------------------------'
' Write errors to the log file. Check if the file already exists, if not create
' one. Returns true if successful otherwise it returns false.
'------------------------------------------------------------------------------'
Dim objFile As TextStream
Dim objFSO As Scripting.FileSystemObject
Dim bFileCheck As Boolean
Dim sFolderName As String
    
    Set objFSO = New Scripting.FileSystemObject
    
    bFileCheck = FileAlreadyExists(sFile)
    
    If bFileCheck Then
        Set objFile = objFSO.OpenTextFile(sFile, ForAppending)
    Else
        Set objFile = objFSO.CreateTextFile(sFile, True)
    End If
    
    objFile.WriteLine frmErrors.rtbErrMsg.Text & AddErrorsToRTB(frmErrors)
    objFile.WriteLine msLOG_SEPARATOR
    
    objFile.Close
    
    Set objFSO = Nothing
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub WriteLogError(sSource As String, sMessage As String)
'----------------------------------------------------------------------------------------'
'TA 22/08/2005: log all errors to file
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
Dim objFile As TextStream
Dim objFSO As Scripting.FileSystemObject
Dim sFileName As String

    'ignore all errors
    On Error GoTo ErrorLabel
    
    sFileName = gsTEMP_PATH & Format(Now, "ddmmyy") & ".log"
    If FolderExistence(gsTEMP_PATH & "dummy.txt") Then
        Set objFSO = New Scripting.FileSystemObject
        If FileAlreadyExists(sFileName) Then
            Set objFile = objFSO.OpenTextFile(sFileName, ForAppending)
        Else
            Set objFile = objFSO.CreateTextFile(sFileName, True)
        End If
        objFile.WriteLine ""
        objFile.WriteLine Format(Now, "HH:MM:SS") & " - ERROR OCCURRED IN " & sSource
        objFile.WriteLine Replace(sMessage, "|", vbCrLf)
        objFile.WriteLine ""
        objFile.WriteLine msLOG_SEPARATOR
        objFile.Close
        Set objFSO = Nothing
    End If
    
ErrorLabel:
End Sub
