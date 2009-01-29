Attribute VB_Name = "basDocuments"
'---------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       basDocuments.bas
'   Author:         Andrew Newbigging, June 1997
'   Purpose:    Routines for launching external documents using Windows shell.
'---------------------------------------------------------------------------------------------'
'---------------------------------------------------------------------------------------------'
'   Revisions:
 '  1               Andrew Newbigging   4/06/97
'   2       Nap         4/06/97
'   3       Andrew Newbigging   13/07/98
'----------------------------------------------------------------------------------------------'
'-----------------------------------------------------------
' MODULE:   Documents
'
' Common routines for launching documents.
'
' The Shell Function Help topic implies in the description
' of the 'pathname' argument that it will launch documents
' with their associated applications:
' "pathname   Name of the program to execute and any required
' arguments or command line switches; may include directory
' and drive.  May also be the name of a document that has been
' associated with an executable program."
'
' In fact, Shell does not have the capability of launching
' files based on their file association.
' In order to launch a document using its file association
' use the ShellExecute Windows API call.
'
' Requires:
'   COMMON.BAS
'
'-----------------------------------------------------------
Option Explicit
Option Compare Binary
Option Base 0

#If Win32 Then
    Private Declare Function ShellExecute Lib _
        "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long

#Else
    Private Declare Function ShellExecute Lib _
        "shell.dll" (ByVal _
        hWnd As Integer, ByVal lpOperation _
        As String, ByVal lpFile _
        As String, ByVal lpParameters _
        As String, ByVal lpDirectory _
        As String, ByVal nShowCmd _
        As Integer) As Integer
#End If

Private Const SW_SHOWNORMAL = 1

'SDM 15/12/99 SR1214
Private Const ERROR_SUCCESS = 32&
Private Const ERROR_NO_ASSOC = 31&
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&


'-----------------------------------------------------------
' SUB:  ShowDocument
'
' Checks that the document exists and launches the
' appropriate application to view it.
'
' Assumes that the startup directory can be the directory
' of the parent application.
'
'-----------------------------------------------------------
Public Sub ShowDocument(ByRef rWindowHandle As Long, _
                        ByRef rDocumentPath As String)

' rWindowHandle is the handle of the window in the parent
' application from which the document is launched
On Error GoTo ErrHandler

Dim mnReturnCode As Long
Dim varTaskID As Variant

'If FileExists(rDocumentPath) Then

'Screen.MousePointer = vbHourglass

    mnReturnCode = ShellExecute(rWindowHandle, _
                                vbNullString, _
                                rDocumentPath, _
                                vbNullString, _
                                gsAddDirSep(App.Path), _
                                SW_SHOWNORMAL)


    'SDM 15/12/99 SR1214
    If mnReturnCode < ERROR_SUCCESS Then
        Select Case mnReturnCode
            Case ERROR_NO_ASSOC:
                'Try the OpenWith dialog
                'NOTE:  Do not try and clean the code between the quotes.
                '       A space after the comma will cause a crash.
                varTaskID = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & rDocumentPath, SW_SHOWNORMAL)
                mnReturnCode = (varTaskID <> 0)
            Case ERROR_OUT_OF_MEM:
                MsgBox "Error: Out of Memory/Resources.", vbInformation
            Case ERROR_FILE_NOT_FOUND:
                MsgBox "Error: File not found.", vbInformation
            Case ERROR_PATH_NOT_FOUND:
                MsgBox "Error: Path not found.", vbInformation
            Case ERROR_BAD_FORMAT:
                MsgBox "Error: Bad File Format.", vbInformation
            Case Else:
        End Select
    End If


Screen.MousePointer = vbDefault

'Else

'    MsgBox "Document " & rDocumentPath & " could not be found.", _
        vbOKOnly
        
'End If

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "ShowDocument", "basDocuments")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

