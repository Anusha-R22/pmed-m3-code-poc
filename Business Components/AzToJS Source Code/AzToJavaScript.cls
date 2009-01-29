VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AzToJavaScript"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------
'
'   File: AzToJavaScript.cls
'   Copyright InferMed 2001 All Rights Reserved
'   Author: Nicky Johns
'   Purpose: Conversion of Arezzo terms to JavaScript for MACRO
'
'------------------------------------------------------------
' REVISIONS:
'   NCJ 10 Sep 01 - Initial development
'   NCJ 11 Sep 01 - Changed ConvertTerms to return array rather than collection
'   NCJ 24 Sep 01 - Changed separator expected from Prolog from | to `
'   NCJ 8 Oct 01 - Return an error string from Initialise
'   NCJ 9 Oct 01 - ConvertTerms now returns a Variant
'   NCJ 30 Aug 02 - Conversion to ALM5 and PfToJavaScript3
'   NCJ 1 Oct 02 - Allow a "subsidiary" eForm to be specified in ConvertTerms (to handle MACRO's Visit eForms)
'------------------------------------------------------------

Option Explicit

Public Enum eAzTermType
    Derivation = 1
    CollectIf = 2
    Validation = 3
    ValidationMsg = 4
End Enum

' The fixed error token used to denote some sort of Prolog error
Private Const msErrorToken = "||"

' Our copy of the ALM
Private moALM As ALM5

' Our term collection
Private mcolTerms As Collection

' Are we doing expressions or conditions?
Private mbExpressions As Boolean

'------------------------------------------------------------
Public Function Initialise(oALM As ALM5) As String
'------------------------------------------------------------
' Assume the ALM already started and initialised.
' Returns an error string resulting from the loading of the Prolog file (empty if OK)
'------------------------------------------------------------
Dim sFile As String
Dim sGoal As String
Dim sR As String
Dim sPrologErr As String
    
    On Error GoTo ErrLabel
    
    Set moALM = oALM
    
    ' Load the PfToJavaScript file (ignore result)
    ' NCJ 30 Aug 02 - Changed to use PfToJavaScript3.pc for MACRO 3.0
    sFile = App.Path & "\PfToJavaScript3.pc"
    sGoal = "ensure_loaded('" & sFile & "'), write('0000'). "
    sPrologErr = moALM.GetPrologResult(sGoal, sR)
    Initialise = sPrologErr
 
    Exit Function
    
ErrLabel:
    Initialise = "AzToJavaScript Initialisation error " & Err.Number & ", " & Err.Description
    
End Function

'------------------------------------------------------------
Public Sub InitTerms(ByVal nTermType As eAzTermType)
'------------------------------------------------------------
' Initialise our term collection with the given type
'------------------------------------------------------------

    Select Case nTermType
    Case eAzTermType.Derivation, eAzTermType.ValidationMsg
        ' These are expressions
        mbExpressions = True
    Case eAzTermType.CollectIf, eAzTermType.Validation
        ' These are conditions
        mbExpressions = False
    End Select
    
    ' Initialise our terms collection
    Set mcolTerms = New Collection
    
End Sub

'------------------------------------------------------------
Public Sub AddTerm(ByVal sQCode As String, ByVal sAzTerm As String)
'------------------------------------------------------------
' Add a term for this question to our collection
'------------------------------------------------------------
Dim oTerm As AzTerm

    Set oTerm = New AzTerm
    oTerm.QuestionCode = LCase(sQCode)
    oTerm.ArezzoTerm = sAzTerm
    mcolTerms.Add oTerm
    
    Set oTerm = Nothing

End Sub

'------------------------------------------------------------
Public Function ConvertTerms(ByVal lMainEFormTaskId As Long, Optional lSubsidiaryEFormTaskId As Long = 0) As Variant
'------------------------------------------------------------
' Convert all the terms stored to JavaScript
' Return an array of results in the same order as they were added
' lMainEFormTaskId identifies the eForm instance to which these terms belong
' lSubsidiaryEFormTaskId optionally identifies an additional eForm instance to which local data names might belong
'   (this is to support Visit eForms in MACRO 3.0)
'------------------------------------------------------------
' NCJ 24 Sep 01 - Changed separator in Split from | to `
' NCJ 9 Oct 01 - Made vJSTerms and ConvertTerms a variant
' NCJ 1 Oct 02 - Added lSubsidiaryEFormTaskId parameter
'------------------------------------------------------------
Dim oTerm As AzTerm
Dim sPrologList As String
Dim sComma As String
Dim sGoal As String
Dim sResults As String
Dim sR As String
Dim nTermCount As Integer
Dim vJSTerms As Variant
Dim i As Long

    On Error GoTo ErrLabel
    
    ' Initialise results
'    Set colResults = New Collection
    nTermCount = mcolTerms.Count
    If nTermCount > 0 Then
        ' Create Prolog list of QName(`Term`) terms
        sPrologList = "["       ' Start list
        sComma = ""
        For Each oTerm In mcolTerms
            sPrologList = sPrologList & sComma & oTerm.QuestionCode & "(`" _
                            & oTerm.ArezzoTerm & "`)"
            sComma = ","
        Next
        sPrologList = sPrologList & "]"     ' End list
        
        ' Have we got expressions or conditions?
        If mbExpressions Then
            sGoal = "vb_exprs_as_jscript"
        Else
            sGoal = "vb_conds_as_jscript"
        End If
        sGoal = sGoal & "( " & sPrologList & ", " & lMainEFormTaskId & ", " & lSubsidiaryEFormTaskId & " ). "
        ' Get results as delimited string
        sResults = moALM.GetPrologResult(sGoal, sR)
        
        If sR = "0000" Then     ' Was the result code OK?
            ' Now peel off the results collection into a string array
            vJSTerms = Split(sResults, "`")
        Else
            ' Return special 2-element array, the first element being our fixed error token
            ' and the second element being what Prolog said about it all
            vJSTerms = Array(msErrorToken, sResults)
        End If
        
    End If

    ' Return results
    ConvertTerms = vJSTerms
    
    ' Tidy up
    Set oTerm = Nothing
    
    Exit Function

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|AzToJavaScript.ConvertTerms"

End Function

'------------------------------------------------------------
Private Sub Class_Initialize()
'------------------------------------------------------------
'------------------------------------------------------------

    Set mcolTerms = Nothing
    Set moALM = Nothing
    
End Sub

'------------------------------------------------------------
Private Sub Class_Terminate()
'------------------------------------------------------------
'------------------------------------------------------------

    Set mcolTerms = Nothing
    Set moALM = Nothing

End Sub
