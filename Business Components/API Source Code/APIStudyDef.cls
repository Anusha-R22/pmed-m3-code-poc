VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APIStudyDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------
' File: APISudyDef.cls
' Copyright: InferMed Ltd 2005 All Rights Reserved
' Author: Toby Aldridge, February 2005
' Purpose: MACRO interface, to update study defintions
'------------------------------------------------------------------
' REVISIONS
'------------------------------------------------------------------

'NOT CURENTLY USED
Option Explicit


Private moALM As ALM5
Private m_clmGuideline As Guideline
'---------------------------------------------------------------------
Public Function UpdateQuestionDerivation(ByRef sSerialisedUser As String, _
                                    ByVal lStudyId As Long, sStudyName As String, _
                                    ByVal lQuestionId As Long, _
                                    sDerivation As String) As Boolean
'---------------------------------------------------------------------
' Update he derivation expression for a question
'---------------------------------------------------------------------
Dim oCon As ADODB.Connection
Dim sTempPath As String
Dim oUser As MACROUser
Dim sKey As String
Dim clmDItem As DataItem

    On Error GoTo ErrLabel
    
    UpdateQuestionDerivation = False
    
    ' Create the MACRO User
    Set oUser = New MACROUser
    Call oUser.SetStateHex(sSerialisedUser)
       
    ' Get settings file (TRUE means look one level up because we're in a DLL)
    Call InitialiseSettingsFile(True)
    sTempPath = GetMACROSetting("Temp", App.Path & "\..\Temp\")

    'start clm
    
    StartUpCLM oUser.CurrentDBConString, sTempPath, lStudyId


    oCon.Open oUser.CurrentDBConString
    oCon.CursorLocation = adUseClient
    
    LoadProformaTrial oCon, sStudyName
    oCon.BeginTrans
    
    oCon.Execute "Update DATAITEM set derivation=" & SQL_ValueToStringValue(sDerivation) & " where Clinicaltrialid =" & lStudyId & " and dataitemid =" & lQuestionId

    
    'update arezzo
    sKey = CStr(lQuestionId)
    Set clmDItem = m_clmGuideline.colDataItems.Item(sKey)
    clmDItem.Derivation = RTrim(sDerivation)
    
    'save this
    SaveCLMGuideline oCon, sStudyName
    
    oCon.CommitTrans
    
    oCon.Close
    
    ShutDownCLM
    
    Set oUser = Nothing
    
    UpdateQuestionDerivation = True
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|APISudyDef.UpdateQuestionDerivation"
End Function


'------------------------------------------------------------------
Private Sub LoadProformaTrial(oCon As Connection, ByVal sClinicalTrialName As String)
'------------------------------------------------------------------
' Load the definition of the given clinical trial into the CLM
' Assume the trial does exist
'------------------------------------------------------------------
Dim sSQL As String
Dim oRS As ADODB.Recordset


    On Error GoTo ErrHandler
    
    Set oRS = New ADODB.Recordset
    
    sSQL = "SELECT ArezzoFile FROM Protocols WHERE FileName = '" & sClinicalTrialName & "'"
    oRS.Open sSQL, oCon, adOpenKeyset, adLockOptimistic
    
    moALM.ArezzoFile = oRS.fields("ArezzoFile").Value

    Set oRS = Nothing
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|APISudyDef.LoadProformaTrial"
    
End Sub


'---------------------------------------------------------
Private Sub SaveCLMGuideline(oCon As Connection, sProtocolName)
'---------------------------------------------------------
' Save the current version of Arezzo guideline to the PSS
' NCJ 11 Aug 99
' NCJ 22 Feb 00 - Only save if gbDoCLMSave is TRUE
'---------------------------------------------------------
Dim sSQL As String
Dim rsProtocols As ADODB.Recordset
        sSQL = "SELECT * FROM Protocols WHERE FileName = '" & sProtocolName & "'"
        Set rsProtocols = New ADODB.Recordset
        rsProtocols.Open sSQL, oCon, adOpenKeyset, adLockOptimistic, adCmdText
        'Update Protocols Arezzo file and Modifed columns now
        rsProtocols.fields("ArezzoFile").Value = moALM.ArezzoFile
        rsProtocols.fields("Modified").Value = SQL_ValueToStringValue(Now)
        rsProtocols.Update
        rsProtocols.Close
        Set rsProtocols = Nothing
    
    
End Sub

'---------------------------------------------------------
Private Sub StartUpCLM(sConnection As String, sTempPath As String, Optional lStudyId As Long = 0)
'---------------------------------------------------------
' Start up the CLM
'---------------------------------------------------------
Dim sR As String
Dim sFile As String
Dim oArezzoMemory As clsAREZZOMemory

    ' Remove any previous one
    Call ShutDownCLM
    
    Set moALM = New ALM5
    
    Set oArezzoMemory = New clsAREZZOMemory
    'ASH 25/11/2002 StudyId = 0 gets the maximum settings value from the database
    'NCJ 16 Jan 03 - StudyId = -1 gets default settings
    ' NCJ 28 Jan 03 - Pass in DB connection string
    Call oArezzoMemory.Load(lStudyId, sConnection)
    
    Call moALM.StartALM(oArezzoMemory.AREZZOSwitches)
    
    ' NCJ 26/1/00 - Set temp directory for CLM
    moALM.TempDirectory = sTempPath
    
    ' NCJ 1/2/01 - Load cycling data add-on
    'REM 08/07/02 - Changed name from Macro_Arezzo.pc to Macro3_Arezzo.pc
    sFile = App.Path & "\..\Macro3_Arezzo.pc"
    Call moALM.GetPrologResult("ensure_loaded('" & sFile & "'), write('0000'). ", sR)
 
    Set m_clmGuideline = moALM.Guideline
 
    Set oArezzoMemory = Nothing

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|APIStudyDef.StartUpCLM"
    
End Sub

'---------------------------------------------------------
Private Sub ShutDownCLM()
'---------------------------------------------------------
' MM 9/8/99
' Close down ALM.DLL if not already done
'---------------------------------------------------------
    
    
  
    If Not m_clmGuideline Is Nothing Then
        m_clmGuideline.Clear
        Set m_clmGuideline = Nothing
    End If
    If Not moALM Is Nothing Then
        moALM.CloseALM
        Set moALM = Nothing
    End If
     
End Sub
