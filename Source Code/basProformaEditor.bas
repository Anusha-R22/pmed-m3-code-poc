Attribute VB_Name = "basProformaEditor"
'---------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998-2006. All Rights Reserved
'   File:       basProformaEditor.bas
'   Author:     Andrew Newbigging, September 1997
'   Purpose:    Routines to update the Proforma database and maintain it in parallel with
'   the MACRO database.
'----------------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------------'
'   Revisions:
'  1-18   Andrew Newbigging et al.      10/09/97 - 05/08/99
'
'   19  Nicky Johns 09-13 Aug 99
'       Major changes for migration to Arezzo CLM
'   20  Paul Norris             19/08/99    Replaced MasgBox to MsgBox
'   21  NCJ, 23 Aug 99
'       Bug fix to MTMToArezzoDataType
'       Set up top level plan when loading a trial
'   22  NCJ 26 Aug 99
'       Use global strings for plan/action name prefixes
'   23  NCJ 27 Aug 99
'       Set cycling info for CRF pages
'   24  NCJ 7 Sep 99
'       Changed task prefixes to task Captions instead
'       NCJ 8 Sep 99
'       Lock top level plan in trial
'        REMOVED NewLPAQueryOnce, NewLPACallGoal, NewLPAExitGoal, DeleteProformaCRFElement,
'               CreateDataCycleTasks, gnNewDataTag, gnNewNodeTag, UpdateProformaDataValidation,
'               DeleteProformaInstance, DeleteProformaTask, InitialiseDataItemValidation,
'               PassDataDefinitionsToProlog, WriteProformaDataDefinitions, InsertProformaCRFElement
'
'   25  NCJ 10 Sep 99
'       UniqueVisitFormDataitemCode moved here from Main and
'       changed to use new CLM call
'   NCJ 17 Sept 99
'       Moved PSS stuff out to separate module modPSSSupport
'   NCJ 22 Sept 99
'       Set cycling tasks to "optional" to ensure correct behaviour at runtime
'   NCJ 27 Sept 99
'       Use new CLM2
'   WillC 10/11/99  Added the error handlers
'   NCJ 13 Dec 99 - Ids to Long
'   NCJ 14 Jan 00 - Revised calculation of plan coords when adding to Arezzo
'                   Ensure Longs are used, to avoid overflow errors
'   NCJ 26 Jan 00 - New CLM (1.0.9) which allows setting of temp folder
'   NCJ 27 Jan 00 - UniqueVisitFormDataitemCode now displays its own message
'   NCJ 3 Feb 00 - Use CLM to do date format validations
'   NCJ 25/9/00 - Added ReadValidDate function (uses new routine in CLM3)
'   NCJ 26/9/00 - Added ConvertDateFromArezzo
'   NCJ 2 Jan 01 - Added handling for collections of Categories and Warning Conditions
'   NCJ 16/2/2001 Added study "integrity check" routine (see CheckStudyHoldsWater)
'   NCJ 17 Apr 01 - IsValidTrialName
'   NCJ 17 May 01 - Changed to use new ALM4 (Arezzo Logic Module)
'   NCJ 29 Jun 01 - Changed Prolog switches in StartUpCLM
'   ATO 12 July 01 - Added CopyProformaTrial
'   NCJ 18 Oct 01 - Changed StartUpCLM to use new GetPrologSwitches
'   REM 08/07/02 - Changed name from Macro_Arezzo.pc to Macro3_Arezzo.pc in StartUpCLM
'   NCJ 29 Aug 02 - Changed from ALM4 to ALM5
'   NCJ 10 Sep 02 - Added SetVisitRepeats and SetEFormRepeats
'   ZA 23/09/2002 - Replaced PSSProtocols object with SQL
'   NCJ 25 Sept 02 - Tidied up SQL for Protocols table
'   NCJ 10 Oct 02 - In CopyProformaTrial, must rename top level plan in Arezzo File
'   ASH 15/11/2002 - Fix for error ORA-01704: string literal too long in sub SaveCLMGuideline
'   NCJ 29 Nov 02 - Fixed bug in CopyProformaTrial, and made sure Protocols always selected by FileName
'   NCJ 6 Mar 03 - Store Protocol name as msProtocolName to avoid problems with copied protocols from earlier versions of MACRO
'   NCJ 3 Apr 03 - Added GetVisitCycles (Bug 870)
'   NCJ 8 Apr 03 - Fixed crash in CopyProformaTrial - must set msProtocolName (Bug 1526)
'   NCJ 22 Jun 04 - Bug 2301 - Do not keep a reference to moTopLevelPlanTask (it gets splatted after RemoveComponent)
'   NCJ 25 Nov 04 - Issue 2471 - Don't do save in CloseProformaTrial if there's no guideline
'   ic 15/06/2005 added clinical coding
'   NCJ 8 May 06 - Standardised all error handlers to allow file to be used in Arezzo Rebuild DLL
'------------------------------------------------------------------------------------'
Option Explicit
Option Base 0
Option Compare Binary

' These are for graphical placing of Arezzo tasks - NCJ 10/8/99
Private Const mnPlanW = 1500     ' Width of Arezzo plan task
Private Const mnPlanH = 750     ' Height of Arezzo plan task
Private Const mnTaskGap = 400    ' Gap between tasks
Private Const mnActionW = 1200     ' Width of Arezzo action task
Private Const mnActionH = 1200     ' Height of Arezzo action task

'Changed by Mo Morris 9/8/99 - Declarations for use of CLM.DLL and PSS.DLL
' Use CLM2 - NCJ 27/9/99
' Use CLM3 - NCJ 12/9/00
' Use ALM4 - NCJ 17/5/01 (Changed from gclmCLM to goALM)
' Use ALM5 - NCJ 29/8/02
Public goALM As ALM5
Public gclmGuideline As Guideline

' NCJ 22 Feb 00 - Switch for CLM saving of Arezzo file
' (to speed up Arezzo creation for GGB imports)
Public gbDoCLMSave As Boolean

Private msTopLevelPlanKey As String ' Arezzo key of top level plan (i.e. Trial ID)
'Private moTopLevelPlanTask As Task    ' Arezzo top level plan task

Private msProtocolName As String    ' NCJ 6 Mar 03

'------------------------------------------------------------------
Public Sub InsertProformaVisit(ByVal lClinicalTrialId As Long, _
                               ByVal lVisitId As Long, _
                               ByVal lVisitOrder As Long)
'------------------------------------------------------------------
' Rewritten for Arezzo - NCJ 10/8/99
' Add the visit as a component of the top level plan
' Assume visit plan already created and lVisitID is its Arezzo key
' NCJ 14/1/00 - Improved calculation for LTWH coordinates within plan
' NCJ 22 Jun 04 - Use local oTopPlan instead of moTopLevelPlanTask
'------------------------------------------------------------------
Dim sKey As String
Dim oVisitTask As Task
Dim lTaskLeft As Long
Dim lTaskTop As Long
Dim lVisitColumn As Long
Dim lVisitRow As Long
Dim oTopPlan As Task

    On Error GoTo ErrHandler

    sKey = CStr(lVisitId)   ' Remember that Arezzo keys are strings
    Set oVisitTask = gclmGuideline.colTasks.Item(sKey)
    ' Set the caption to have special prefix NCJ 7 Sept 99
    oVisitTask.Caption = gsVisitPlanPrefix & oVisitTask.Name
    
    Set oTopPlan = gclmGuideline.colTasks.Item(msTopLevelPlanKey)
    oTopPlan.AddComponent sKey
    ' Set its coordinates within the plan
    ' NCJ 14/1/00 - Revised coordinate calculation
    ' Place in rows of 5 each
    lVisitColumn = lVisitOrder Mod 5
    lVisitRow = lVisitOrder \ 5
    lTaskLeft = mnTaskGap + (lVisitColumn * CLng((mnPlanW + mnTaskGap)))
    lTaskTop = mnTaskGap + (lVisitRow * CLng((mnPlanH + mnTaskGap)))
'    lTaskLeft = mnTaskGap + (lVisitOrder - 1) * CLng((mnPlanW + mnTaskGap))
    oTopPlan.SetLTWH sKey, lTaskLeft, lTaskTop, mnPlanW, mnPlanH
    
    SaveCLMGuideline
    'That's all
    
    Set oTopPlan = Nothing
    Set oVisitTask = Nothing
 
Exit Sub
ErrHandler:
      Err.Raise Err.Number, , Err.Description & "|basProformaEditor.InsertProformaVisit"
   
End Sub

'------------------------------------------------------------------
Public Sub InsertProformaStudyVisitCRFPage(ByVal lClinicalTrialId As Long, _
                                            ByVal lVisitId As Long, _
                                            ByVal lCRFPageId As Long, _
                                            ByVal nCRFPageOrder As Integer)
'------------------------------------------------------------------
' Rewritten for Arezzo - NCJ 10/8/99
' We add the existing CRFPage as a component of the Visit plan
' NCJ 14/1/00 - Improved calculation for LTWH coordinates within plan
'------------------------------------------------------------------
Dim oVisit As Task
'Dim clmPage As Task
Dim sKey As String
Dim lTaskLeft As Long
Dim lTaskTop As Long
Dim lPageColumn As Long
Dim lPageRow As Long

    On Error GoTo ErrHandler

    Set oVisit = gclmGuideline.colTasks.Item(CStr(lVisitId))
    ' Get Arezzo key of CRF Page
    sKey = CStr(lCRFPageId)
    'line commented out Mo Morris 15/2/00
    'Set clmPage = gclmGuideline.colTasks.Item(sKey)
    ' Add it as a component of the Visit plan
    oVisit.AddComponent sKey
    ' Calculate its coordinates (although this doesn't work!!)
    ' NCJ 14/1/00 - Revised coordinate calculation
    ' Place in rows of 5 each
    lPageColumn = nCRFPageOrder Mod 5
    lPageRow = nCRFPageOrder \ 5
'    lTaskLeft = mnTaskGap + (nCRFPageOrder - 1) * CLng((mnPlanW + mnTaskGap))
    lTaskLeft = mnTaskGap + (lPageColumn * CLng((mnPlanW + mnTaskGap)))
    lTaskTop = mnTaskGap + (lPageRow * CLng((mnPlanH + mnTaskGap)))
    oVisit.SetLTWH sKey, lTaskLeft, lTaskTop, mnPlanW, mnPlanH
    
    SaveCLMGuideline
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.InsertProformaStudyVisitCRFPage"
        
End Sub

'------------------------------------------------------------------
Public Sub DeleteProformaStudyVisitCRFPage(ByVal lVisitId As Long, _
                                           ByVal lCRFPageId As Long)
'------------------------------------------------------------------
' Remove a CRF page from a visit
' Remove the Arezzo component from its containing plan
' NCJ 10/8/99
' NB Because the CRFPage task is locked, its task definition won't be deleted
'------------------------------------------------------------------

On Error GoTo ErrHandler

    gclmGuideline.colTasks.RemoveComponent CStr(lVisitId), CStr(lCRFPageId)
    SaveCLMGuideline
 
Exit Sub
ErrHandler:
        Err.Raise Err.Number, , Err.Description & "|basProformaEditor.DeleteProformaStudyVisitCRFPage"
    
End Sub

'------------------------------------------------------------------
Public Sub SetVisitRepeats(ByVal lVisitId As Long, ByVal nNoCycles As Integer)
'------------------------------------------------------------------
' NCJ 10 Sept 2002
' Set the given Visit to be a repeating task
' with the given no. of cycles (assume a valid integer)
' NCJ 22 Jun 04 - Use local oTopPlan instead of moTopLevelPlanTask
'------------------------------------------------------------------
Dim sVisitKey As String
Dim oTopPlan As Task

    On Error GoTo ErrHandler

    sVisitKey = CStr(lVisitId)
    
    ' Visits belong to the top level plan
    Set oTopPlan = gclmGuideline.colTasks.Item(msTopLevelPlanKey)
    oTopPlan.NumberOfCycles(sVisitKey) = CStr(nNoCycles)
    
    ' If cycling, also set as "optional"
    oTopPlan.TaskOptional(sVisitKey) = (nNoCycles <> 1)
    
    Call SaveCLMGuideline
    
    Set oTopPlan = Nothing
 
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.SetVisitRepeats"

End Sub

'------------------------------------------------------------------
Public Function GetVisitRepeats(ByVal lVisitId As Long) As Integer
'------------------------------------------------------------------
' NCJ 3 Apr 2003
' Get the no. of cycles for the given Visit
' May be 1 (for non-repeating), -1 (for infinitely repeating), or another integer
' NCJ 22 Jun 04 - Use local oTopPlan instead of moTopLevelPlanTask
'------------------------------------------------------------------
Dim sCycles As String
Dim oTopPlan As Task

    On Error GoTo ErrHandler

    Set oTopPlan = gclmGuideline.colTasks.Item(msTopLevelPlanKey)
    sCycles = oTopPlan.NumberOfCycles(CStr(lVisitId))
    If IsNumeric(sCycles) Then
        GetVisitRepeats = CInt(sCycles)
    Else
        GetVisitRepeats = 1     ' Default to 1
    End If

    Set oTopPlan = Nothing

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.GetVisitRepeats"

End Function

'------------------------------------------------------------------
Public Sub SetEFormRepeats(ByVal lVisitId As Long, ByVal lCRFPageId As Long, ByVal nNoCycles As Integer)
'------------------------------------------------------------------
' NCJ 10 Sept 2002 - To be used when we implement cycle nos. for eForms!
' Set the given eForm to be repeating within given Visit
' with the given no. of cycles (assume a valid integer)
'------------------------------------------------------------------
Dim sCRFKey As String
Dim oVPlan As Task

    On Error GoTo ErrHandler

    sCRFKey = CStr(lCRFPageId)
    
    ' Get the visit plan
    Set oVPlan = gclmGuideline.colTasks.Item(CStr(lVisitId))
    
    oVPlan.NumberOfCycles(sCRFKey) = CStr(nNoCycles)
    
    ' If cycling, also set as "optional"
    oVPlan.TaskOptional(sCRFKey) = (nNoCycles <> 1)
    
    Call SaveCLMGuideline
    
    Set oVPlan = Nothing
 
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.SetEFormRepeats"

End Sub

'------------------------------------------------------------------
Public Sub SetCyclingTask(ByVal lVisitId As Long, _
                            ByVal lCRFPageId As Long, _
                            ByVal bIsCycling As Boolean)
'------------------------------------------------------------------
' Set CRF page within the given Visit to be a repeating/non-repeating form
' bIsCycling says whether it should cycle or not
' NCJ 27/8/99
'------------------------------------------------------------------
Dim sVisitKey As String
Dim sCRFKey As String
Dim clmVPlan As Task

    On Error GoTo ErrHandler

    sVisitKey = CStr(lVisitId)
    sCRFKey = CStr(lCRFPageId)
    ' Get the visit plan
    Set clmVPlan = gclmGuideline.colTasks.Item(sVisitKey)
    If bIsCycling Then
        ' Also set as "optional" - NCJ 22/9/99
        clmVPlan.NumberOfCycles(sCRFKey) = "-1"
        clmVPlan.TaskOptional(sCRFKey) = True
    Else
        clmVPlan.NumberOfCycles(sCRFKey) = "1"
        clmVPlan.TaskOptional(sCRFKey) = False
    End If
    
    SaveCLMGuideline
    
    Set clmVPlan = Nothing
 
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.SetCyclingTask"
    
End Sub

'------------------------------------------------------------------
Public Sub DeleteProformaVisit(ByVal lVisitId As Long, _
                               ByVal lClinicalTrialId As Long)
'------------------------------------------------------------------
' Delete a visit from the trial
' Rewritten for Arezzo - NCJ 10/8/99
' lVisitID is the Arezzo key of the Visit plan
'------------------------------------------------------------------
Dim sVisitKey As String
Dim oVisitTask As Task

    On Error GoTo ErrHandler

    sVisitKey = CStr(lVisitId)
    Set oVisitTask = gclmGuideline.colTasks.Item(sVisitKey)
    ' We must unlock it before deleting it so its definition goes too
    ' NB its CRF components remain locked so won't be deleted
    oVisitTask.Locked = False
    gclmGuideline.colTasks.RemoveComponent msTopLevelPlanKey, sVisitKey
    
    SaveCLMGuideline
 
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.DeleteProformaVisit"
        
End Sub

'------------------------------------------------------------------
Public Sub CreateProformaTrial(ByVal lClinicalTrialId As Long, _
                               ByVal sClinicalTrialName As String)
'------------------------------------------------------------------
' CLM Version - NCJ 9/8/99
' Create a new guideline to represent this trial
' Lock top level plan - NCJ 8 Sept 99
' NCJ 22 Jun 04 - Use local oTopPlan instead of moTopLevelPlanTask
'------------------------------------------------------------------
Dim sSQL As String
Dim oTopPlan As Task

' Default values for Protocols table
Const sVersion = "1"
Const sAREZZO_VERSION = "5"

    On Error GoTo ErrHandler
    
    ' Start a new CLM guideline
    gclmGuideline.StartNew sClinicalTrialName
    
    ' NCJ 6 Mar 03 - Store name for later access to Protocols table
    msProtocolName = sClinicalTrialName
    
    ' Get the top plan task, save it and its key, and set its caption
    msTopLevelPlanKey = gclmGuideline.TopLevelPlanKey   ' Store the key
    Set oTopPlan = gclmGuideline.colTasks.Item(msTopLevelPlanKey)
    oTopPlan.Caption = sClinicalTrialName
    oTopPlan.Locked = True
    
    'ZA 23/09/2002 - removed PssProtocol and use SQL for data insertion
    ' NCJ 25 Sept 02 - Do not set ArezzoFile here, but do set Description
    sSQL = "INSERT INTO Protocols (FileName, ProtocolName, Version, Description, " & _
            " Created, Modified, ArezzoVersion )" & _
            " VALUES ('" & sClinicalTrialName & "', '" & sClinicalTrialName & "', '" & sVersion & "', '" & sClinicalTrialName & "', " & _
            SQLStandardNow & "," & SQLStandardNow & ", '" & sAREZZO_VERSION & "' )"
    
    MacroADODBConnection.Execute sSQL
    
    ' Save initialised Arezzo file
    SaveCLMGuideline
 
    Set oTopPlan = Nothing
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.CreateProformaTrial"
        
End Sub

'------------------------------------------------------------------
Public Sub LoadProformaTrial(ByVal sClinicalTrialName As String)
'------------------------------------------------------------------
' Load the definition of the given clinical trial into the CLM
' Assume the trial does exist
'------------------------------------------------------------------
Dim sSQL As String
Dim oRS As ADODB.Recordset


    On Error GoTo ErrHandler
    
    Set oRS = New ADODB.Recordset
    
    sSQL = "SELECT ArezzoFile FROM Protocols WHERE FileName = '" & sClinicalTrialName & "'"
    oRS.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic
    
    goALM.ArezzoFile = oRS.Fields("ArezzoFile").Value
    
    ' NCJ 6 Mar 03 - Store name for later access to Protocols table
    msProtocolName = sClinicalTrialName
    
    'Set up the top level plan NCJ 23/8/99
    msTopLevelPlanKey = gclmGuideline.TopLevelPlanKey
    ' NCJ 22 Jun 04 - Do NOT keep top level task reference (Bug 2301)
'    Set moTopLevelPlanTask = gclmGuideline.colTasks.Item(msTopLevelPlanKey)
    
    Set oRS = Nothing
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.LoadProformaTrial"
    
End Sub

'------------------------------------------------------------------
Public Function gsCLMVisitName(ByVal sVisitCode As String) As String
'------------------------------------------------------------------
' Create the CLM plan name for an MTM Visit
' NCJ 9/8/99
' Use just the VisitCode, and add the prefix to the task Caption later - NCJ 7 Sept 99
'------------------------------------------------------------------

    On Error GoTo ErrHandler

    ' gsCLMVisitName = gsVisitPlanPrefix & sVisitCode
    gsCLMVisitName = sVisitCode
 
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.gsCLMVisitName"
    
End Function

'------------------------------------------------------------------
Public Function gsCLMFormName(ByVal sCRFPageCode As String) As String
'------------------------------------------------------------------
' Create the CLM action name for an MTM CRF Page
' NCJ 9/8/99
'------------------------------------------------------------------
    
    gsCLMFormName = gsCRFActionPrefix & sCRFPageCode
 
End Function

'------------------------------------------------------------------
Public Function gsCLMCRFName(ByVal sCRFPageCode As String) As String
'------------------------------------------------------------------
' Create the CLM plan name for an MTM CRF Page
' (See gsCLMFormName for the "action" task name)
' NCJ 9/8/99
' Use just the CRFPageCode, and add the prefix to the task Caption later - NCJ 7 Sept 99
'------------------------------------------------------------------

    gsCLMCRFName = sCRFPageCode
     
End Function

'------------------------------------------------------------------
Public Function gnNewCLMPlan(ByVal sPlanName As String) As Long
'------------------------------------------------------------------
' Create a new CLM "plan" task and return its integer key
' Lock it to stop it being deleted when removed from a Visit plan
' NCJ 9/8/99
' NCJ 13/12/99  Now returns Long
'------------------------------------------------------------------
Dim clmtask As Task

    On Error GoTo ErrHandler
    
    Set clmtask = gclmGuideline.colTasks.AddPlan(sPlanName)
    gnNewCLMPlan = CLng(clmtask.TaskKey)    ' TaskKey is a string
    clmtask.Locked = True       ' Lock to prevent deletion
 
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.gnNewCLMPlan"
    
End Function

'------------------------------------------------------------------
Public Function gnNewCLMDataItem(ByVal sDItemCode As String) As Long
'------------------------------------------------------------------
' Create a new Arezzo data item with the given name
' Default to type "text"
' Return integer key as ID
' NCJ 13/12/99 - Now returns a Long
'------------------------------------------------------------------
Dim oDItem As DataItem

    On Error GoTo ErrHandler
    
    Set oDItem = gclmGuideline.colDataItems.Add(sDItemCode, "text")
    oDItem.Locked = True
    gnNewCLMDataItem = CLng(oDItem.DataItemKey)
    
    SaveCLMGuideline
 
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.gnNewCLMDataItem"
    
End Function

'------------------------------------------------------------------
Public Sub InsertProformaCRFPage(ByVal lClinicalTrialId As Long, _
                                ByVal lCRFPageId As Long, _
                                ByVal sCRFPageCode As String)
'------------------------------------------------------------------
' When inserting a CRF page, we create a single action inside it
' Assume the CRF Page plan already created, with lCRFPageID its task key
'------------------------------------------------------------------
Dim clmCRF As Task
Dim sTaskKey As String
Dim clmAction As Task
Dim sActionName As String

    On Error GoTo ErrHandler

    sTaskKey = CStr(lCRFPageId)
    Set clmCRF = gclmGuideline.colTasks.Item(sTaskKey) ' The CRF plan
    ' Set the caption to have special prefix NCJ 7 Sept 99
    clmCRF.Caption = gsCRFPlanPrefix & clmCRF.Name
    
    ' Create new action task with appropriate name
    sActionName = gsCLMFormName(sCRFPageCode)
    Set clmAction = gclmGuideline.colTasks.AddAction(sActionName)
    clmAction.Caption = "Data Entry Control for " & sCRFPageCode
    clmAction.Locked = True
    ' Set Procedure
    clmAction.Procedure = "Data entry for " & sCRFPageCode
    ' Add it to the CRF plan
    clmCRF.AddComponent clmAction.TaskKey
    ' Set its coordinates
    clmCRF.SetLTWH clmAction.TaskKey, mnTaskGap, mnTaskGap, mnActionW, mnActionH
    ' (We'll later move it into an appropriate visit plan)
    
    SaveCLMGuideline
 
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.InsertProformaCRFPage"
    
End Sub

'------------------------------------------------------------------
Public Sub DeleteProformaCRFPage(ByVal lCRFPageId As Long, _
                                 ByVal lClinicalTrialId As Long)
'------------------------------------------------------------------
' Delete the Arezzo task definition of a CRF page
' lCRFPageID is its Arezzo task key
' Must unlock it first (and its data entry action)
'------------------------------------------------------------------
Dim clmCRFPlan As Task
Dim clmtask As Task
Dim sKey As String
Dim vKey As Variant
    
    On Error GoTo ErrHandler
    
    sKey = CStr(lCRFPageId)
    Set clmCRFPlan = gclmGuideline.colTasks.Item(sKey)
    clmCRFPlan.Locked = False
    ' Unlock the action component (actually only one)
    For Each vKey In clmCRFPlan.Components
        Set clmtask = gclmGuideline.colTasks.Item(CStr(vKey))
        clmtask.Locked = False
    Next
    gclmGuideline.colTasks.Remove sKey
    
    SaveCLMGuideline
    
' And that's all
 
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.DeleteProformaCRFPage"
    
End Sub

'------------------------------------------------------------------
Public Sub DeleteProformaDataItem(ByVal vDataItemId As Long)
'------------------------------------------------------------------
' Delete a data item from Arezzo
' NCJ 12/8/99
'------------------------------------------------------------------

    On Error GoTo ErrHandler

    gclmGuideline.colDataItems.Remove CStr(vDataItemId)
    
    SaveCLMGuideline
 
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.DeleteProformaDataItem"

End Sub

'------------------------------------------------------------------
Public Sub UpdateProformaDataItem(ByVal vDataItemId As Long, _
                                  ByVal vDataItemCode As String, _
                                  ByVal vDataItemName As String, _
                                  ByVal vDataItemType As Integer, _
                                  ByVal vDerivation As String, _
                                  ByVal vUnit As String)
'------------------------------------------------------------------
' Update the properties of an Arezzo data item
' CLM version - NCJ 12/8/99
' Assume data item already created, with vDataItemID as its key
' Assume all fields are valid...
' NB The validation conditions are saved elsewhere
'------------------------------------------------------------------
Dim clmDItem As DataItem
Dim sKey As String

    On Error GoTo ErrHandler
    
    sKey = CStr(vDataItemId)
    Set clmDItem = gclmGuideline.colDataItems.Item(sKey)
    
    clmDItem.Name = vDataItemCode
    clmDItem.Caption = vDataItemName
    clmDItem.DataItemType = MTMToArezzoDataType(vDataItemType)
    'changed by Mo Morris 12/1/00, SR2640
    'rtrim required to remove an SQL Server single space on a blank derivation
    clmDItem.Derivation = RTrim(vDerivation)
    clmDItem.Unit = vUnit
    
    SaveCLMGuideline
 
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.UpdateProformaDataItem"
  
End Sub

'------------------------------------------------------------------
Private Function MTMToArezzoDataType(iMTMType As Integer) As String
'------------------------------------------------------------------
' Convert an MTM data type to an Arezzo data type
' ic 13/06/2005 added clinical coding
'------------------------------------------------------------------
Dim sArezzoType As String

    On Error GoTo ErrHandler

    Select Case iMTMType
        'ic 13/06/2005 clinical coding: added thesaurus datatype
        Case DataType.Text, _
                DataType.Category, _
                DataType.Multimedia, DataType.Thesaurus
            sArezzoType = "text"
            
        Case DataType.IntegerData
            sArezzoType = "integer"
            
        Case DataType.Real
            sArezzoType = "real"
            
        Case DataType.Date
            sArezzoType = "datetime"
            
        Case Else
            sArezzoType = "text"
    
    End Select
    MTMToArezzoDataType = sArezzoType
 
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.MTMToArezzoDataType"
    
End Function

'------------------------------------------------------------------
Public Sub SaveProformaWarningCondition(ByVal vDataItemId As Long, _
                    ByVal vWarningFlag As Integer, _
                    ByVal vCondition As String)
'------------------------------------------------------------------
' Pass a warning condition to the CLM
' Assume Warning condition is valid
' NCJ 12/8/99
'------------------------------------------------------------------
Dim clmDItem As DataItem
Dim sKey As String

    On Error GoTo ErrHandler
    
    sKey = CStr(vDataItemId)
    Set clmDItem = gclmGuideline.colDataItems.Item(sKey)
    clmDItem.WarningCondition(CStr(vWarningFlag)) = vCondition
    
    SaveCLMGuideline
 
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.SaveProformaWarningCondition"

    
End Sub
                    
'------------------------------------------------------------------
Public Sub SaveProformaWarningConditions(ByVal vDataItemId As Long, _
                    colWarningFlags As Collection, _
                    colConditions As Collection)
'------------------------------------------------------------------
' NCJ 2/1/01
' Add a pile of warning conditions to the CLM
' Assume collections of flags and conditions are in sync.,
' i.e. i-th WarningFlag is for i-th Condition
' This does NOT remove existing Warning Conditions - see DeleteProformaWarningConditions
'------------------------------------------------------------------
Dim clmDItem As DataItem
Dim sKey As String

    On Error GoTo ErrHandler
    
    sKey = CStr(vDataItemId)
    Set clmDItem = gclmGuideline.colDataItems.Item(sKey)
    Call clmDItem.AddWarningConditions(colWarningFlags, colConditions)
    
    SaveCLMGuideline
 
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.SaveProformaWarningConditions"

End Sub
                    
'------------------------------------------------------------------
Public Sub DeleteProformaWarningCondition(ByVal vDataItemId As Long, _
                    ByVal vWarningFlag As Integer)
'------------------------------------------------------------------
' Delete the warning condition with this Flag
' NCJ 12/8/99
'------------------------------------------------------------------
Dim clmDItem As DataItem
Dim sKey As String

    On Error GoTo ErrHandler
    
    sKey = CStr(vDataItemId)
    Set clmDItem = gclmGuideline.colDataItems.Item(sKey)
    clmDItem.WarningCondition(CStr(vWarningFlag)) = ""
    
    SaveCLMGuideline
 
Exit Sub
ErrHandler:
      Err.Raise Err.Number, , Err.Description & "|basProformaEditor.DeleteProformaWarningCondition"
    
End Sub

'------------------------------------------------------------------
Public Sub DeleteProformaWarningConditions(ByVal vDataItemId As Long)
'------------------------------------------------------------------
' NCJ 2/1/01 - Delete ALL warning conditions & flags for a data item
'------------------------------------------------------------------
Dim clmDItem As DataItem
Dim sKey As String

    On Error GoTo ErrHandler
    
    sKey = CStr(vDataItemId)
    Set clmDItem = gclmGuideline.colDataItems.Item(sKey)
    Call clmDItem.RemoveWarningConditions
    
    SaveCLMGuideline
 
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.DeleteProformaWarningConditions"

End Sub

'------------------------------------------------------------------
Public Sub DeleteProformaRangeValues(ByVal vDataItemId As Long)
'------------------------------------------------------------------
' NCJ 2/1/01 - Delete ALL category values for a data item
'------------------------------------------------------------------
Dim clmDItem As DataItem
Dim sKey As String

    On Error GoTo ErrHandler
    
    sKey = CStr(vDataItemId)
    Set clmDItem = gclmGuideline.colDataItems.Item(sKey)
    Call clmDItem.RemoveRangeValues
    
    SaveCLMGuideline
 
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.DeleteProformaRangeValues"

End Sub

'------------------------------------------------------------------
Public Sub SaveProformaRangeValue(ByVal vDataItemId As Long, _
                    ByVal vNewRangeValue As String, _
                    ByVal vOldRangeValue As String, _
                    ByVal bIsNew As Boolean)
'------------------------------------------------------------------
' Pass a data item's range value to the CLM
' NCJ 12/8/99
' bIsNew is True if this is a new range value
'------------------------------------------------------------------
Dim clmDItem As DataItem
Dim sKey As String

    On Error GoTo ErrHandler
    
    sKey = CStr(vDataItemId)
    Set clmDItem = gclmGuideline.colDataItems.Item(sKey)
        
    If bIsNew Then
        ' the item does not exist so add it
        clmDItem.AddRangeValue (vNewRangeValue)
        
    Else
        ' update the item so remove then add
        clmDItem.RemoveRangeValue (vOldRangeValue)
        clmDItem.AddRangeValue (vNewRangeValue)
    
    End If
    SaveCLMGuideline
 
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.SaveProformaRangeValue"

End Sub

'------------------------------------------------------------------
Public Sub DeleteProformaRangeValue(ByVal vDataItemId As Long, _
                    ByVal vRangeValue As String)
'------------------------------------------------------------------
' Delete a data item's range value
' NCJ 12/8/99
'------------------------------------------------------------------
Dim clmDItem As DataItem
Dim sKey As String

    On Error GoTo ErrHandler
    
    sKey = CStr(vDataItemId)
    Set clmDItem = gclmGuideline.colDataItems.Item(sKey)
    clmDItem.RemoveRangeValue (vRangeValue)
    
    SaveCLMGuideline
 
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.DeleteProformaRangeValue"

End Sub

'------------------------------------------------------------------
Public Sub SaveProformaRangeValues(ByVal vDataItemId As Long, _
                                    colValues As Collection)
'------------------------------------------------------------------
' NCJ 2/1/01
' Add a collection of category values in one go
' NB This doesn't remove existing values - see DeleteProformaRangeValues
'------------------------------------------------------------------
Dim clmDItem As DataItem
Dim sKey As String

    On Error GoTo ErrHandler
    
    sKey = CStr(vDataItemId)
    Set clmDItem = gclmGuideline.colDataItems.Item(sKey)
    Call clmDItem.AddRangeValues(colValues)
    
    SaveCLMGuideline
 
Exit Sub
ErrHandler:
      Err.Raise Err.Number, , Err.Description & "|basProformaEditor.SaveProformaRangeValues"

End Sub

'------------------------------------------------------------------
Public Sub CloseProformaTrial(ByVal lClinicalTrialId As Long, _
                              ByVal nVersionId As Integer, _
                              ByVal sClinicalTrialName As String)
'------------------------------------------------------------------
' Close the PROforma guideline
' NCJ 25 Nov 04 - Issue 2471 - Don't save if there's no guideline!
'------------------------------------------------------------------

    On Error GoTo ErrHandler

    ' NCJ 25 Nov 04 - Do nothing if no guideline (can happen if errors occur during exit of SD)
    If msTopLevelPlanKey = "" Then Exit Sub
    
   ' Save it to the protocols database
   SaveCLMGuideline

    ' Clear the current guideline
    gclmGuideline.Clear
    msTopLevelPlanKey = ""
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.CloseProformaTrial"
    
End Sub

'---------------------------------------------------------
Public Sub ShutDownCLM()
'---------------------------------------------------------
' MM 9/8/99
' Close down ALM.DLL if not already done
'---------------------------------------------------------
    
    On Error GoTo ErrHandler
  
    If Not gclmGuideline Is Nothing Then
        gclmGuideline.Clear
        Set gclmGuideline = Nothing
    End If
    If Not goALM Is Nothing Then
        goALM.CloseALM
        Set goALM = Nothing
    End If
 
Exit Sub
ErrHandler:
        Err.Raise Err.Number, , Err.Description & "|basProformaEditor.ShutDownCLM"

End Sub

'---------------------------------------------------------
Public Sub StartUpCLM(Optional lStudyId As Long = 0)
'---------------------------------------------------------
' MM 9/8/99
' Start up the CLM
' Changed to use new CLM2 - NCJ 27 Sept 99
' NCJ 26/1/00 - Set CLM temp directory
' Changed to use new CLM3 - NCJ 12 Sept 00
' NCJ 1 Feb 01 - Load the Macro_Arezzo file
' NCJ 17 May 01 - Use new ALM
' NCJ 18 Oct 01 - Use new GetPrologSwitches for memory settings
' REM 08/07/02 - Changed name from Macro_Arezzo.pc to Macro3_Arezzo.pc
' NCJ 28 Nov 02 - Use new clsAREZZOMemory for Prolog settings
' NCJ 17 Jan 03 - Added optional StudyId argument
'               If StudyId = -1, then we use default memory settings (for a new study)
'---------------------------------------------------------
Dim sR As String
Dim sFile As String
Dim oArezzoMemory As clsAREZZOMemory

    ' Remove any previous one
    Call ShutDownCLM
    
    Set goALM = New ALM5
    
    Set oArezzoMemory = New clsAREZZOMemory
    'ASH 25/11/2002 StudyId = 0 gets the maximum settings value from the database
    'NCJ 16 Jan 03 - StudyId = -1 gets default settings
    ' NCJ 28 Jan 03 - Pass in DB connection string
    Call oArezzoMemory.Load(lStudyId, goUser.CurrentDBConString)
    
    Call goALM.StartALM(oArezzoMemory.AREZZOSwitches)
    
    ' NCJ 26/1/00 - Set temp directory for CLM
    goALM.TempDirectory = gsTEMP_PATH
    
    ' NCJ 1/2/01 - Load cycling data add-on
    'REM 08/07/02 - Changed name from Macro_Arezzo.pc to Macro3_Arezzo.pc
    sFile = App.Path & "\Macro3_Arezzo.pc"
    Call goALM.GetPrologResult("ensure_loaded('" & sFile & "'), write('0000'). ", sR)
 
    Set gclmGuideline = goALM.Guideline
 
    Set oArezzoMemory = Nothing

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.StartUpCLM"
    
End Sub

'---------------------------------------------------------
Public Sub SaveCLMGuideline()
'---------------------------------------------------------
' Save the current version of Arezzo guideline to the PSS
' NCJ 11 Aug 99
' NCJ 22 Feb 00 - Only save if gbDoCLMSave is TRUE
'---------------------------------------------------------
Dim sSQL As String
Dim rsProtocols As ADODB.Recordset
'Dim sProtocolName As String

    If gbDoCLMSave Then
'        'ZA 23/09/2002
'        sProtocolName = goALM.Guideline.Name

        ' NCJ 6 Mar 03 - Use stored msProtocolName
        
        '   ASH 15/11/2002 fix for error ORA-01704: string literal too long
        sSQL = "SELECT * FROM Protocols WHERE FileName = '" & msProtocolName & "'"
        Set rsProtocols = New ADODB.Recordset
        rsProtocols.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
        'Update Protocols Arezzo file and Modifed columns now
        rsProtocols.Fields("ArezzoFile").Value = goALM.ArezzoFile
        rsProtocols.Fields("Modified").Value = SQLStandardNow
        ' DPH 21/11/2002 - Update required to recordset
        rsProtocols.Update
        rsProtocols.Close
        Set rsProtocols = Nothing
    End If
    
End Sub

'---------------------------------------------------------------------
Public Function IsValidTrialName(ByVal sTrialName As String, ByRef sMsg As String) As Boolean
'---------------------------------------------------------------------
' NCJ 17 Apr 01
' Check that sTrialName is a valid Arezzo task name
'---------------------------------------------------------------------

    sMsg = ""
    
    On Error GoTo BadName
    
    IsValidTrialName = gclmGuideline.IsUniqueName(sTrialName)
    
    Exit Function
    
BadName:
    ' Arezzo spat it out
    sMsg = "Study names cannot be reserved words of AREZZO."
    IsValidTrialName = False
    
    Exit Function


End Function

'---------------------------------------------------------------------
Public Function UniqueVisitFormDataitemCode(ByVal sCode As String, sMsg As String) As Boolean
'---------------------------------------------------------------------
' Checks that sCode (for a new Visit, Form or data item) is unique.
' Returns True if sCode is not being used for any other visit, form or data item,
' otherwise returns False.
' Uses a single CLM call - NCJ 10 Sept 99
' NCJ 27 Jan 00 - Give message here (to save repetitions all over MACRO SD)
'---------------------------------------------------------------------
Dim bIsUnique As Boolean

    UniqueVisitFormDataitemCode = False
    
    On Error GoTo BadName
    
    bIsUnique = gclmGuideline.IsUniqueName(sCode)
    If bIsUnique Then
        UniqueVisitFormDataitemCode = True
    Else
        sMsg = "A Study, Visit, eForm or Question with this code already exists."
    End If
    
    Exit Function
    
BadName:
    ' Arezzo spat it out
    sMsg = "Codes cannot be reserved words of AREZZO."
    
    Exit Function

End Function

'-------------------------------------------------------------
Public Function ValidateDateFormatString(sFormat As String) As Integer
'-------------------------------------------------------------
' NCJ 3 Feb 00
' Returns 0 if NOT valid
' Otherwise returns type of format (as enumeration eDateFormatType)
'-------------------------------------------------------------

    ValidateDateFormatString = goALM.ValidateDateFormat(sFormat)
    
End Function

'----------------------------------------------------------------
Public Function ReadValidDate(ByVal sDateString As String, _
                                ByVal sFormatString As String, _
                                ByRef sArezzoDate As String) As String
'----------------------------------------------------------------
' NCJ 9 Feb 00 - Read a valid date from sDateString according to the format given
' Returns correctly validated date string
' and sArezzoDate is the "actual" date value as an Arezzo term
' (see also ConvertDateFromArezzo)
' NCJ 25/9/00 - Copied in from frmArezzo
'----------------------------------------------------------------

    ReadValidDate = goALM.ReadFormattedDate(sFormatString, sDateString, sArezzoDate)

End Function

'----------------------------------------------------------------
Public Function ConvertDateFromArezzo(ByVal sArezzoDate As String) As Double
'----------------------------------------------------------------
' Convert from Arezzo date string
' to VB's internal format as double
' Returns 0 if sDate not a valid date
' Assume sArezzoDate is one of
'   "date(Y,M,D)"
'   "time(H,Mn,S)"
'   "date(Y,M,D,H,Mn,S)"
' NCJ 26/9/00 - Copied here from DM
'----------------------------------------------------------------

    ConvertDateFromArezzo = goALM.ArezzoDateToDouble(sArezzoDate)

End Function

'----------------------------------------------------------------
Public Function CheckStudyHoldsWater() As Boolean
'----------------------------------------------------------------
' NCJ 15 Feb 01
' Try and alert the user to "funnies" in their study
' To be called when closing a study
' To being with, just warn about infinitely cycling visits
' If it's OK, or if user chooses to ignore warning, return TRUE
' Return FALSE if something is wrong and the user wants to stick around and change it
'----------------------------------------------------------------
Dim sR As String
Dim sMsg As String
Dim sBadVisit As String

    CheckStudyHoldsWater = True
    
    ' Only do this for SD (not Exchange)
    #If SD = 1 Then
        ' Only check study if it exists
        If frmMenu.ClinicalTrialId <= 0 Then
            Exit Function
        End If
        
        sBadVisit = goALM.GetPrologResult("macro_is_infinite. ", sR)
        If sR = "0000" And sBadVisit > "" Then
            ' Yes, there were infinite cycles
            sMsg = "WARNING This study contains the cycling visit '" _
                    & sBadVisit & "' which contains only cycling eForms."
            sMsg = sMsg & vbCrLf & "This study cannot be used for data entry " _
                    & "because the visit will cycle indefinitely."
            sMsg = sMsg & vbCrLf & vbCrLf & "You should ensure that every cycling visit contains at least one non-cycling eForm."
            sMsg = sMsg & vbCrLf & vbCrLf & "Are you sure you wish to close this study?"
            
            CheckStudyHoldsWater = (DialogWarning(sMsg, , True) = vbOK)
        Else
            ' OK, so give no warning
        End If
        
    #End If
    
End Function

'--------------------------------------------------------------------------------------------------------
Public Sub CopyProformaTrial(ByVal sClinicalTrialName As String, ByVal sNewClinicalTrialName As String)
'--------------------------------------------------------------------------------------------------------
' Makes a copy of an Arezzo file in Protocols table
' NCJ 22 Jun 04 - Use local oTopPlan instead of moTopLevelPlanTask
'--------------------------------------------------------------------------------------------------------
Dim sSQL As String
Dim rsProt As ADODB.Recordset
Dim rsNewProt As ADODB.Recordset
Dim oTopPlan As Task

    On Error GoTo ErrHandler
    
    Set rsProt = New ADODB.Recordset
    
    'Get the protocol
    sSQL = "SELECT * FROM Protocols WHERE FileName = '" & sClinicalTrialName & "'"
    
    rsProt.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic
    
    'Insert the new protocol row
    sSQL = "INSERT INTO Protocols (FileName, ProtocolName) VALUES ( '" & _
            sNewClinicalTrialName & "', '" & sNewClinicalTrialName & "' )"
            
    MacroADODBConnection.Execute sSQL
    
    'Update the newly added Protocol record
    ' (Do it by SELECT and UPDATE so it works in Oracle)
    sSQL = "SELECT * FROM Protocols WHERE FileName = '" & sNewClinicalTrialName & "'"
    Set rsNewProt = New ADODB.Recordset
    rsNewProt.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic
    
    rsNewProt.Fields("Version").Value = rsProt.Fields("Version").Value
    rsNewProt.Fields("Description").Value = rsProt.Fields("Description").Value
    rsNewProt.Fields("ArezzoFile").Value = rsProt.Fields("ArezzoFile").Value
    rsNewProt.Fields("ArezzoVersion").Value = rsProt.Fields("ArezzoVersion").Value
    rsNewProt.Fields("Created").Value = SQLStandardNow
    rsNewProt.Fields("Modified").Value = rsNewProt.Fields("Created").Value
    rsNewProt.Update
    rsNewProt.Close
    Set rsNewProt = Nothing
    
    ' NCJ 10 Oct 02
    ' Also load the guideline and rename the top level task
    ' (otherwise the Guideline.Name property will return incorrect results!)
    goALM.ArezzoFile = rsProt.Fields("ArezzoFile").Value
    Set oTopPlan = gclmGuideline.colTasks.Item(gclmGuideline.TopLevelPlanKey)
    ' Rename to new trial name
    oTopPlan.Name = sNewClinicalTrialName
    ' NCJ 8 Apr 03 - Store name for later access to Protocols table (in SaveCLMGuideline)
    msProtocolName = sNewClinicalTrialName
    Call SaveCLMGuideline
    Call gclmGuideline.Clear
    Set oTopPlan = Nothing
    msProtocolName = ""
    
    Set rsProt = Nothing
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basProformaEditor.CopyProformaTrial"

End Sub

