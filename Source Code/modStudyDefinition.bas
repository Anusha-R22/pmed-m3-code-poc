Attribute VB_Name = "modStudyDefinition"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999. All Rights Reserved
'   File:       modStudyDefinition.bas
'   Author      Paul Norris, 20/09/99
'   Purpose:    All common functions for the StudyDefintion project are in this module.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   WillC 10/11/99  Added the Error handler
'   Mo Morris   16/11/99    DAO to ADO conversion
'   NCJ 13 Dec 99   Ids to Long
'   Mo Morris   17/12/99,   AdjustSpaceOnForm added for the purpose of inserting/removing
'               space on a form
'   Mo Morris   5/6/00 SR 3472, changes made to AdjustSpaceOnForm
'   NCJ 30 Nov 01 - Added gsgMouseDownOnFormX
'   NCJ 10 Jan 02 - Make sure ALL question lists are updated when deleting eForm
'   NCJ 6 Sept 06 - Make sure we mark study as changed in these routines (MUSD)
'   NCJ 7 Sept 06 - Use new RefreshIsNeeded function in AttemptDeleteCRFPage
'----------------------------------------------------------------------------------------'

Option Explicit

Public gsgMouseDownOnFormY As Single
Public gsgMouseDownOnFormX As Single

'------------------------------------------------------------------------------------'
Public Function AttemptDeleteCRFPage(sUpdateMode As String, sCRFPageName As String, _
                                     lClinicalTrialId As Long, nVersionId As Integer, _
                                     lCRFPageID As Long) As Boolean
'------------------------------------------------------------------------------------'
' PN 18/08/99 - routine added
' attempt to delete a crfpage
' if succeeds then update all open windows
' NCJ 10 Jan 02 - Update ALL question lists for study
' NCJ 7 Sept 06 - Use generic RefreshIsNeeded function
'------------------------------------------------------------------------------------'
Dim sMsg As String
'Dim oForm As Form
'Dim rsCRFPages As ADODB.Recordset
'Dim lNextCRFPageId As Long

    On Error GoTo ErrLabel
    
    AttemptDeleteCRFPage = False
    
    If sUpdateMode = gsREAD Then
        sMsg = "You cannot delete an eForm from this study definition or library"
        Call DialogInformation(sMsg)
        
    ElseIf frmMenu.TrialStatus > 1 Then
        sMsg = "You cannot delete an eForm once a study has been opened"
        Call DialogInformation(sMsg)
        
    Else
        sMsg = "Are you sure you want to delete eForm : " & sCRFPageName & " ?"
        If DialogQuestion(sMsg) = vbYes Then
            ' do the deletion
            Call DeleteCRFPage(lClinicalTrialId, nVersionId, lCRFPageID)
            
            ' NCJ 7 Sept 06 - We now have a new routine which does all the refreshing
            Call frmMenu.RefreshIsNeeded(True)
            
'            ' repaint datalist list
'            ' NCJ 10 Jan 02 - Make sure ALL question lists are updated
'            Call frmMenu.RefreshQuestionLists(lClinicalTrialId)
'
'            ' repaint study visits list
'            Set oForm = frmMenu.gFindForm(lClinicalTrialId, nVersionId, "frmStudyVisits")
'            If Not oForm Is Nothing Then
'                oForm.RefreshStudyVisits
'            End If
'
'            ' repaint the study design
'            Set oForm = frmMenu.gFindForm(lClinicalTrialId, nVersionId, "frmCRFDesign")
'            If Not oForm Is Nothing Then
'                oForm.RefreshCRFPageList
'            End If
'
'            ' frmCRFPageDefiniton now modal so should never find form
'            Set oForm = frmMenu.gFindForm(lClinicalTrialId, nVersionId, "frmCRFPageDefinition")
'            If Not oForm Is Nothing Then
'                If oForm.CRFPageId = lCRFPageId Then
'                    ' repaint the page definition window
'                    ' by displaying another page
'
'                    ' first get the next page
'                    Set rsCRFPages = gdsCRFPageList(lClinicalTrialId, nVersionId)
'                    lNextCRFPageId = -1
'                    With rsCRFPages
'                        Do While Not .EOF
'                            lNextCRFPageId = .Fields("CRFPageID")
'                            Exit Do
'                        Loop
'                    End With
'                    rsCRFPages.Close
'                    Set rsCRFPages = Nothing
'
'                    If lNextCRFPageId > -1 Then
'                        ' then display it
'                        oForm.CRFPageId = lNextCRFPageId
'
'                    Else
'                        ' no more pages to display
'                        oForm.Hide
'
'                    End If
'
'                End If
'
'            End If
            
            AttemptDeleteCRFPage = True
            ' NCJ 6 Sept 06 - Make sure we notify other users of change
            Call frmMenu.MarkStudyAsChanged
        Else
            AttemptDeleteCRFPage = False
            
        End If
    
    End If
 
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modStudyDefinition.AttemptDeleteCRFPage"
    
End Function

'---------------------------------------------------------------------
Public Sub AdjustSpaceOnForm(sgAmount As Single)
'---------------------------------------------------------------------
'For the purspose of inserting/removing space on a form
'called by frmMenu.mnuFLDInsertSpace and frmMenu.mnuFLDRemoveSpace
'
'Changed Mo Morris 5/6/00 SR3472. Comments, Pictures and Lines should always have CaptionY
'and CaptionX set to 0. The previous version  of this sub did not keep to this rule.
'The new version updates Y and CaptionY in separate SQL statements
'---------------------------------------------------------------------
Dim sSQL As String
Dim lScrollPosition As Long

    On Error GoTo ErrLabel
    
    sSQL = "UPDATE CRFElement" _
        & " SET Y = Y + " & sgAmount _
        & " Where ClinicalTrialId = " & frmCRFDesign.ClinicalTrialId _
        & " AND VersionId = " & frmCRFDesign.VersionId _
        & " AND CRFPageid = " & frmCRFDesign.CRFPageId _
        & " AND (Y > " & gsgMouseDownOnFormY & ")"
    
    MacroADODBConnection.Execute sSQL
    
    sSQL = "UPDATE CRFElement" _
        & " SET CaptionY = CaptionY + " & sgAmount _
        & " Where ClinicalTrialId = " & frmCRFDesign.ClinicalTrialId _
        & " AND VersionId = " & frmCRFDesign.VersionId _
        & " AND CRFPageid = " & frmCRFDesign.CRFPageId _
        & " AND (CaptionY > " & gsgMouseDownOnFormY & ")"
    
    MacroADODBConnection.Execute sSQL
    
    'If Space was being removed check that no form elements have been given
    'negative co-ordintates and set them to 100 twips
    sSQL = "UPDATE CRFElement" _
        & " SET Y = 100" _
        & " Where ClinicalTrialId = " & frmCRFDesign.ClinicalTrialId _
        & " AND VersionId = " & frmCRFDesign.VersionId _
        & " AND CRFPageid = " & frmCRFDesign.CRFPageId _
        & " AND (Y < 0 )"
    
    MacroADODBConnection.Execute sSQL
    
    sSQL = "UPDATE CRFElement" _
        & " SET CaptionY = 100" _
        & " Where ClinicalTrialId = " & frmCRFDesign.ClinicalTrialId _
        & " AND VersionId = " & frmCRFDesign.VersionId _
        & " AND CRFPageid = " & frmCRFDesign.CRFPageId _
        & " AND (CaptionY < 0)"
    
    MacroADODBConnection.Execute sSQL
    
    'SDM 25/02/00 SR2881 Get scroll position
    lScrollPosition = frmCRFDesign.vsbCRFPage.Value
    
    'Rebuild the form based on the changed positions
    frmCRFDesign.RefreshMe
    'BuildCRFPage frmCRFDesign, frmCRFDesign.CRFPageId

    'SDM 25/02/00 SR2881 Reset scroll position
    frmCRFDesign.vsbCRFPage.Value = lScrollPosition

    ' NCJ 7 Sept 06 - Make sure we notify other users of change
    Call frmMenu.MarkStudyAsChanged
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modStudyDefinition.AdjustSpaceOnForm"
    
End Sub

