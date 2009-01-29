VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsOC.cls
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Toby Aldridge, March 2002
'   Purpose:    class to maintain reference to exteranl object for Discrepancy
'                  cut/paste functionality
'----------------------------------------------------------------------------------------'
' Revisions:
'----------------------------------------------------------------------------------------'

Option Explicit

'reference to external class
Private WithEvents moOCDiscs As OCDiscrepancies
Attribute moOCDiscs.VB_VarHelpID = -1

'----------------------------------------------------------------------------------------'
Public Property Get OCId() As Long
'----------------------------------------------------------------------------------------'
    If Not moOCDiscs Is Nothing Then
        OCId = moOCDiscs.OCId
    End If
End Property

'----------------------------------------------------------------------------------------'
Public Property Get DiscrepancyText() As String
'----------------------------------------------------------------------------------------'

    If Not moOCDiscs Is Nothing Then
        DiscrepancyText = moOCDiscs.DiscrepancyText
    End If
End Property

'----------------------------------------------------------------------------------------'
Public Property Get HaveInfo() As Boolean
'----------------------------------------------------------------------------------------'
'do we have discrepancy info for raising?
'----------------------------------------------------------------------------------------'

    HaveInfo = moOCDiscs.FormOpen
    
End Property

'----------------------------------------------------------------------------------------'
Private Sub Class_Initialize()
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'
    Set moOCDiscs = New OCDiscrepancies

End Sub

'----------------------------------------------------------------------------------------'
Public Sub ShowForm()
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    moOCDiscs.Display gsADOConnectString, goUser
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_Terminate()
'----------------------------------------------------------------------------------------'

    Set moOCDiscs = Nothing

End Sub

'----------------------------------------------------------------------------------------'
Private Sub moOCDiscs_OpenSubject(sStudy As String, sSite As String, sSubjectLabel As String)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'
Dim lSubjectId As Long
Dim oQD As New IMEDQD20.QueryDef
Dim oqds As New IMEDQDS20.QueryDefs
Dim lStudyId As Long
Dim vData As Variant

    On Error GoTo Errlabel
    
    lStudyId = TrialIdFromName(sStudy)
    oQD.InitSelect "TrialSubject", "PersonId", Array("ClinicalTrialId", "TrialSite"), _
                    Array(lStudyId, sSite)
    oQD.QueryFilters.Add "LocalIdentifier1", "like", "%" & sSubjectLabel & "%"
    oQD.Code = "1"
    oqds.Add oQD
    

    oqds.Init gsADOConnectString, ""
    vData = oqds.GetDataArray("1")
    lSubjectId = vData(0, 0)

    'open subject schedule
    frmMenu.SubjectOpen lStudyId, sSite, lSubjectId
    
    Exit Sub
Errlabel:
    
    DialogError "Subject not found"
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub CheckItem(lOCId As Long)
'----------------------------------------------------------------------------------------'

    Call moOCDiscs.MarkAsRaised(lOCId)
    
End Sub

