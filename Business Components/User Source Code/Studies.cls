VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Studies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2002-2003. All Rights Reserved
'   File:       Studies.cls
'   Author:     Richard Meinesz, September 2002
'   Purpose:    Contains all info about the studies a user has access to
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   NCJ 27 Nov 03 - Added class Terminate
'   REM 17/12/03 - In Load routine only do IN statment if not all studies
'   DPH 24/03/2004 - Clean up collections in class_terminate / close recordsets
'   NCJ 22 Mar 05 - Check for blank study descriptions (Bug 2549)
'----------------------------------------------------------------------------------------'

Option Explicit

Private mcolStudies As Collection

'----------------------------------------------------------------------------------------'
Friend Function Item(Key As Variant) As Study
'----------------------------------------------------------------------------------------'
' emulate collection's item function
'----------------------------------------------------------------------------------------'

    Set Item = mcolStudies("K" & Key)

End Function

'----------------------------------------------------------------------------------------'
Friend Function Studies() As Collection
'----------------------------------------------------------------------------------------'

    Set Studies = mcolStudies

End Function

'----------------------------------------------------------------------------------------'
Friend Sub Load(conMACRO As Connection, Optional sStudies As String = "")
'----------------------------------------------------------------------------------------'
'REM 25/09/02
'Loads all the studies that a user has access to for the Role selected
'REVISIONS:
'REM 17/12/03 - only do IN statment if not all studies
'DPH 24/03/2004 - always close recordset
' NCJ 22 Mar 05 - Bug 2549 - Check for blank study desc (to prevent NULL RTE in Oracle)
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsStudies As ADODB.Recordset
Dim vStudies As Variant
Dim colStudies As Collection
Dim oStudy As Study
Dim i As Integer

    On Error GoTo ErrLabel

    Set mcolStudies = New Collection
    
    If sStudies = "" Then Exit Sub

    sSQL = "SELECT * FROM ClinicalTrial"
    If sStudies <> "AllStudies" Then
        sSQL = sSQL & " WHERE ClinicalTrialId IN (" & sStudies & ")"
    Else
        sSQL = sSQL & " WHERE ClinicalTrialId <> 0"
    End If
    sSQL = sSQL & " ORDER BY CLINICALTRIALNAME"
    
    Set rsStudies = New ADODB.Recordset
    rsStudies.Open sSQL, conMACRO, adOpenKeyset, adLockReadOnly, adCmdText
        
    If rsStudies.RecordCount > 0 Then
        vStudies = rsStudies.GetRows
        
        For i = 0 To UBound(vStudies, 2)
            Set oStudy = New Study
            oStudy.StudyId = vStudies(0, i)
            oStudy.StudyName = vStudies(1, i)
            ' NCJ 22 Mar 05 - Check for NULL
            oStudy.StudyDescription = ConvertFromNull(vStudies(2, i), vbString)
            oStudy.PhaseId = vStudies(3, i)
            oStudy.StatusId = vStudies(4, i)
            oStudy.StudyTypeId = vStudies(7, i)
                
            oStudy.Key = "K" & oStudy.StudyId
            mcolStudies.Add oStudy, "K" & oStudy.StudyId
        Next
    
    End If
   
    ' DPH 24/03/2004 - always close recordset
    rsStudies.Close
    Set rsStudies = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Studies.Load"
End Sub

'----------------------------------------------------------------------------------------'
Friend Function NewEnum() As IUnknown
'----------------------------------------------------------------------------------------'

    Set NewEnum = mcolStudies.[_NewEnum]

End Function

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    Set mcolStudies = CollectionDeSerialise(PropBag.ReadProperty("C"), True)
End Sub

'----------------------------------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------------------------------
' NCJ 27 Nov 03 - Tidy up on termination
' DPH 24/03/2004 - use terminate
'----------------------------------------------------------------------

    Call Terminate
    
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
      PropBag.WriteProperty "C", CollectionSerialise(mcolStudies)
End Sub

'----------------------------------------------------------------------------------------'
Public Function StudyById(lStudyId) As Study
'----------------------------------------------------------------------------------------'
'public method to retrieve study info by id
'----------------------------------------------------------------------------------------'

    Set StudyById = Item(lStudyId)
    
End Function

'----------------------------------------------------------------------------------------'
Public Function StudyByName(sStudyName) As Study
'----------------------------------------------------------------------------------------'
'public method to retrieve study info by id
'----------------------------------------------------------------------------------------'
Dim oStudy As Study

    For Each oStudy In mcolStudies
        If oStudy.StudyName = sStudyName Then
            Set StudyByName = oStudy
            Exit Function
        End If
    Next
    
    'if we get here it hasn't found a matching studyname
    Err.Raise vbObjectError + 1100, , "No such study for this user|Studies.StudyByName"
    
End Function

'----------------------------------------------------------------------
Friend Sub Terminate()
'----------------------------------------------------------------------
' DPH 24/03/2004 - terminate all study objects in collection
' Remove references in collection to study objects / set to nothing
'----------------------------------------------------------------------

Dim oStudy As Study

    If Not (mcolStudies Is Nothing) Then
        For Each oStudy In mcolStudies
            oStudy.Terminate
        Next
        
        Call CollectionRemoveAll(mcolStudies)
    End If
    
    Set mcolStudies = Nothing
    
End Sub
