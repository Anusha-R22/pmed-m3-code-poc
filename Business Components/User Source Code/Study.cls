VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Study"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2002-2003. All Rights Reserved
'   File:       Study.cls
'   Author:     Richard Meinesz, September 2002
'   Purpose:
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   NCJ 27 Nov 03 - Added class Terminate
'   DPH 24/03/2004 - Clean up collections in class_terminate
'----------------------------------------------------------------------------------------'

Option Explicit

Public StudyId As Long
Public StudyName As String
Public StudyDescription As String
Public PhaseId As Integer
Public StatusId As Integer
Public StudyTypeId As Integer

Public Key As String

Private mcolSites As New Collection

'----------------------------------------------------------------------------------------'
Friend Sub AddSite(oSite As Site)
'----------------------------------------------------------------------------------------'

    Call CollectionAddAnyway(mcolSites, oSite, "K" & oSite.Site)
        
End Sub

'----------------------------------------------------------------------------------------'
Public Function StudySites() As Collection
'----------------------------------------------------------------------------------------'

    Set StudySites = mcolSites

End Function

'----------------------------------------------------------------------------------------'
Public Function ItemSite(Key As Variant) As Site
'----------------------------------------------------------------------------------------'

    Set ItemSite = mcolSites("K" & Key)
    
End Function

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    StudyId = PropBag.ReadProperty("ID")
    StudyName = PropBag.ReadProperty("NA")
    StudyDescription = PropBag.ReadProperty("DE")
    PhaseId = PropBag.ReadProperty("PH")
    StatusId = PropBag.ReadProperty("ST")
    StudyTypeId = PropBag.ReadProperty("TY")
    Key = PropBag.ReadProperty("KE")
End Sub

'----------------------------------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------------------------------
' NCJ 27 Nov 03 - Tidy up on termination
' DPH 24/03/2004 - call terminate function
'----------------------------------------------------------------------

    Call Terminate
    
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "ID", StudyId
    PropBag.WriteProperty "NA", StudyName
    PropBag.WriteProperty "DE", StudyDescription
    PropBag.WriteProperty "PH", PhaseId
    PropBag.WriteProperty "ST", StatusId
    PropBag.WriteProperty "TY", StudyTypeId
    PropBag.WriteProperty "KE", Key
End Sub

'----------------------------------------------------------------------
Friend Sub Terminate()
'----------------------------------------------------------------------
' DPH 24/03/2004 - terminate all site objects in collection
' Remove references in collection to site objects / set to nothing
'----------------------------------------------------------------------

Dim oSite As Site
    
    If Not (mcolSites Is Nothing) Then
        For Each oSite In mcolSites
            oSite.Terminate
        Next
    
        Call CollectionRemoveAll(mcolSites)
    
    End If
    
    Set mcolSites = Nothing

End Sub
