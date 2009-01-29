VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Site"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2002-2003. All Rights Reserved
'   File:       Site.cls
'   Author:     Richard Meinesz, September 2002
'   Purpose:    Contains info about a specific site a user has permission to access
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   NCJ 27 Nov 03 - Added class Terminate
'   REM 05/01/04 - in AddStudy routine Added 3rd parameter, studyid, to CollectionAddAnyway function as we were getting duplicates
'   DPH 24/03/2004 - Clean up collections in class_terminate
'----------------------------------------------------------------------------------------'

Option Explicit

Public Site As String
Public SiteDescription As String
Public SiteStatus As Integer
Public SiteLocation As Integer

Public CountryName As String
Public Key As String

Private mcolStudies As New Collection

'----------------------------------------------------------------------------------------'
Friend Sub AddStudy(oStudy As Study)
'----------------------------------------------------------------------------------------'
'REM 05/01/04 - Added 3rd parameter, studyid, to CollectionAddAnyway as we were getting duplicates
'----------------------------------------------------------------------------------------'

    Call CollectionAddAnyway(mcolStudies, oStudy, "K" & oStudy.StudyId)
        
End Sub

'----------------------------------------------------------------------------------------'
Public Function SiteStudies() As Collection
'----------------------------------------------------------------------------------------'

    Set SiteStudies = mcolStudies

End Function

'----------------------------------------------------------------------
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'----------------------------------------------------------------------
    Site = PropBag.ReadProperty("SI")
    SiteDescription = PropBag.ReadProperty("DE")
    SiteStatus = PropBag.ReadProperty("ST")
    SiteLocation = PropBag.ReadProperty("SL")
    CountryName = PropBag.ReadProperty("CO")
    Key = PropBag.ReadProperty("KE")
End Sub

'----------------------------------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------------------------------
' NCJ 27 Nov 03 - Tidy up on termination
' DPH 24/03/2004 - Call terminate function
'----------------------------------------------------------------------

    Call Terminate
    
End Sub

'----------------------------------------------------------------------
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------------------------
    PropBag.WriteProperty "SI", Site
    PropBag.WriteProperty "DE", SiteDescription
    PropBag.WriteProperty "ST", SiteStatus
    PropBag.WriteProperty "SL", SiteLocation
    PropBag.WriteProperty "CO", CountryName
    PropBag.WriteProperty "KE", Key
End Sub

'----------------------------------------------------------------------
Friend Sub Terminate()
'----------------------------------------------------------------------
' DPH 24/03/2004 - Remove references in collection to Study objects / set to nothing
'----------------------------------------------------------------------

    Call CollectionRemoveAll(mcolStudies)
    
    Set mcolStudies = Nothing
    
End Sub
