VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2002-2003. All Rights Reserved
'   File:       Sites.cls
'   Author:     Richard Meinesz, September 2002
'   Purpose:    Contains all info about the sites a user has access to
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   NCJ 27 Nov 03 - Added class Terminate
'   REM 17/12/03 - In Load routine only do IN statment if not all sites
'   DPH 24/03/2004 - Clean up collections in class_terminate / close recordsets
'----------------------------------------------------------------------------------------'

Option Explicit

Private mcolSites As Collection

'----------------------------------------------------------------------------------------'
Friend Function Item(Key As Variant) As Site
'----------------------------------------------------------------------------------------'
' emulate collection's item function
'----------------------------------------------------------------------------------------'

    Set Item = mcolSites("K" & Key)

End Function

'----------------------------------------------------------------------------------------'
Friend Function Sites() As Collection
'----------------------------------------------------------------------------------------'

    Set Sites = mcolSites

End Function


'----------------------------------------------------------------------------------------'
Friend Sub Load(enDBType As eMACRODatabaseType, conMACRO As Connection, Optional sSites As String = "")
'----------------------------------------------------------------------------------------'
'REM 24/09/02
'loads the  sites the user has permission to access for the RoleCode selected
'REVISIONS:
'REM 17/12/03 - only do IN statment if not all sites
' DPH 24/03/2004 - always close recordset
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsSites As ADODB.Recordset
Dim vSites As Variant
Dim i As Integer
Dim oSite As Site

    On Error GoTo ErrLabel
    
    Set mcolSites = New Collection
    
    If sSites = "" Then Exit Sub
    
    'get all the sites
    If enDBType = mdtOracle80 Then
        sSQL = "SELECT SITE.SITE, SITE.SITEDESCRIPTION, SITE.SITESTATUS, SITE.SITELOCATION, MACROCOUNTRY.COUNTRYDESCRIPTION FROM Site, MACROCOUNTRY WHERE  SITECOUNTRY = COUNTRYID(+)"
        'REM 17/12/03 - only do IN statment if not all sites
        If sSites <> "AllSites" Then
            sSQL = sSQL & " AND SITE IN (" & sSites & ")"
        End If
        
        sSQL = sSQL & " ORDER BY SITE"
        
    Else
        sSQL = "SELECT SITE.SITE, SITE.SITEDESCRIPTION, SITE.SITESTATUS, SITE.SITELOCATION, MACROCOUNTRY.COUNTRYDESCRIPTION FROM SITE LEFT OUTER JOIN MACROCOUNTRY ON SITECOUNTRY = COUNTRYID"
        'REM 17/12/03 - only do IN statment if not all sites
        If sSites <> "AllSites" Then
            sSQL = sSQL & " WHERE  SITE IN (" & sSites & ")"
        End If
        
        sSQL = sSQL & " ORDER BY SITE"
    End If
    
    Set rsSites = New ADODB.Recordset
    rsSites.Open sSQL, conMACRO, adOpenKeyset, adLockReadOnly, adCmdText
    
    'check there is at least one record
    If rsSites.RecordCount > 0 Then
        vSites = rsSites.GetRows
        
        
        'load each sites properties
        For i = 0 To UBound(vSites, 2)
            Set oSite = New Site
            oSite.Site = vSites(0, i)
            oSite.SiteDescription = vSites(1, i)
            oSite.SiteStatus = vSites(2, i)
            oSite.SiteLocation = Val(ConvertFromNull(vSites(3, i), vbString))
            oSite.CountryName = ConvertFromNull(vSites(4, i), vbString)
            oSite.Key = "K" & oSite.Site
            'add site to coolection
            mcolSites.Add oSite, "K" & oSite.Site
        Next
    
    End If
    
    ' dph 24/03/2004 - always close recordset
    rsSites.Close
    Set rsSites = Nothing
    
Exit Sub
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "Sites.Load"
End Sub

'----------------------------------------------------------------------
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'----------------------------------------------------------------------
    Set mcolSites = CollectionDeSerialise(PropBag.ReadProperty("C"), True)
End Sub

'----------------------------------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------------------------------
' NCJ 27 Nov 03 - Tidy up on termination
' DPH 24/03/2004 - terminate function
'----------------------------------------------------------------------

    Call Terminate
    
End Sub

'----------------------------------------------------------------------
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------------------------
      PropBag.WriteProperty "C", CollectionSerialise(mcolSites)
End Sub

'----------------------------------------------------------------------
Friend Sub Terminate()
'----------------------------------------------------------------------
' DPH 24/03/2004 - Remove references in collection to site objects / set to nothing
'----------------------------------------------------------------------

    Call CollectionRemoveAll(mcolSites)
    
    Set mcolSites = Nothing
    
End Sub

