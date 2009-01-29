VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReasonsForChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       clsReasonsForChange.cls
'   Author:     Paul Norris 18/08/99
'   Purpose:    Class to wrap up a collection of StudyDefinition ReasonsForChange-
'               reading and saving to the database
'               It uses the clsReasonForChange class to store the data retrieved.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   Mo Morris   15/11/99    DAO to ADO conversion
'   Mo Morris   13/12/99    Id's from integer to Long
'   NCJ 7 Oct 02 - Class now supports either RFCs or RFOs (must set type on creation)
'   NCJ 8 Oct 02 - Get next ID from database rather than generating from current objects
'----------------------------------------------------------------------------------------'
Option Explicit

Public Event AddedNewItem()

Private moReasonsForChange As Collection
Private WithEvents moNewReason As clsReasonForChange
Attribute moNewReason.VB_VarHelpID = -1

' NCJ 7 Oct 02 - The type of the reasons in this collection
' (Reasons For Change or Reasons For Overrule)
Private menReasonsType As eReasonType

Private mlStudyId As Long
Private mnVersionId As Integer

'---------------------------------------------------------------------
Public Property Get ReasonsType() As eReasonType
'---------------------------------------------------------------------
' Whether the class contains RFOs or RFCs
'---------------------------------------------------------------------

    ReasonsType = menReasonsType
    
End Property

'---------------------------------------------------------------------
Public Property Let ReasonsType(enRT As eReasonType)
'---------------------------------------------------------------------
' Whether the class contains RFOs or RFCs
'---------------------------------------------------------------------

    menReasonsType = enRT

End Property

'---------------------------------------------------------------------
Public Property Get IsDirty() As Boolean
'---------------------------------------------------------------------
Dim bIsDirty As Boolean
Dim oReason As clsReasonForChange
   
    bIsDirty = False
    For Each oReason In moReasonsForChange
        If oReason.IsDirty Then
            bIsDirty = True
            Exit For
        End If
    Next oReason
    IsDirty = bIsDirty
    
End Property

'---------------------------------------------------------------------
Private Sub Class_Initialize()
'---------------------------------------------------------------------

    Set moReasonsForChange = New Collection
    mlStudyId = 0
    mnVersionId = 1

End Sub

'---------------------------------------------------------------------
Public Property Get IsValid() As Boolean
'---------------------------------------------------------------------
Dim bIsValid As Boolean
Dim oReason As clsReasonForChange
   
    On Error GoTo ErrHandler

    bIsValid = True
    For Each oReason In moReasonsForChange
        If Not oReason.IsValid Then
            bIsValid = False
            Exit For
        End If
    Next oReason
    IsValid = bIsValid
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|clsReasonsForChange.IsValid"

End Property

'---------------------------------------------------------------------
Public Sub AddNew()
'---------------------------------------------------------------------
' Add a new (empty) reason to our collection
'---------------------------------------------------------------------
   
    On Error GoTo ErrHandler

    Set moNewReason = Nothing
    Set moNewReason = New clsReasonForChange
    
    ' keep a reference to this item because it will
    ' raise an event when it is valid
    ' we then know when to add a new row to the grid
    With moNewReason
        .ReasonID = GetNextUniqueId(Nothing)
        .IsNew = True
        .IsDirty = False
        .ReasonType = menReasonsType
    End With
    Call Add(moNewReason)
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|clsReasonsForChange.AddNew"

End Sub

'---------------------------------------------------------------------
Public Sub Add(oReason As clsReasonForChange)
'---------------------------------------------------------------------

    With oReason
        moReasonsForChange.Add oReason
        Set oReason = Nothing
    End With

End Sub

'---------------------------------------------------------------------
Public Sub Remove(lId As Long)
'---------------------------------------------------------------------
Dim oReason As clsReasonForChange
   
    On Error GoTo ErrHandler

    With moReasonsForChange
        Set oReason = GetItemFromID(lId)
        If Not oReason Is Nothing Then
            oReason.IsDeleted = True
        Else
            Err.Raise vbObjectError + 10024, "Invalid value id"
        End If
    End With
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|clsReasonsForChange.Remove"

End Sub

'---------------------------------------------------------------------
Public Sub Save(lClinicalTrialId As Long, nVersionId As Integer)
'---------------------------------------------------------------------
Dim oReason As clsReasonForChange
   
    On Error GoTo ErrHandler

    ' ensure that the records to be deleted are saved first
    ' to free up the ValueID unique reference field
    For Each oReason In moReasonsForChange
        If oReason.IsDeleted Then
            oReason.Save lClinicalTrialId, nVersionId
            
        End If
    Next oReason

    For Each oReason In moReasonsForChange
        If Not oReason.IsDeleted Then
            oReason.Save lClinicalTrialId, nVersionId
        End If
    Next oReason
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|clsReasonsForChange.Save"

End Sub

'---------------------------------------------------------------------
Public Function GetItemFromID(lId As Long) As clsReasonForChange
'---------------------------------------------------------------------
Dim oReason As clsReasonForChange
   
    On Error GoTo ErrHandler

    For Each oReason In moReasonsForChange
        If oReason.ReasonID = lId And oReason.IsDeleted = False Then
            Set GetItemFromID = oReason
            Exit For
        End If
    Next oReason
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|clsReasonsForChange.GetItemFromID"

End Function

'---------------------------------------------------------------------
Public Sub Load(lClinicalTrialId As Long, nVersionId As Integer)
'---------------------------------------------------------------------
' Load all reasons for the given study according to this collection's type (RFO or RFC)
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsReasonsForChange As ADODB.Recordset
Dim oReason As clsReasonForChange
   
    On Error GoTo ErrHandler

    ' NCJ 8 Oct 02 - Store study and version
    mlStudyId = lClinicalTrialId
    mnVersionId = nVersionId
    
    ' read the database for the reasons that belong to the study
    ' NCJ 7 Oct 02 - Added filter ReasonType
    sSQL = "SELECT ReasonForChangeID, ReasonForChange FROM ReasonForChange WHERE "
    sSQL = sSQL & "ClinicalTrialId  = " & lClinicalTrialId & " AND VersionId = " & nVersionId
    sSQL = sSQL & " AND ReasonType  = " & menReasonsType
    
    Set rsReasonsForChange = New ADODB.Recordset
    rsReasonsForChange.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With rsReasonsForChange
        Do While Not .EOF
            Set oReason = New clsReasonForChange
            With oReason
                .ReasonID = rsReasonsForChange.Fields(0)
                .Reason = rsReasonsForChange.Fields(1)
                .IsNew = False
                .IsDirty = False
                .ReasonType = menReasonsType
                moReasonsForChange.Add oReason
            End With
            Set oReason = Nothing
            .MoveNext
        Loop
        .Close
        Set rsReasonsForChange = Nothing
    End With
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|clsReasonsForChange.Load"
    
End Sub

'---------------------------------------------------------------------
Public Sub Discard()
'---------------------------------------------------------------------

    Set moReasonsForChange = Nothing
    mlStudyId = 0

End Sub

'---------------------------------------------------------------------
Private Function GetNextUniqueId(oReasons As Collection, _
                                Optional lStartCodeValue As Long = 1) As Long
'---------------------------------------------------------------------
' code must be unique among all non deleted items
' get the next available unique id
'---------------------------------------------------------------------
Dim lNextUniqueID As Long

    On Error GoTo ErrHandler

    If oReasons Is Nothing Then
        Set oReasons = moReasonsForChange
    End If

    lNextUniqueID = lStartCodeValue
    While IsUniqueIDUsed(lNextUniqueID, oReasons)
        lNextUniqueID = lNextUniqueID + 1
    Wend

    GetNextUniqueId = lNextUniqueID

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|clsReasonsForChange.GetNextUniqueID"

End Function

'---------------------------------------------------------------------
Private Function IsUniqueIDUsed(lId As Long, oReasons As Collection) As Boolean
'---------------------------------------------------------------------
' determine if the unique id is being used in the collection
'---------------------------------------------------------------------
Dim oReason As clsReasonForChange

    On Error GoTo ErrHandler

    For Each oReason In oReasons
        If Not oReason.IsDeleted Then
            If oReason.ReasonID = lId Then
                IsUniqueIDUsed = True
                Exit For

            End If

        End If

    Next oReason

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|clsReasonsForChange.IsUniqueIDUsed"

End Function

'---------------------------------------------------------------------
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
'---------------------------------------------------------------------
' The enumerator to allow For...Each iterations
'---------------------------------------------------------------------

    Set NewEnum = moReasonsForChange.[_NewEnum]

End Function

'---------------------------------------------------------------------
Public Function Count() As Long
'---------------------------------------------------------------------

    Count = moReasonsForChange.Count

End Function

'---------------------------------------------------------------------
Public Function CountValid() As Long
'---------------------------------------------------------------------
' How many "real" (non-deleted) reasons in the collection
'---------------------------------------------------------------------
Dim oReason As clsReasonForChange
Dim lCount As Long
   
    On Error GoTo ErrHandler

    For Each oReason In moReasonsForChange
        If Not oReason.IsDeleted Then
            lCount = lCount + 1
        End If
    Next oReason
    CountValid = lCount
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|clsReasonsForChange.CountValid"

End Function

'---------------------------------------------------------------------
Private Sub moNewReason_Valid(bIsValid As Boolean)
'---------------------------------------------------------------------

    ' the last item in the collection is now valid so add a new one
    Call AddNew
    RaiseEvent AddedNewItem
    
End Sub

''---------------------------------------------------------------------
'Private Function GetNextUniqueId() As Integer
''---------------------------------------------------------------------
'' NCJ 8 Oct 02 - Can't use this because we don't save to DB between creating Reasons!
'' Returns the next available unique id for a new Reason in this study
'' NB RFOs and RFCs share the same sequence of IDS (i.e. an RFC will never have the same ID as an RFO)
''---------------------------------------------------------------------
'Dim rsTmp As ADODB.Recordset
'Dim sSQL As String
'
'    On Error GoTo ErrHandler
'
'    ' Ignore ReasonType when getting next ID
'    sSQL = "SELECT MAX(ReasonForChangeId) as MaxReasonId FROM ReasonForChange "
'    If mlStudyId > 0 Then
'        ' If we have a study, filter on it, otherwise just trawl the whole table
'        sSQL = sSQL & "  WHERE ClinicalTrialId  = " & mlStudyId _
'                    & "  AND VersionId  = " & mnVersionId
'    End If
'    Set rsTmp = New ADODB.Recordset
'    rsTmp.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
'
'    If IsNull(rsTmp!MaxReasonId) Then     'if no reasons in this study
'        GetNextUniqueId = gnFIRST_ID
'    Else
'        GetNextUniqueId = rsTmp!MaxReasonId + gnID_INCREMENT
'    End If
'
'    rsTmp.Close
'    Set rsTmp = Nothing
'
'Exit Function
'ErrHandler:
'    Err.Raise Err.Number, , Err.Description & "|clsReasonsForChange.GetNextUniqueID"
'
'End Function



