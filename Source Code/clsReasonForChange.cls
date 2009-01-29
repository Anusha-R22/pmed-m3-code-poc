VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReasonForChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       clsReasonForChange.cls
'   Author:     Paul Norris 18/08/99
'   Purpose:    Class to wrap up a collection of StudyDefinition ReasonsForChange-reading from the database,
'               saving and loading from a text file or odbc source.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   Mo Morris   15/11/99    DAO to ADO conversion
'   Mo Morris   13/12/99    Id's from integer to Long
'   NCJ 7 Oct 02 - Resons object now has a type (RFC or RFO)
'----------------------------------------------------------------------------------------'
Option Explicit

Public Event Valid(bIsValid As Boolean)

Private mlReasonID As Long
Private msReason As String

' persistence management properties
Private mbIsDirty As Boolean
Private mbIsDeleted As Boolean
Private mbIsNew As Boolean
Private mbIsValid As Boolean

' NCJ 7 Oct 02 - The type of the reason
' (Reason For Change or Reason For Overrule)
Private menReasonType As eReasonType


'---------------------------------------------------------------------
Public Property Get ReasonType() As eReasonType
'---------------------------------------------------------------------
' Whether this is an RFO or RFC
'---------------------------------------------------------------------

    ReasonType = menReasonType
    
End Property

'---------------------------------------------------------------------
Public Property Let ReasonType(enRT As eReasonType)
'---------------------------------------------------------------------
' Whether this is an RFO or RFC
'---------------------------------------------------------------------

    menReasonType = enRT

End Property


'---------------------------------------------------------------------
Public Property Get IsValid() As Boolean
'---------------------------------------------------------------------

    IsValid = mbIsValid

End Property

'---------------------------------------------------------------------
Private Property Let IsValid(bIsValid As Boolean)
'---------------------------------------------------------------------

    If mbIsValid <> bIsValid Then
        mbIsValid = bIsValid
        RaiseEvent Valid(mbIsValid)
    End If

End Property

'---------------------------------------------------------------------
Private Sub Validate()
'---------------------------------------------------------------------

    If mlReasonID >= 0 And msReason <> vbNullString Then
        IsValid = True
    Else
        IsValid = False
    End If

End Sub

'---------------------------------------------------------------------
Public Property Get ReasonID() As Long
'---------------------------------------------------------------------

    ReasonID = mlReasonID

End Property

'---------------------------------------------------------------------
Public Property Let ReasonID(lId As Long)
'---------------------------------------------------------------------

    If lId >= 0 Then
        If mlReasonID <> lId Then
            mlReasonID = lId
            Call Validate
        End If
    Else
        ' raise the invalid property value error
        ' no negative values
        Err.Raise 380
        
    End If

End Property

'---------------------------------------------------------------------
Public Property Get Reason() As String
'---------------------------------------------------------------------

    Reason = msReason

End Property

'---------------------------------------------------------------------
Public Property Let Reason(sValue As String)
'---------------------------------------------------------------------

    ' Do not raise an error for empty strings - ATN 16/12/99 SR 2351
    ' raise the invalid property value error
    'If sValue = vbNullString Then Err.Raise 380
    ' catch the string too long error
    If Len(sValue) > 255 Then
        Err.Raise vbObjectError + 10025, , "String too long"
    End If
    If Not gblnValidString(sValue, valOnlySingleQuotes) Then Err.Raise 380

    If msReason <> sValue Then
        IsDirty = True
        msReason = sValue
        Call Validate
    End If

End Property

'---------------------------------------------------------------------
Public Property Get IsDirty() As Boolean
'---------------------------------------------------------------------

    IsDirty = mbIsDirty

End Property

'---------------------------------------------------------------------
Public Property Let IsDirty(bIsDirty As Boolean)
'---------------------------------------------------------------------

    If mbIsDirty <> bIsDirty Then
        mbIsDirty = bIsDirty
    End If

End Property

'---------------------------------------------------------------------
Public Property Get IsNew() As Boolean
'---------------------------------------------------------------------

    IsNew = mbIsNew

End Property

'---------------------------------------------------------------------
Public Property Let IsNew(bIsNew As Boolean)
'---------------------------------------------------------------------

    If bIsNew <> mbIsNew Then
        mbIsNew = bIsNew
        mbIsDirty = True
    End If

End Property

'---------------------------------------------------------------------
Public Property Get IsDeleted() As Boolean
'---------------------------------------------------------------------

    IsDeleted = mbIsDeleted

End Property

'---------------------------------------------------------------------
Public Property Let IsDeleted(bIsDeleted As Boolean)
'---------------------------------------------------------------------

    If bIsDeleted <> mbIsDeleted Then
        mbIsDeleted = bIsDeleted
        mbIsDirty = True
    End If

End Property

'---------------------------------------------------------------------
Public Sub Save(lClinicalTrialId As Long, nVersionId As Integer)
'---------------------------------------------------------------------
' Save this reason to the database
' NCJ 8 Oct 02 - Added ReasonType
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsRecord As ADODB.Recordset

    On Error GoTo ErrHandler
    
    If mbIsDirty Then
        ' changes have been made
       ' WillC SR3211 Added the ability to delete a ReasonForChange if the Reason(msReason) was a null string
       ' the reasons for change in audit trails are stored by the string value so there shouldn't be a conflict
       ' with IDs
        ' If mbIsDeleted Then
        If mbIsDeleted Or msReason = vbNullString Then
            If Not mbIsNew Then
                ' delete this ReasonForChange
                sSQL = "DELETE FROM ReasonForChange WHERE ClinicalTrialId = " & lClinicalTrialId
                sSQL = sSQL & " AND VersionId = " & nVersionId
                sSQL = sSQL & " AND ReasonForChangeID = " & mlReasonID
                sSQL = sSQL & " AND ReasonType  = " & menReasonType
                MacroADODBConnection.Execute sSQL
                
            End If
            
        Else
            If mbIsNew Then
                ' insert this ReasonForChange
                sSQL = "Select * FROM ReasonForChange"
                Set rsRecord = New ADODB.Recordset
                rsRecord.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockPessimistic, adCmdText
                With rsRecord
                    .AddNew
                    .Fields("ClinicalTrialId") = lClinicalTrialId
                    .Fields("VersionId") = nVersionId
                End With
            Else
                ' update this ReasonForChange
                sSQL = "Select * FROM ReasonForChange WHERE ClinicalTrialId = " & lClinicalTrialId
                sSQL = sSQL & " AND VersionId = " & nVersionId
                sSQL = sSQL & " AND ReasonForChangeID = " & mlReasonID
                sSQL = sSQL & " AND ReasonType  = " & menReasonType
                Set rsRecord = New ADODB.Recordset
                rsRecord.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockPessimistic, adCmdText
            End If
            
            With rsRecord
                .Fields("ReasonForChangeID") = mlReasonID
                .Fields("ReasonForChange") = msReason
                .Fields("ReasonType") = menReasonType
            End With
            rsRecord.Update
            rsRecord.Close
            Set rsRecord = Nothing
        End If
        mbIsDirty = False
        
    End If
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|clsReasonForChange.Save"

End Sub

'---------------------------------------------------------------------
Private Sub Class_Initialize()
'---------------------------------------------------------------------

    mbIsValid = False
    mbIsDirty = True

End Sub

