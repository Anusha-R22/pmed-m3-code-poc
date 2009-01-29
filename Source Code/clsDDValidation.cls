VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDDValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       clsDDValidation.cls
'   Author:     Paul Norris 03/08/99
'   Purpose:    Class to wrap up the DataItemValidation.
'               It encapsulates the rules that a Validation must obey and is
'               populated by the clsDDValidations class.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   NCJ 12/8/99 Changed to CLM validation of conditions, also saving/deleting CLM validations
'   NCJ 24/8/99 Moved saving of Proforma validation to more appropriate place!
'   PN  24/09/99    Renamed class name from clsDataDefinitionValidation to clsDDValidation
'                   because prog id is too long with original name
'   NCJ 28 Oct 99   Deal with empty strings; new IsEmpty property
'                   IsTermValid renamed to IsConditionValid
'   NCJ 9/11/99 Save mlValidationTypeID instead of mlValidationID (SR 2081)
'   NCJ 15/11/99 Reversed change of 9/11 because it was the wrong fix!
'   Mo Morris   16/11/99    DAO to ADO conversion
'   Mo Morris   13/12/99     Id's from integer to Long
'   NCJ 24 Jan 03 - Check for max AREZZO expression length
'----------------------------------------------------------------------------------------'
Option Explicit

Public Event Valid(bIsValid As Boolean)

' Class properties
Private mlValidationID As Long      ' Internal ID
Private mlValidationTypeID As Long  ' ID of type of validation (Reject, Warn, Inform etc.)
Private msValidation As String      ' The validation Condition
Private msValidationMessage As String   ' The validation Message

' Persistence management properties
Private mbIsDirty As Boolean
Private mbIsDeleted As Boolean
Private mbIsNew As Boolean
Private mbIsValid As Boolean

'---------------------------------------------------------------------
Private Sub Validate()
'---------------------------------------------------------------------
' Validate the Validation object (sets IsValid)
' NB It's invalid if ValCond and ValMsg are both empty
'---------------------------------------------------------------------

    If msValidation <> vbNullString And _
       msValidationMessage <> vbNullString Then
        IsValid = True
    Else
        IsValid = False
    End If

End Sub

'---------------------------------------------------------------------
Public Property Get IsEmpty() As Boolean
'---------------------------------------------------------------------
' Returns TRUE if Validation and ValidationMessage are both empty strings
' NCJ 28/10/99
'---------------------------------------------------------------------
    
    IsEmpty = (msValidation = vbNullString And msValidationMessage = vbNullString)
    
End Property

'---------------------------------------------------------------------
Public Property Get IsConditionValid() As Boolean
'---------------------------------------------------------------------
' Validate the validation condition as an Arezzo term
' Renamed from IsTermValid - NCJ 28/10/99
' NCJ 24 Jan 03 - Check the max length allowed for AREZZO terms
'---------------------------------------------------------------------
Dim bResult As Boolean
    
    bResult = True
    
    If Not gblnValidString(msValidation, valOnlySingleQuotes) Then
        bResult = False

    ElseIf Len(msValidation) > glMAX_AREZZO_EXPR_LEN Then
        bResult = False
        
    ElseIf Not gclmGuideline.IsValidCondition(msValidation) Then
            bResult = False
    End If
    
    IsConditionValid = bResult

End Property

'---------------------------------------------------------------------
Public Property Get IsValid() As Boolean
'---------------------------------------------------------------------

    IsValid = mbIsValid

End Property

'---------------------------------------------------------------------
Private Property Let IsValid(bIsValid As Boolean)
'---------------------------------------------------------------------
' Set whether this object is valid or not
' and raise a "Valid" event accordingly
'---------------------------------------------------------------------

    If mbIsValid <> bIsValid Then
        mbIsValid = bIsValid
        RaiseEvent Valid(mbIsValid)
    End If

End Property

'---------------------------------------------------------------------
Public Property Get ValidationID() As Long
'---------------------------------------------------------------------

    ValidationID = mlValidationID

End Property

'---------------------------------------------------------------------
Public Property Let ValidationID(lValidationID As Long)
'---------------------------------------------------------------------

    If lValidationID > 0 Then
        If mlValidationID <> lValidationID Then
            mlValidationID = lValidationID
            Call Validate
        End If
    Else
        ' Raise the invalid property value error
        Err.Raise 380
        
    End If
End Property

'---------------------------------------------------------------------
Public Property Get ValidationTypeID() As Long
'---------------------------------------------------------------------

    ValidationTypeID = mlValidationTypeID

End Property

'---------------------------------------------------------------------
Public Property Let ValidationTypeID(lValidationTypeID As Long)
'---------------------------------------------------------------------

    If mlValidationTypeID <> lValidationTypeID Then
        mlValidationTypeID = lValidationTypeID
        mbIsDirty = True
        Call Validate
    End If
End Property

'---------------------------------------------------------------------
Public Property Get Validation() As String
'---------------------------------------------------------------------

    Validation = msValidation

End Property

'---------------------------------------------------------------------
Public Property Let Validation(sValidation As String)
'---------------------------------------------------------------------
' Assign a validation condition
'---------------------------------------------------------------------

    ' Do not raise an error for empty strings - NCJ 28/10/99
    ' raise the invalid property value error
    ' If sValidation = vbNullString Then Err.Raise 380
    
'Mo Morris 16/12/99, Len validation removed. DataItemValidation now a memo field
'    ' catch the string too long error
'    If Len(sValidation) > 255 Then
'        Err.Raise vbObjectError + 10025, , "Validation condition too long"
'    End If
    If Not gblnValidString(sValidation, valOnlySingleQuotes) Then Err.Raise 380

    If msValidation <> sValidation Then
        mbIsDirty = True
        msValidation = sValidation
        Call Validate
    End If

End Property

'---------------------------------------------------------------------
Public Property Get ValidationMessage() As String
'---------------------------------------------------------------------

    ValidationMessage = msValidationMessage

End Property

'---------------------------------------------------------------------
Public Property Let ValidationMessage(sValidationMessage As String)
'---------------------------------------------------------------------
' Assign a validation message
'---------------------------------------------------------------------

    ' Do not raise an error for empty strings - NCJ 28/10/99
    ' raise the invalid property value error
    ' If sValidationMessage = vbNullString Then Err.Raise 380
    
'Mo Morris 16/12/99, Len Message removed. ValidationMessage now a memo field
'    ' catch the string too long error
'    If Len(sValidationMessage) > 255 Then
'        Err.Raise vbObjectError + 10025, , "Validation message too long"
'    End If
    If Not gblnValidString(sValidationMessage, valOnlySingleQuotes) Then Err.Raise 380
    
    If Len(sValidationMessage) > glMAX_AREZZO_EXPR_LEN Then Err.Raise 380
    
    If msValidationMessage <> sValidationMessage Then
        mbIsDirty = True
        msValidationMessage = sValidationMessage
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

    mbIsDirty = bIsDirty

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
Private Sub DeleteValidation(iDataItemId As Long, _
                             iClinicalTrialId As Long, _
                             iVersionId As Integer)
'---------------------------------------------------------------------
' Delete validation from database
' and delete "warning condition" from Proforma
'---------------------------------------------------------------------
Dim sSQL As String
        
    On Error GoTo ErrHandler
    
    sSQL = "DELETE FROM DataItemValidation WHERE ClinicalTrialId = " & iClinicalTrialId
    sSQL = sSQL & " AND VersionId = " & iVersionId
    sSQL = sSQL & " AND DataItemId = " & iDataItemId
    sSQL = sSQL & " AND ValidationID = " & mlValidationID
    MacroADODBConnection.Execute sSQL
    
    ' Delete it from Proforma
    ' NCJ 12/9/99
    DeleteProformaWarningCondition iDataItemId, mlValidationID
        
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "DeleteValidation", "clsDDValidation")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select

End Sub

'---------------------------------------------------------------------
Public Sub Save(iDataItemId As Long, iClinicalTrialId As Long, _
                    iVersionId As Integer)
'---------------------------------------------------------------------
' Save this validation object
' Check for empty and delete - NCJ 28/10/99
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As adodb.Recordset
        
    On Error GoTo ErrHandler
    
    If mbIsDirty And mbIsValid Then
        ' Changes have been made
        ' NB mbIsValid means both strings are non-empty
        If mbIsDeleted Then
            If Not mbIsNew Then
                ' Delete this Validation if it's not new
                ' (if it's new then there's nothing to delete)
                Call DeleteValidation(iDataItemId, iClinicalTrialId, iVersionId)
            End If
            
        Else    ' It's not deleted
        
            If mbIsNew Then
                ' Insert this Validation
                ' and prepare to add current values
                sSQL = "Select * FROM DataItemValidation"
                Set rsTemp = New adodb.Recordset
                rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockPessimistic, adCmdText
                With rsTemp
                    .AddNew
                    .Fields("ClinicalTrialId") = iClinicalTrialId
                    .Fields("VersionId") = iVersionId
                    .Fields("DataItemId") = iDataItemId
                End With

            Else
                ' Prepare to save this Validation
                sSQL = "Select * FROM DataItemValidation WHERE ClinicalTrialId = " & iClinicalTrialId
                sSQL = sSQL & " AND VersionId = " & iVersionId
                sSQL = sSQL & " AND DataItemId = " & iDataItemId
                sSQL = sSQL & " AND ValidationID = " & mlValidationID
                Set rsTemp = New adodb.Recordset
                rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockPessimistic, adCmdText
                'rsTemp.Edit
                
            End If
            
            ' Now store the new values
            ' using rsTemp previously set up
            With rsTemp
                .Fields("ValidationID") = mlValidationID
                .Fields("ValidationTypeID") = mlValidationTypeID
                .Fields("DataItemValidation") = msValidation
                .Fields("ValidationMessage") = msValidationMessage
            End With
            rsTemp.Update
            rsTemp.Close
            Set rsTemp = Nothing
            ' Save it to Proforma
            ' NCJ 12/9/99 (Moved here 24/8/99)
            ' SR 2081 - NCJ 9/11/99 Save mlValidationTypeID instead of mlValidationID
            ' NCJ 15/11/99 No!! Save mlValidationID instead of mlValidationTypeID
            SaveProformaWarningCondition iDataItemId, _
                                        mlValidationID, msValidation
        End If      ' If mbIsDeleted
        mbIsDirty = False
    
    ' Check to see if it's empty - if so delete if not new
    ' (if it's new then there's nothing to delete)
    ElseIf mbIsDirty And IsEmpty Then
        If Not mbIsNew Then
            Call DeleteValidation(iDataItemId, iClinicalTrialId, iVersionId)
        End If
        mbIsDirty = False
    
    End If
        
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "save", "clsDDValidation")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub Class_Initialize()
'---------------------------------------------------------------------
' NCJ 13/1/00 - Initialise ValidationTypeId to 1 ("Warning")
'---------------------------------------------------------------------

    mbIsDirty = True
    mlValidationTypeID = 1

End Sub
