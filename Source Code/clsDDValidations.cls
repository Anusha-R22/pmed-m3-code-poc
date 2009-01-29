VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDDValidations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       clsDDValidations.cls
'   Author:     Paul Norris 03/08/99
'   Purpose:    Class to wrap up a collection of DataItemValidations
'               -reading from the database, saving and loading
'               It uses the clsDDValidation class to store the data retrieved.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   PN  24/09/99    Renamed class name from clsDataDefinitionValidations to clsDDValidations
'                   because prog id is too long with original name
'   NCJ 28/10/99    Exclude "empty" validations from Validate method
'                   IsTermValid renamed to IsConditionValid
'   Mo Morris   16/11/99    DAO to ADO conversion
'   Mo Morris   13/12/99    Id's from integer to Long
'   NCJ 13 Feb 03 - Added CountReal property
'----------------------------------------------------------------------------------------'

Option Explicit

Private miClinicalTrialId As Long
Private miVersionID As Integer
Private msClinicalTrialName As String
Private miDataItemId As Long
Private moValidations As Collection

'---------------------------------------------------------------------
Public Property Get IsDirty() As Boolean
'---------------------------------------------------------------------
Dim bIsDirty As Boolean
Dim oValidation As clsDDValidation
   
    bIsDirty = False
    For Each oValidation In moValidations
        If oValidation.IsDirty Then
            bIsDirty = True
            Exit For
        End If
    Next oValidation
    IsDirty = bIsDirty
    
End Property

'---------------------------------------------------------------------
Public Property Get ClinicalTrialId() As Long
'---------------------------------------------------------------------

    ClinicalTrialId = miClinicalTrialId

End Property

'---------------------------------------------------------------------
Public Property Let ClinicalTrialId(iClinicalTrialId As Long)
'---------------------------------------------------------------------

    miClinicalTrialId = iClinicalTrialId

End Property

'---------------------------------------------------------------------
Public Property Get DataItemId() As Long
'---------------------------------------------------------------------

    DataItemId = miDataItemId

End Property

'---------------------------------------------------------------------
Public Property Let DataItemId(iDataItemId As Long)
'---------------------------------------------------------------------

    miDataItemId = iDataItemId

End Property

'---------------------------------------------------------------------
Public Property Get ClinicalTrialName() As String
'---------------------------------------------------------------------

    ClinicalTrialName = msClinicalTrialName

End Property

'---------------------------------------------------------------------
Public Property Let ClinicalTrialName(sClinicalTrialName As String)
'---------------------------------------------------------------------

    msClinicalTrialName = sClinicalTrialName

End Property

'---------------------------------------------------------------------
Public Property Get VersionId() As Integer
'---------------------------------------------------------------------

    VersionId = miVersionID

End Property

'---------------------------------------------------------------------
Public Property Let VersionId(iVersionId As Integer)

    miVersionID = iVersionId

End Property

'---------------------------------------------------------------------
Private Sub Class_Initialize()
'---------------------------------------------------------------------

    Set moValidations = New Collection

End Sub

'---------------------------------------------------------------------
Public Property Get IsValid() As Boolean
'---------------------------------------------------------------------

    IsValid = Validate

End Property

'---------------------------------------------------------------------
Public Sub Add(oValidation As clsDDValidation)
'---------------------------------------------------------------------

    moValidations.Add oValidation
    Set oValidation = Nothing

End Sub

'---------------------------------------------------------------------
Public Sub Copy(oCopyFromValidations As clsDDValidations)
'---------------------------------------------------------------------
Dim oCopyFromValidation As clsDDValidation
Dim oCopyToValidation As clsDDValidation
       
    On Error GoTo ErrHandler
    
    ' copy all validation items in the collection
    For Each oCopyFromValidation In oCopyFromValidations
        Set oCopyToValidation = New clsDDValidation
        oCopyToValidation.Validation = oCopyFromValidation.Validation
        oCopyToValidation.ValidationTypeID = oCopyFromValidation.ValidationTypeID
        oCopyToValidation.ValidationID = oCopyFromValidation.ValidationID
        oCopyToValidation.ValidationMessage = oCopyFromValidation.ValidationMessage
        oCopyToValidation.IsNew = True
        moValidations.Add oCopyToValidation
        Set oCopyToValidation = Nothing
    Next oCopyFromValidation
    
            
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "copy", "clsDDValidations")
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
Public Sub Remove(lValueID As Long)
'---------------------------------------------------------------------
Dim oValidation As clsDDValidation
    
    With moValidations
        Set oValidation = GetItemFromValueID(lValueID)
        If Not oValidation Is Nothing Then
            oValidation.IsDeleted = True
        Else
            Err.Raise vbObjectError + 10024, "Invalid value id"
        End If
    End With
End Sub

'---------------------------------------------------------------------
Public Sub Save()
'---------------------------------------------------------------------
Dim oValidation As clsDDValidation
       
    On Error GoTo ErrHandler
    
    
    ' first check if the categories are all valid
    If IsValid Then
    
        ' ensure that the records to be deleted are saved first
        ' to free up the ValueID unique reference field
        For Each oValidation In moValidations
            If oValidation.IsDeleted Then
                oValidation.Save miDataItemId, miClinicalTrialId, miVersionID
                
            End If
        Next oValidation

        For Each oValidation In moValidations
            If Not oValidation.IsDeleted Then
                oValidation.Save miDataItemId, miClinicalTrialId, miVersionID
            End If
        Next oValidation
    End If
            
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "save", "clsDDValidations")
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
Public Function GetItemFromValueID(lValueID As Long) As clsDDValidation
'---------------------------------------------------------------------
Dim oValidation As clsDDValidation
       
    On Error GoTo ErrHandler
    
    
    For Each oValidation In moValidations
        If oValidation.ValidationID = lValueID And oValidation.IsDeleted = False Then
            Set GetItemFromValueID = oValidation
            Exit For
        End If
    Next oValidation
            
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "GetItemFromValueID", "clsDDValidations")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select

End Function

'---------------------------------------------------------------------
Public Function Load(iDataItemId As Long, iClinicalTrialId As Long, _
                    iVersionId As Integer) As Boolean
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsValidations As ADODB.Recordset
Dim oValidation As clsDDValidation
    
    ' read the database for the categories that belong to the dataitem id
    sSQL = "SELECT ValidationID,ValidationTypeID,DataItemValidation,ValidationMessage"
    sSQL = sSQL & " FROM DataItemValidation WHERE  ClinicalTrialId  = " & iClinicalTrialId
    sSQL = sSQL & " AND DataItemId = " & iDataItemId
    sSQL = sSQL & " AND VersionId = " & iVersionId
    sSQL = sSQL & " ORDER BY ValidationID, ValidationTypeID"
    Set rsValidations = New ADODB.Recordset
    rsValidations.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With rsValidations
        Do While Not .EOF
            Set oValidation = New clsDDValidation
            With oValidation
                .ValidationID = rsValidations.Fields(0)
                .ValidationTypeID = rsValidations.Fields(1)
                If rsValidations.Fields(2) <> vbNullString Then
                    .Validation = rsValidations.Fields(2)
                End If
                If rsValidations.Fields(3) <> vbNullString Then
                    .ValidationMessage = rsValidations.Fields(3)
                End If
                
                .IsNew = False
                .IsDirty = False
                    
                moValidations.Add oValidation
            End With
            Set oValidation = Nothing
            .MoveNext
        Loop
        .Close
        Set rsValidations = Nothing
    End With
    
    miDataItemId = iDataItemId
    miClinicalTrialId = iClinicalTrialId
    miVersionID = iVersionId
            
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Load", "clsDDValidations")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select

End Function

'---------------------------------------------------------------------
Public Sub Discard()
'---------------------------------------------------------------------

    Call RemoveAll(moValidations)
    Set moValidations = Nothing

End Sub

'---------------------------------------------------------------------
Public Sub Delete()
'---------------------------------------------------------------------
Dim oValidation As clsDDValidation
       
    On Error GoTo ErrHandler
    
    For Each oValidation In moValidations
        oValidation.IsDeleted = True
        oValidation.Save miDataItemId, miClinicalTrialId, miVersionID
    Next oValidation
            
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Delete", "clsDDValidations")
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
Public Sub MarkAllAsDelete()
'---------------------------------------------------------------------
Dim oValidation As clsDDValidation
       
    On Error GoTo ErrHandler
    
    For Each oValidation In moValidations
        oValidation.IsDeleted = True
    Next oValidation
            
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "MarkAllAsDelete", "clsDDValidations")
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
Private Sub RemoveAll(oValidations As Collection)
'---------------------------------------------------------------------

Dim oValidation As clsDDValidation
       
    On Error GoTo ErrHandler
    
    While oValidations.Count > 0
        Set oValidation = oValidations(1)
        oValidations.Remove 1
        Set oValidation = Nothing
    Wend
            
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "RemoveAll", "clsDDValidations")
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
Private Function Validate() As Boolean
'---------------------------------------------------------------------
' Validate all the validations in the current collection
' Allow "empty" validations through - NCJ 28/10/99
'---------------------------------------------------------------------
Dim oValidation As clsDDValidation
       
    On Error GoTo ErrHandler
    
    Validate = True
    For Each oValidation In moValidations
        ' Pick up non-deleted, non-empty objects
        ' If Not oValidation.IsDeleted Then
        If Not (oValidation.IsDeleted Or oValidation.IsEmpty) Then
            ' if the validation has been edited by the user and
            ' has an invalid term or is not valid
                If oValidation.IsDirty And ((Not oValidation.IsConditionValid) Or (Not oValidation.IsValid)) Then
                Validate = False
                Exit For
            End If
        End If
    Next oValidation
            
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Validate", "clsDDValidations")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select

End Function

'---------------------------------------------------------------------
Public Function GetNextUniqueId(oValidations As Collection, Optional lStartCodeValue As Long = 1) As Long
'---------------------------------------------------------------------
' when setting the ValueCode property on the clsDDValidation the
' code must be unique among all non deleted items
' get the next available unique id
'---------------------------------------------------------------------
Dim lNextUniqueID As Long
       
    On Error GoTo ErrHandler
    
    If oValidations Is Nothing Then
        Set oValidations = moValidations
    End If
    
    lNextUniqueID = lStartCodeValue
    While IsUniqueIDUsed(lNextUniqueID, oValidations)
        lNextUniqueID = lNextUniqueID + 1
    Wend
    
    GetNextUniqueId = lNextUniqueID
            
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "GetNextUniqueID", "clsDDValidations")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select


End Function

'---------------------------------------------------------------------
Private Function IsUniqueIDUsed(lId As Long, oValidations As Collection) As Boolean
'---------------------------------------------------------------------
' determine if the unique id is being used in the collection
'---------------------------------------------------------------------
Dim oValidation As clsDDValidation
       
    On Error GoTo ErrHandler
    

    For Each oValidation In oValidations
        If Not oValidation.IsDeleted Then
            If oValidation.ValidationID = lId Then
                IsUniqueIDUsed = True
                Exit For
                
            End If
        
        End If
        
    Next oValidation
                
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "IsUniqueIDUsed", "clsDDValidations")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select


End Function

'---------------------------------------------------------------------
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
'---------------------------------------------------------------------
' The enumerator to allow For...Each iterations
'---------------------------------------------------------------------

    Set NewEnum = moValidations.[_NewEnum]

End Function

'---------------------------------------------------------------------
Public Function CountValid() As Long
'---------------------------------------------------------------------
Dim oValidation As clsDDValidation
Dim lCount As Long
    
    For Each oValidation In moValidations
        If Not oValidation.IsDeleted Then
            lCount = lCount + 1
            
        End If
        
    Next oValidation
    CountValid = lCount
    
End Function

'---------------------------------------------------------------------
Public Function Count() As Long
'---------------------------------------------------------------------

    Count = moValidations.Count

End Function

'---------------------------------------------------------------------
Public Function CountReal() As Long
'---------------------------------------------------------------------
' NCJ 13 Feb 03
' Count of "real" entries, i.e. not deleted or blank
'---------------------------------------------------------------------
Dim oValidation As clsDDValidation
Dim lCount As Long
    
    lCount = 0
    For Each oValidation In moValidations
        If (Not oValidation.IsDeleted) And (Not oValidation.IsEmpty) Then
            lCount = lCount + 1
        End If
    Next oValidation
    
    CountReal = lCount
    
End Function

