VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDDCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       clsDDCategory.cls
'   Author:     Paul Norris 27/07/99
'   Purpose:    Class to wrap up the DataItemCategory.
'               It encapsulates the rules that a category must obey and is
'               populated by the clsDDCategories class.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   12/8/99 NCJ Added code for saving to/from Arezzo data items
'   PN  24/09/99    Renamed class name from clsDataDefinitionCategory to clsDDCategory
'                   because prog id is too long with original name
'   NCJ 27-8/10/99 - Deal with max code/value lengths and "empty" categories
'   Mo Morris   16/11/99    DAO to ADO conversion
'   Mo 13/12/99     Id's from integer to Long
'   NCJ 15 Jan 00   SR2721 Do not allow cat codes to start with number but then have letters
'   NCJ 2/6/00 SR3514 Allow underscore in category codes
'   NCJ 02/01/01 SR3825 (efficiency issues) Changed Save to do INSERT instead of get recordset for new categories
'               Do not save Proforma values here (only now done in clsDDCategories)
'   Mo Morris   30/8/01 Db Audit (All aspects of DefaultCat removed)
'----------------------------------------------------------------------------------------'
Option Explicit

Public Event Valid(bIsValid As Boolean)

' class properties
Private msValueCode As String
Private msValue As String
Private mlValueID As Long
Private miActive As Integer
Private miValueOrder As Integer
' this code is the msValueCode that was read from the database
' it is used to determine if the user edits msValueCode
Private msPersistentValueCode As String

' persistence management properties
Private mbIsDirty As Boolean
Private mbIsDeleted As Boolean
Private mbIsNew As Boolean
Private mbIsValid As Boolean

' NCJ 27 Oct 99
Private Const mnMaxCodeLength = 15
Private Const mnMaxValueLength = 255

'---------------------------------------------------------------------
Public Property Get IsEmpty() As Boolean
'---------------------------------------------------------------------
' Category is empty if both code and value are empty
' This will signify that category should be deleted
' NCJ 27/10/99
'---------------------------------------------------------------------
    
    IsEmpty = (msValue = vbNullString And msValueCode = vbNullString)
    
End Property

'---------------------------------------------------------------------
Public Property Get IsValid() As Boolean
'---------------------------------------------------------------------

    IsValid = mbIsValid

End Property

'---------------------------------------------------------------------
Private Property Let IsValid(bIsValid As Boolean)
'---------------------------------------------------------------------
' Set whether this category is valid
' Raise an event to notify validity
'---------------------------------------------------------------------

    If mbIsValid <> bIsValid Then
        mbIsValid = bIsValid
        RaiseEvent Valid(mbIsValid)
    End If

End Property

'---------------------------------------------------------------------
Public Property Get ValueID() As Long
'---------------------------------------------------------------------

    ValueID = mlValueID

End Property

'---------------------------------------------------------------------
Public Property Let ValueID(lValueID As Long)
'---------------------------------------------------------------------

    If lValueID > 0 Then
        If mlValueID <> lValueID Then
            mlValueID = lValueID
            Call Validate
        End If
    Else
        ' raise the invalid property value error
        Err.Raise 380
        
    End If

End Property

'---------------------------------------------------------------------
Public Property Get Active() As Integer
'---------------------------------------------------------------------

    Active = miActive

End Property

'---------------------------------------------------------------------
Public Property Let Active(iActive As Integer)
'---------------------------------------------------------------------

    If iActive = 0 Or iActive = 1 Then
        If miActive <> iActive Then
            mbIsDirty = True
            miActive = iActive
            Call Validate
        End If
    Else
        ' raise the invalid property value error
        Err.Raise 380
    End If

End Property

'---------------------------------------------------------------------
Public Property Get ValueOrder() As Integer
'---------------------------------------------------------------------

    ValueOrder = miValueOrder

End Property

'---------------------------------------------------------------------
Public Property Let ValueOrder(iValueOrder As Integer)
'---------------------------------------------------------------------

    If miValueOrder <> iValueOrder Then
        mbIsDirty = True
        miValueOrder = iValueOrder
        Call Validate
    End If

End Property

'---------------------------------------------------------------------
Public Property Get ValueCode() As String
'---------------------------------------------------------------------

    ValueCode = msValueCode

End Property

'---------------------------------------------------------------------
Public Property Let ValueCode(sValueCode As String)
'---------------------------------------------------------------------
' Assign a new category value code
' Check for max 15 characters - NCJ 27/10/99
'---------------------------------------------------------------------
Dim sErrDesc As String
Dim sValCode As String

    ' NCJ 15/1/00 - Get rid of spaces
    sValCode = Trim(sValueCode)
    
    ' Do not raise error for empty string - NCJ 27/10/99
    ' If sValCode = vbNullString Then Err.Raise 380
    
    ' raise the invalid property value error
    If Not gblnValidString(sValCode, valOnlySingleQuotes) Then Err.Raise 380
    ' NCJ 2/6/00 SR3514 Allow underscore in category codes
    If Not gblnValidString(sValCode, valAlpha + valNumeric + valUnderscore) Then Err.Raise 380

    ' NCJ 15/1/00, SR2721 - Code which starts with a digit must not contain non-numeric chars
    ' (This is an Arezzo restriction)
    If IsNumeric(Left(sValCode, 1)) Then
        If Not IsNumeric(sValCode) Then
            Err.Raise 380
        End If
    End If
    
    ' catch the string too long error
    'If Len(sValCode) > 255 Then
    If Len(sValCode) > mnMaxCodeLength Then    ' Changed NCJ 27/10/99
        sErrDesc = "Category code may not be more than " & mnMaxCodeLength & " characters"
        Err.Raise vbObjectError + 10025, , sErrDesc
    End If

    If msValueCode <> sValCode Then
        msValueCode = sValCode
        mbIsDirty = True
        Call Validate
    End If

End Property

'---------------------------------------------------------------------
Public Property Let PersistentValueCode(sValueCode As String)
'---------------------------------------------------------------------
' This Property is used to keep track of changes made to the ValueCode string
' It should only be used by the clsDDCategories class when
' loading from the database
'---------------------------------------------------------------------

    ' raise the invalid property value error
    If sValueCode = vbNullString Then Err.Raise 380
    ' catch the string too long error
    If Len(sValueCode) > 15 Then
        Err.Raise vbObjectError + 10025, , "String too long"
    End If
    If Not gblnValidString(sValueCode, valOnlySingleQuotes) Then Err.Raise 380
    ' NCJ 2/6/00 SR3514 Allow underscore in category codes
    If Not gblnValidString(sValueCode, valAlpha + valNumeric + valUnderscore) Then Err.Raise 380
    
    msPersistentValueCode = sValueCode
    
End Property

'---------------------------------------------------------------------
Public Property Get Value() As String
'---------------------------------------------------------------------

    Value = msValue

End Property

'---------------------------------------------------------------------
Public Property Let Value(sValue As String)
'---------------------------------------------------------------------
' Assign a category value
' Changed max length check - NCJ 27 Oct 99
'---------------------------------------------------------------------
Dim sErrDesc As String

    ' Do not raise error for empty string - NCJ 27/10/99
    'If sValue = vbNullString Then Err.Raise 380
    
    ' raise the invalid property value error for invalid characters
    If Not gblnValidString(sValue, valOnlySingleQuotes) Then Err.Raise 380
    ' catch the string too long error
    If Len(sValue) > mnMaxValueLength Then
        sErrDesc = "Category value may not be more than " & mnMaxValueLength & " characters"
        Err.Raise vbObjectError + 10025, , sErrDesc
    End If
    
    If msValue <> sValue Then
        mbIsDirty = True
        msValue = sValue
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
Public Sub Validate()
'---------------------------------------------------------------------
' Validate a category - check its values are OK
' NB If Value and Code are both empty then it's not valid
'    (but IsEmpty will be true) - NCJ 27/10/99
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler

    ' Check internal stuff first
    If mlValueID > 0 And _
       miValueOrder > 0 And _
       (miActive = 0 Or miActive = 1) Then
       
            ' Check for non-empty and correct lengths
            If msValue <> vbNullString And _
             msValueCode <> vbNullString And _
             Len(msValue) <= mnMaxValueLength And _
             Len(msValueCode) <= mnMaxCodeLength Then
                 IsValid = True
            Else
                 IsValid = False
            End If
    Else
        ' Something wrong with ValueID or ValueOrder or Active
        ' (Surely shouldn't happen?)
        IsValid = False

    End If
        
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Validate", "clsDDCategory")
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
Private Sub DeleteCategoryValue(iDataItemId As Long, _
                                iClinicalTrialId As Long, _
                                iVersionId As Integer)
'---------------------------------------------------------------------
' Delete category value from the database and from Proforma NCJ 28/10/99
' NCJ 2/1/01 Do not delete from Proforma here
'---------------------------------------------------------------------
Dim sSQL As String
    
    On Error GoTo ErrHandler

    sSQL = "DELETE FROM ValueData WHERE ClinicalTrialId = " & iClinicalTrialId
    sSQL = sSQL & " AND VersionId = " & iVersionId
    sSQL = sSQL & " AND DataItemId = " & iDataItemId
    sSQL = sSQL & " AND ValueID = " & mlValueID
    MacroADODBConnection.Execute sSQL
    
' NCJ 2/1/01 Do not delete from Proforma here
'    ' Delete from Proforma NCJ 12/8/99
'    ' Must delete SAVED value NCJ 28/10/99
'    If msPersistentValueCode <> vbNullString Then   ' it always should be!
'        DeleteProformaRangeValue iDataItemId, msPersistentValueCode
'    End If
        
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "DeleteCategoryValue", "clsDDCategory")
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
Public Sub Save(lDataItemId As Long, lClinicalTrialId As Long, _
                    iVersionId As Integer)
'---------------------------------------------------------------------
' Save the current category values to the DB for given DataItem
' Extra comments added; also check for "empty" categories - NCJ 28/10/99
' NCJ 2/1/01 Do not save to Proforma here (only now done in clsDDCategories)
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsRecord As ADODB.Recordset
    
    On Error GoTo ErrHandler

If mbIsDirty And mbIsValid Then
        ' changes have been made and it's valid
        ' NB mbIsValid is FALSE if both code and value are empty
        ' Check if it's deleted
        If mbIsDeleted Then
            ' If it's deleted but not new we must remove from database
            ' (if it's new there'll be nothing to remove)
            If Not mbIsNew Then
                ' delete this category
                Call DeleteCategoryValue(lDataItemId, lClinicalTrialId, iVersionId)
            End If
            
        Else    ' It's changed but not deleted
        
            If mbIsNew Then
                ' NCJ 02/01/01 - Changed this into a straight INSERT
                '                   for improved efficiency
                'Mo Morris 30/8/01 Db Audit (DefaultCat removed)
                sSQL = "INSERT INTO ValueData " _
                        & "(ClinicalTrialId, VersionId, DataItemId, ValueId, " _
                        & " ValueCode, ItemValue, Active, ValueOrder ) "
                sSQL = sSQL & " VALUES( " & lClinicalTrialId & ", " & iVersionId & ", " _
                        & lDataItemId & ", " & mlValueID & ", '" _
                        & msValueCode & "', '" & ReplaceQuotes(msValue) & "', " _
                        & miActive & ", " & miValueOrder & " ) "
                        
                MacroADODBConnection.Execute sSQL

            Else    ' It's not new, it's not deleted
                ' Prepare to save this category
                sSQL = "Select * FROM ValueData WHERE ClinicalTrialId = " & lClinicalTrialId
                sSQL = sSQL & " AND VersionId = " & iVersionId
                sSQL = sSQL & " AND DataItemId = " & lDataItemId
                sSQL = sSQL & " AND ValueID = " & mlValueID
                Set rsRecord = New ADODB.Recordset
                rsRecord.open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockPessimistic, adCmdText
            
                ' We now store the changes
                ' using rsRecord which has been set up for editing
                With rsRecord
                    .Fields("ValueCode") = msValueCode
                    ' PN 08/09/99 changed field Value to ItemValue
                    .Fields("ItemValue") = msValue
                    .Fields("ValueId") = mlValueID
                    .Fields("Active") = miActive
                    .Fields("ValueOrder") = miValueOrder
                End With
                
                rsRecord.Update
                rsRecord.Close
                Set rsRecord = Nothing
            
            End If
                       
        End If      ' If deleted
        
        mbIsDirty = False
     
    ' Finally if it's changed to empty and it's not new then delete the category
    ElseIf (mbIsDirty And IsEmpty) Or (mbIsDirty And IsDeleted) Then
    
        If Not mbIsNew Then
            ' delete this category
            Call DeleteCategoryValue(lDataItemId, lClinicalTrialId, iVersionId)
        End If
    
        mbIsDirty = False
     
    End If      ' If dirty and valid
        
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                        "Save", "clsDDCategory")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

    
End Sub

'---------------------------------------------------------------------
Private Sub Class_Initialize()
'---------------------------------------------------------------------

    mbIsDirty = True
    miActive = 1

End Sub

