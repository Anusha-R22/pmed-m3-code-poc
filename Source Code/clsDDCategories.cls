VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDDCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998-2006. All Rights Reserved
'   File:       clsDDCategories.cls
'   Author:     Paul Norris 27/07/99
'   Purpose:    Class to wrap up a collection of DataItemCategories -reading from the database,
'               saving and loading from a text file or odbc source.
'               It uses the clsDDCategory class to store the data retrieved.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   PN  24/09/99    Renamed class name from clsDataDefinitionCategories to clsDDCategories
'                   because prog id is too long with original name
'   NCJ 27/10/99    Check for empty categories when validating
'   NCJ 28/10/99    MaxCatLength for max. length of values/codes
'   Mo Morris   15/11/99    DAO to ADO conversion
'   Mo 13/12/99     Id's from integer to Long
'   NCJ 29/12/99    Fixed problems with deleting categories (and then adding a new one)
'   NCJ 6 Jan 00    SRs 2415, 1798 Check uniqueness of values against codes
'   NCJ 15 Jan 00   SR2722 Use lower case text comparison for code uniqueness check
'   NCJ 2 Jan 01    SR3825 (efficiency issues) Save Arezzo category codes in one dollop in Save
'   NCJ 3 Jan 01    Improved efficiency of loading categories from file
'   ZA  02/08/01    Added IsAllEmpty property
'   Mo Morris   30/8/01 Db Audit (All aspects of DefaultCat removed)
'   DPH 05/11/2001  Added IsAllDeletedOrEmpty property
'   NCJ 13 Feb 03 - Added CountReal property
'   NCJ 10 Jan 06 - Bug 2638 - Don't include deleted items in MaxCatLength
'   NCJ 19 Jan 06 - Also check for active items in MaxCatLength
'----------------------------------------------------------------------------------------'

Option Explicit

Public Enum LoadOption
    ELoadAllRecords
    ELoadAllValidRecords
    ELoadNoneIfErrors
End Enum

Public Enum FieldID
    EFieldIDValueCode
    EFieldIDValueID
    EFieldIDValueOrder
End Enum

Public Enum CategoryLoadResult
    ENoError
    EFileDoesNotExist
    EBadODBCSource
    EBadSQL
    ETooFewFields
    ECodesNotUnique
    EValueBlank
    EInvalidValue
    ExclusiveNoPassword
End Enum

Private mlClinicalTrialId As Long
Private miVersionID As Integer
Private msClinicalTrialName As String
Private mlDataItemId As Long
Private moCategories As Collection
' NCJ 6 Jan 00 - Store a code that is not unique
Private msCodeNotUnique As String

'---------------------------------------------------------------------
Public Property Get IsAllEmpty() As Boolean
'---------------------------------------------------------------------
' Return true if each category is empty
' ZA 02/08/01
'---------------------------------------------------------------------
Dim oCategory As clsDDCategory
Dim bIsAllEmpty As Boolean
    
    bIsAllEmpty = True
    For Each oCategory In moCategories
        If Not oCategory.IsEmpty Then
            bIsAllEmpty = False
            Exit For
        End If
    Next oCategory
    
    IsAllEmpty = bIsAllEmpty
    
End Property

'---------------------------------------------------------------------
Public Property Get IsAllDeletedOrEmpty() As Boolean
'---------------------------------------------------------------------
' Return true if each category is Empty/deleted
' DPH 05/11/2001
'---------------------------------------------------------------------
Dim oCategory As clsDDCategory
Dim bIsAllDeletedOrEmpty As Boolean
    
    bIsAllDeletedOrEmpty = True
    For Each oCategory In moCategories
        If Not (oCategory.IsDeleted Or oCategory.IsEmpty) Then
            bIsAllDeletedOrEmpty = False
            Exit For
        End If
    Next 'oCategory
    
    IsAllDeletedOrEmpty = bIsAllDeletedOrEmpty
    
End Property

'---------------------------------------------------------------------
Public Property Get IsDirty() As Boolean
'---------------------------------------------------------------------
Dim bIsDirty As Boolean
Dim oCategory As clsDDCategory
   
    bIsDirty = False
    For Each oCategory In moCategories
        If oCategory.IsDirty Then
            bIsDirty = True
            Exit For
        End If
    Next oCategory
    IsDirty = bIsDirty
    
End Property

'---------------------------------------------------------------------
Public Sub CopyProperties(oCategories As clsDDCategories)
'---------------------------------------------------------------------

    With oCategories
        mlClinicalTrialId = .ClinicalTrialId
        miVersionID = .VersionId
        msClinicalTrialName = .ClinicalTrialName
        mlDataItemId = .DataItemId
    End With
    
End Sub

'---------------------------------------------------------------------
Public Property Get ClinicalTrialId() As Long
'---------------------------------------------------------------------

    ClinicalTrialId = mlClinicalTrialId

End Property

'---------------------------------------------------------------------
Public Property Let ClinicalTrialId(lClinicalTrialId As Long)
'---------------------------------------------------------------------

    mlClinicalTrialId = lClinicalTrialId

End Property

'---------------------------------------------------------------------
Public Property Get DataItemId() As Long
'---------------------------------------------------------------------

    DataItemId = mlDataItemId

End Property

'---------------------------------------------------------------------
Public Property Let DataItemId(lDataItemId As Long)
'---------------------------------------------------------------------
    
    mlDataItemId = lDataItemId

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
Public Property Get CodeNotUnique() As String
'---------------------------------------------------------------------
' Returns first Code value found that is not unique
' (Only available after calling IsValid)
'---------------------------------------------------------------------

    CodeNotUnique = msCodeNotUnique
    
End Property

'---------------------------------------------------------------------
Public Property Get VersionId() As Integer
'---------------------------------------------------------------------
    
    VersionId = miVersionID

End Property

'---------------------------------------------------------------------
Public Property Let VersionId(iVersionId As Integer)
'---------------------------------------------------------------------
    
    miVersionID = iVersionId

End Property

'---------------------------------------------------------------------
Private Sub Class_Initialize()
'---------------------------------------------------------------------

    Set moCategories = New Collection
    msCodeNotUnique = ""
    
End Sub

'---------------------------------------------------------------------
Public Property Get IsValid() As Boolean
'---------------------------------------------------------------------

    IsValid = ValidateCategories

End Property

'---------------------------------------------------------------------
Public Property Get CodesUnique() As Boolean
'---------------------------------------------------------------------

    ' this function will determine if the codes are unique
    CodesUnique = ValidateCategoryCodes(moCategories)

End Property

'---------------------------------------------------------------------
Public Property Get MaxCatLength() As Integer
'---------------------------------------------------------------------
' Return the maximum length (in characters)
' of any category description or code
' NCJ 10 Jan 06 - Bug 2638 - Check that item isn't deleted!
' NCJ 19 Jan 06 - Also check it's active
'---------------------------------------------------------------------
Dim nMaxLength As Integer
Dim oCategory As clsDDCategory

    nMaxLength = 0
    For Each oCategory In moCategories
        With oCategory
            ' Consider the Value (description) and the Code
            If (Not oCategory.IsDeleted) And oCategory.Active Then     ' NCJ 10 Jan 06 - Check not deleted
                If Len(.Value) > nMaxLength Then
                    nMaxLength = Len(.Value)
                End If
                If Len(.ValueCode) > nMaxLength Then
                    nMaxLength = Len(.ValueCode)
                End If
            End If
        End With
    Next
    MaxCatLength = nMaxLength
    
End Property

'---------------------------------------------------------------------
Public Sub Copy(oCopyFromCategories As clsDDCategories)
'---------------------------------------------------------------------

Dim oCopyFromCategory As clsDDCategory
Dim oCopyToCategory As clsDDCategory
    
    On Error GoTo ErrHandler
    
    ' copy all validation items in the collection
    For Each oCopyFromCategory In oCopyFromCategories
        Set oCopyToCategory = New clsDDCategory
        oCopyToCategory.Active = oCopyFromCategory.Active
        oCopyToCategory.Value = oCopyFromCategory.Value
        oCopyToCategory.ValueCode = oCopyFromCategory.ValueCode
        oCopyToCategory.ValueID = oCopyFromCategory.ValueID
        oCopyToCategory.ValueOrder = oCopyFromCategory.ValueOrder
        oCopyToCategory.IsNew = True
        moCategories.Add oCopyToCategory, oCopyToCategory.ValueID & "K"
        Set oCopyToCategory = Nothing
    Next oCopyFromCategory
        
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Copy", "clsDDCategories")
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
Public Sub Add(oCategory As clsDDCategory)
'---------------------------------------------------------------------

    moCategories.Add oCategory, oCategory.ValueID & "K"
    Set oCategory = Nothing

End Sub

'---------------------------------------------------------------------
Public Sub Remove(lValueID As Long)
'---------------------------------------------------------------------
'---------------------------------------------------------------------
Dim oCategory As clsDDCategory
Dim oCatToRemove As clsDDCategory
    
    On Error GoTo ErrHandler
    
    With moCategories
        Set oCatToRemove = GetItemFromValueID(lValueID)
        If Not oCatToRemove Is Nothing Then
            oCatToRemove.IsDeleted = True
            For Each oCategory In moCategories
'                If oCategory.ValueOrder > moCategories(lValueID & "K").ValueOrder Then
                If oCategory.ValueOrder > oCatToRemove.ValueOrder Then
                    oCategory.ValueOrder = oCategory.ValueOrder - 1
                End If
            Next oCategory
''            ' NCJ 29/12/99 - Do a save (to update DB) and actually remove from collection
''            Call oCatToRemove.Save(mlDataItemId, mlClinicalTrialId, miVersionID)
''            Call moCategories.Remove(lValueID & "K")
        Else
            ' NCJ 29/12/99 - Do not raise error here because GetItemFromValueID
            ' will be Nothing if item has been deleted (which is OK)
            ' Err.Raise vbObjectError + 10024, "Invalid value id"
        End If
    End With
    
    Set oCategory = Nothing
        
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Remove", "clsDDCategories")
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
Public Sub Save()
'---------------------------------------------------------------------
' Save the categories
' NCJ 02/01/01 Save all the Proforma categories at once instead of individually
'---------------------------------------------------------------------
Dim oCategory As clsDDCategory
Dim colCatCodes As Collection

    On Error GoTo ErrHandler
    
    Set colCatCodes = New Collection
    
    ' ensure that the records to be deleted are saved first
    ' to free up the ValueID unique reference field
    For Each oCategory In moCategories
        If oCategory.IsEmpty And oCategory.IsDirty Then
            Call Remove(oCategory.ValueID)
        End If
        If oCategory.IsDeleted Then
            oCategory.Save mlDataItemId, mlClinicalTrialId, miVersionID
        End If
    Next oCategory

    For Each oCategory In moCategories
        If Not oCategory.IsDeleted Then
            If Not oCategory.IsEmpty Then
                ' NCJ 2/1/01 Add code to collection to give to Arezzo
                colCatCodes.Add oCategory.ValueCode
            End If
            oCategory.Save mlDataItemId, mlClinicalTrialId, miVersionID
        End If
    Next oCategory
    
    ' NCJ 2/1/01 Replace value codes in Arezzo
    Call DeleteProformaRangeValues(mlDataItemId)
    Call SaveProformaRangeValues(mlDataItemId, colCatCodes)
    
    Set colCatCodes = Nothing
        
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Save", "clsDDCategories")
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
Public Function GetItemFromValueID(lValueID As Long) As clsDDCategory
'---------------------------------------------------------------------
Dim oCategory As clsDDCategory
    
    On Error GoTo ErrHandler
    
    For Each oCategory In moCategories
        If oCategory.ValueID = lValueID And oCategory.IsDeleted = False Then
            Set GetItemFromValueID = oCategory
            Exit For
        End If
    Next oCategory
        
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "GetItemFromValueID", "clsDDCategories")
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
                    iVersionId As Integer) As CategoryLoadResult
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsCategories As ADODB.Recordset
Dim oCategory As clsDDCategory
    
    On Error GoTo ErrHandler
    
    ' read the database for the categories that belong to the dataitem id
    ' PN 08/09/99 changed field Value to ItemValue
    'Mo Morris 30/8/01 Db Audit (DefaultCat removed)
    sSQL = "SELECT ValueCode,ItemValue,Active,ValueOrder,ValueID FROM ValueData WHERE  ClinicalTrialId  = " & iClinicalTrialId
    sSQL = sSQL & " AND VersionId = " & iVersionId
    sSQL = sSQL & " AND DataItemId = " & iDataItemId
    sSQL = sSQL & " ORDER BY ValueOrder"
    Set rsCategories = New ADODB.Recordset
    rsCategories.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    Load = ENoError
    
    With rsCategories
        Do While Not .EOF
            Set oCategory = New clsDDCategory
            With oCategory
                .ValueCode = rsCategories.Fields(0)
                .PersistentValueCode = .ValueCode
                .Value = rsCategories.Fields(1)
                .Active = rsCategories.Fields(2)
                .ValueID = rsCategories.Fields(4)
                .ValueOrder = rsCategories.Fields(3)
                .IsNew = False
                .IsDirty = False
                    
                moCategories.Add oCategory, .ValueID & "K"
            End With
            Set oCategory = Nothing
            .MoveNext
        Loop
        .Close
        Set rsCategories = Nothing
    End With
    
    mlDataItemId = iDataItemId
    mlClinicalTrialId = iClinicalTrialId
    miVersionID = iVersionId
        
Exit Function

ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Load", "clsDDCategories")
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
Private Function LoadRecord(oTempCategories As Collection, sRecord As String, eErrorLoadOptions As LoadOption) As CategoryLoadResult
'---------------------------------------------------------------------
' NCJ 3/1/01 Changed ConvertStringToCollection to Split
'---------------------------------------------------------------------
Dim oLine As New Collection
Dim vStrings As Variant
Dim nPos As Integer

    On Error GoTo ErrHandler
    
    If sRecord <> vbNullString Then
        Set oLine = Nothing
        
        'check if it is a csv file
        nPos = InStr(1, sRecord, ",")
                    
        'za 13/05/02 - add functionality for loading tab delimited file
        'parse the text based on a comma or a tab
        If nPos = 0 Then
            'no commas in record - this must be tab delimited
            vStrings = Split(sRecord, vbTab, 2)
        Else
        
            vStrings = Split(sRecord, ",", 2)
        End If
        
'        Call ConvertStringToCollection(sRecord, oLine, ",")

        ' We expect array elements 0 and 1
        If UBound(vStrings) = 1 Then
            If vStrings(1) = vbNullString Then
                If eErrorLoadOptions = ELoadAllRecords Then
                    ' label is null but load the record anyway
                    LoadRecord = AddCategoryToCollection(CStr(vStrings(0)), CStr(vStrings(1)), oTempCategories, eErrorLoadOptions)
                Else
                    LoadRecord = EValueBlank
                End If
            Else
                ' the text contains at least two fields
                LoadRecord = AddCategoryToCollection(CStr(vStrings(0)), CStr(vStrings(1)), oTempCategories, eErrorLoadOptions)

            End If
            
        Else
            ' the number of fields is less than 2
            If UBound(vStrings) = 0 Then
                If eErrorLoadOptions = ELoadAllRecords Then
                    ' label is null but load the record anyway
                    LoadRecord = AddCategoryToCollection(CStr(vStrings(0)), vbNullString, oTempCategories, eErrorLoadOptions)
                
                Else
                    LoadRecord = EValueBlank
                
                End If
            
            Else
                ' there are not 2 fields in the file
                LoadRecord = ETooFewFields
            
            End If
            
        End If

' NCJ - Previous PN code using results of ConvertStringToCollection
'        If oLine.Count > 1 Then
'            If oLine(2) = vbNullString Then
'                If eErrorLoadOptions = ELoadAllRecords Then
'                    ' label is null but load the record anyway
'                    LoadRecord = AddCategoryToCollection(oLine(1), oLine(2), oTempCategories, eErrorLoadOptions)
'                Else
'                    LoadRecord = EValueBlank
'                End If
'            Else
'                ' the text contains at least two fields
'                LoadRecord = AddCategoryToCollection(oLine(1), oLine(2), oTempCategories, eErrorLoadOptions)
'
'            End If
'
'        Else
'            ' the number of fields is less than 2
'            If oLine.Count = 1 Then
'                If eErrorLoadOptions = ELoadAllRecords Then
'                    ' label is null but load the record anyway
'                    LoadRecord = AddCategoryToCollection(oLine(1), vbNullString, oTempCategories, eErrorLoadOptions)
'
'                Else
'                    LoadRecord = EValueBlank
'
'                End If
'
'            Else
'                ' there are not 2 fields in the file
'                LoadRecord = ETooFewFields
'
'            End If
'
'        End If
        
    End If
        
Exit Function

ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "LoadRecord", "clsDDCategories")
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
Public Function LoadFromFile(sFileName As String, eErrorLoadOptions As LoadOption) As CategoryLoadResult
'---------------------------------------------------------------------
Dim iFileNumber As Integer
Dim sLine As String
Dim oTempCategories As New Collection
Dim eResult As CategoryLoadResult
    
    On Error GoTo ErrHandler
    
    ' read the file to populate the collection
    LoadFromFile = ENoError
    sFileName = TrimNull(sFileName)
    If FileExists(sFileName) Then
        iFileNumber = FreeFile
        Open sFileName For Input As iFileNumber
            Do While Not EOF(iFileNumber)
                Line Input #iFileNumber, sLine
                eResult = LoadRecord(oTempCategories, sLine, eErrorLoadOptions)
                If LoadFromFile = ENoError Then
                    ' only assign the return code
                    ' if it is still set to no error
                    LoadFromFile = eResult
                End If
            Loop
        
        Close iFileNumber
        
    Else
        ' raise the file not found error
        LoadFromFile = EFileDoesNotExist
    End If
    
    Call CompleteLoad(eErrorLoadOptions, oTempCategories, LoadFromFile)
        
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "LoadFromFile", "clsDDCategories")
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
Private Sub CompleteLoad(eErrorLoadOptions As LoadOption, oTempCategories As Collection, eLoadResult As CategoryLoadResult)
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Select Case eErrorLoadOptions
    Case ELoadAllRecords, ELoadAllValidRecords
        ' load the reocrds even if there was an error
        
        ' save the new categories by copying them into the permanent collection
        Call CopyCategories(oTempCategories, moCategories)
        
    Case ELoadNoneIfErrors
        ' only load if no error occurred
         
        If eLoadResult = ENoError Then
            ' save the new categories by copying them into the permanent collection
            Call CopyCategories(oTempCategories, moCategories)
        End If
    
    End Select
    ' properly destroy the temp collection
    Call RemoveCategories(oTempCategories)
        
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CompleteLoad", "clsDDCategories")
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
Public Sub Discard()
'---------------------------------------------------------------------

    Call RemoveCategories(moCategories)
    Set moCategories = Nothing

End Sub

'---------------------------------------------------------------------
Public Sub CopyCategories(oFrom As Collection, oTo As Collection)
'---------------------------------------------------------------------
Dim oCategory As clsDDCategory
    
    For Each oCategory In oFrom
        oTo.Add oCategory
    Next oCategory

End Sub

'---------------------------------------------------------------------
Public Sub Delete()
'---------------------------------------------------------------------
Dim oCategory As clsDDCategory
    
    For Each oCategory In moCategories
        oCategory.IsDeleted = True
        oCategory.Save mlDataItemId, mlClinicalTrialId, miVersionID
    Next oCategory

End Sub

'---------------------------------------------------------------------
Public Sub MarkAllAsDelete()
'---------------------------------------------------------------------
Dim oCategory As clsDDCategory
    
    For Each oCategory In moCategories
        oCategory.IsDeleted = True
    Next oCategory
End Sub

'---------------------------------------------------------------------
Private Sub RemoveCategories(oCategories As Collection)
'---------------------------------------------------------------------
Dim oCat As clsDDCategory
    
    While oCategories.Count > 0
        Set oCat = oCategories(1)
        oCategories.Remove 1
        Set oCat = Nothing
    Wend
    
End Sub

'---------------------------------------------------------------------
Public Function LoadFromODBCSource(sConnect As String, sSQL As String, _
                                    eErrorLoadOptions As LoadOption) As CategoryLoadResult
'---------------------------------------------------------------------
' open the odbc source to load the categories
'---------------------------------------------------------------------
Dim rsExternalDataSource As ADODB.Recordset
Dim oTempCategories As New Collection
Dim eResult As CategoryLoadResult
Dim sFieldValue0 As String
Dim sFieldValue1 As String
Dim cnNewConnection As ADODB.Connection

    ' open the connection using odbc
    On Error GoTo BadConnectString
    sConnect = TrimNull(sConnect)
        
    ' run the SQL
    On Error GoTo BadSQl
    sSQL = TrimNull(sSQL)
    Set rsExternalDataSource = New ADODB.Recordset
    ' Use default database
    If UCase(sConnect) = UCase("MacroADODBConnection") Then
        rsExternalDataSource.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    Else
        ' WillC 5/1/2000
        ' Use non default database ,get connection information from frmADOConnect
        Set cnNewConnection = New ADODB.Connection
        'changed Mo Morris 12/1/00
        'sConnect = frmADOConnect.ConnectString
        'sConnect = frmDataDefinition.txtConnect
        cnNewConnection.ConnectionString = sConnect
        cnNewConnection.Open
        rsExternalDataSource.Open sSQL, cnNewConnection, adOpenKeyset, adLockReadOnly, adCmdText
    End If
        
    If rsExternalDataSource.Fields.Count > 1 Then
        ' at least two fields exist
        With rsExternalDataSource
            LoadFromODBCSource = ENoError
            Do While Not .EOF
                If IsNull(.Fields(0)) Then
                    sFieldValue0 = vbNullString
                Else
                    sFieldValue0 = .Fields(0)
                End If
                If IsNull(.Fields(1)) Then
                    sFieldValue1 = vbNullString
                Else
                    sFieldValue1 = .Fields(1)
                End If
                
                eResult = LoadRecord(oTempCategories, sFieldValue0 & "," & sFieldValue1, eErrorLoadOptions)

                If LoadFromODBCSource = ENoError Then
                    ' only assign the return code
                    ' if it is still set to no error
                    LoadFromODBCSource = eResult
                End If
                .MoveNext
            Loop
        End With

    Else
        ' there are not enough fields in the sql query string
        LoadFromODBCSource = ETooFewFields
    
    End If
    
    Call CompleteLoad(eErrorLoadOptions, oTempCategories, LoadFromODBCSource)

    Set rsExternalDataSource = Nothing
    Exit Function

BadConnectString:
    LoadFromODBCSource = EBadODBCSource
    Set rsExternalDataSource = Nothing
    Exit Function

BadSQl:
    'za 03/05/02 - check if db is password protected
    'database is password protected error provided by MS Access
    If Err.Number = -2147217843 Then
        LoadFromODBCSource = ExclusiveNoPassword
    Else
        LoadFromODBCSource = EBadSQL
    End If
    Set rsExternalDataSource = Nothing

    Set cnNewConnection = Nothing
        
End Function

'---------------------------------------------------------------------
Private Function ValidateCategoryCodes(oCategories As Collection) As Boolean
'---------------------------------------------------------------------
' Validate category codes i.e. check no duplicates
'---------------------------------------------------------------------
Dim oCategory As clsDDCategory
Dim sValueCodes As String
    
    On Error GoTo ErrHandler
    
    ValidateCategoryCodes = True
    
    ' build a string containing codes separated by ";"
    sValueCodes = ";"
    For Each oCategory In oCategories
        ' Check not empty too - NCJ 27/10/99
        If Not (oCategory.IsDeleted Or oCategory.IsEmpty) Then
                sValueCodes = sValueCodes & oCategory.ValueCode & ";"
        End If
    Next oCategory
   ' Debug.Print sValueCodes
    
    ' Check codes unique by searching for the code in the string
    For Each oCategory In oCategories
        If CountStrsInStr(sValueCodes, ";" & oCategory.ValueCode & ";", vbTextCompare) > 1 Then
            ValidateCategoryCodes = False
            Exit For
        End If
    Next oCategory
        
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ValidateCategoryCodes", "clsDDCategories")
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
Private Function CheckValueCodeUnique(sValueCode As String, nValueID As Integer) As Boolean
'---------------------------------------------------------------------
' NCJ 6 Jan 00 - SRs 2415, 1798
' Check that ValueCode is unique
' i.e. that it matches no ValueCode or Value except its own
' nValueID identifies the category
'---------------------------------------------------------------------
Dim oCategory As clsDDCategory
    
    On Error GoTo ErrHandler
    
    For Each oCategory In moCategories
        ' Don't check ourselves
        If oCategory.ValueID <> nValueID Then
            ' Check against the ValueCode and the Value
            If oCategory.ValueCode = sValueCode Or oCategory.Value = sValueCode Then
                CheckValueCodeUnique = False
                Exit Function
            End If
        End If
    Next
        
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CheckValueCodeUnique", "clsDDCategories")
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
Private Function CheckValueUnique(sValue As String, nValueID As Integer) As Boolean
'---------------------------------------------------------------------
' NCJ 6 Jan 00 - SRs 2415, 1798
' Check that Value is unique
' i.e. that it matches no code except its own
' (matching other values doesn't matter)
' nValueID identifies the category
'---------------------------------------------------------------------
Dim oCategory As clsDDCategory
    
    On Error GoTo ErrHandler
    

    For Each oCategory In moCategories
        ' Don't check ourselves
        If oCategory.ValueID <> nValueID Then
            If oCategory.ValueCode = sValue Then
                CheckValueUnique = False
                Exit Function
            End If
        End If
    Next
        
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CheckValueUnique", "clsDDCategories")
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
Private Function CheckValuesUnique() As Boolean
'---------------------------------------------------------------------
' NCJ 6 Jan 00 - SRs 2415, 1798
' Check each code against each other code and each other value
' Codes must be unique and not match any value other than their own
' Values don't have to be unique but they mustn't match any code other than their own
' NCJ 15/1/00, SR2722 - Do lower case text comparison
'---------------------------------------------------------------------
Dim oCategory1 As clsDDCategory
Dim oCategory2 As clsDDCategory
    
    On Error GoTo ErrHandler
    

    msCodeNotUnique = ""
    For Each oCategory1 In moCategories
        ' Check not deleted or empty
        If Not (oCategory1.IsDeleted Or oCategory1.IsEmpty) Then
            For Each oCategory2 In moCategories
                 ' Check not deleted or empty
                If Not (oCategory2.IsDeleted Or oCategory2.IsEmpty) Then
                    ' Don't check ourselves
                    If oCategory1.ValueID <> oCategory2.ValueID Then
                        ' NCJ 15/1/00 - Use LCase conversion
                        If LCase(oCategory1.ValueCode) = LCase(oCategory2.ValueCode) _
                           Or LCase(oCategory1.ValueCode) = LCase(oCategory2.Value) Then
                            CheckValuesUnique = False
                            msCodeNotUnique = oCategory1.ValueCode
                            Exit Function
                        End If
                    End If
                End If
            Next oCategory2
        End If
    Next oCategory1
    
    CheckValuesUnique = True
        
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CheckValuesUnique", "clsDDCategories")
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
Private Function ValidateCategories() As Boolean
'---------------------------------------------------------------------
' Validate all categories
' Check for "empty" categories - NCJ 27/10/99
'---------------------------------------------------------------------
Dim oCategory As clsDDCategory
Dim bIsValid As Boolean
    
    On Error GoTo ErrHandler
    
    ' validate all objects
    ' Validate the codes first
'    bIsValid = ValidateCategoryCodes(moCategories)
    
    ' NCJ 6 Jan 00 - Do more complete uniqueness check (SR 1798)
    bIsValid = CheckValuesUnique
    
    If bIsValid Then
        For Each oCategory In moCategories
            ' Pick up the changed, invalid categories
            If oCategory.IsDirty And Not oCategory.IsValid Then
                ' The category has been edited and is not valid
                'If Not oCategory.IsDeleted Then
                If Not (oCategory.IsDeleted Or oCategory.IsEmpty) Then
                    ' this is not a deleted or empty category so it should be valid
                    bIsValid = False
                    Exit For
                End If
            End If
        Next oCategory
    End If
    ValidateCategories = bIsValid
        
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ValidateCategories", "clsDDCategories")
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
Public Function GetNextUniqueId(oCategories As Collection, Optional lStartCodeValue As Long = 1, _
                                 Optional lFieldID As FieldID = EFieldIDValueCode) As Long
'---------------------------------------------------------------------
' when setting the ValueCode property on the clsDDCategory the
' code must be unique among all non deleted items
' get the next available unique id
' NCJ 29/12/99 - ID is unique amongst ALL category items (i.e. deleted ones too)
'---------------------------------------------------------------------
Dim lNextUniqueID As Long
    
    On Error GoTo ErrHandler
        
    If oCategories Is Nothing Then
        Set oCategories = moCategories
    End If
    
    lNextUniqueID = lStartCodeValue
    While IsUniqueIDUsed(lNextUniqueID, oCategories, lFieldID)
        lNextUniqueID = lNextUniqueID + 1
    Wend
    
    GetNextUniqueId = lNextUniqueID
        
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "GetNextUniqueID", "clsDDCategories")
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
Private Function IsUniqueIDUsed(lId As Long, oCategories As Collection, _
                                Optional lFieldID As FieldID = EFieldIDValueCode) As Boolean
'---------------------------------------------------------------------
' determine if the unique id is being used in the collection
' NCJ 29/12/99 - Consider deleted items too
' (otherwise we can get key clashes when adding new items)
'---------------------------------------------------------------------
Dim oCategory As clsDDCategory
    
    On Error GoTo ErrHandler
    

    IsUniqueIDUsed = False
    
    For Each oCategory In oCategories
'        If Not oCategory.IsDeleted Then
            If lFieldID = EFieldIDValueCode Then
                If oCategory.ValueCode = lId Then
                    IsUniqueIDUsed = True
                    Exit For
                    
                End If
                
            ElseIf lFieldID = EFieldIDValueOrder Then
                If oCategory.ValueOrder = lId Then
                    IsUniqueIDUsed = True
                    Exit For
                    
                End If
                
            ElseIf lFieldID = EFieldIDValueID Then
                If oCategory.ValueID = lId Then
                    IsUniqueIDUsed = True
                    Exit For
                    
                End If
                
            End If
        
'        End If
        
    Next oCategory
        
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "IsUniqueIDUsed", "clsDDCategories")
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
Private Function AddCategoryToCollection(sCode As String, sLabel As String, _
                                         oCategories As Collection, eErrorLoadOptions As LoadOption) _
                                         As CategoryLoadResult
'---------------------------------------------------------------------
' NCJ 3 Jan 01 - Do not validate categories here because it is much too slow
'           Also do not raise errors but return error code
'---------------------------------------------------------------------
Dim oCategory As clsDDCategory
    
    On Error GoTo ErrHandler
    
    ' the first field is numeric
    With oCategory
        Set oCategory = New clsDDCategory
        With oCategory
            If sCode <> vbNullString Then
                .ValueCode = sCode
            End If
            If sLabel <> vbNullString Then
                .Value = sLabel
            End If
            
            ' only set for categories that are not read from the database
            .IsNew = True
            .IsDeleted = False
            ' NCJ 3 Jan 01 - For sequential records just use the collection count
            .ValueID = oCategories.Count + 1
            .ValueOrder = oCategories.Count + 1
'            .ValueID = GetNextUniqueID(oCategories, , EFieldIDValueID)
'            .ValueOrder = GetNextUniqueID(oCategories, , EFieldIDValueOrder)
        End With
        
        oCategories.Add oCategory
        Set oCategory = Nothing
               
        ' NCJ 3 Jan 01 - Do not validate on each category
        ' because it takes AGES! (Gets done when you click OK anyway)
        AddCategoryToCollection = ENoError
'        If ValidateCategoryCodes(oCategories) Then
'            AddCategoryToCollection = ENoError
'        Else
'            AddCategoryToCollection = ECodesNotUnique
'        End If

        
    End With
        
Exit Function
ErrHandler:
    ' NCJ 3/1/01 return generic error code
    AddCategoryToCollection = EInvalidValue
    
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "AddCategoryToCollection", "clsDDCategories")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select

End Function

'---------------------------------------------------------------------
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
'---------------------------------------------------------------------
' The enumerator to allow For...Each iterations
'---------------------------------------------------------------------
    
    Set NewEnum = moCategories.[_NewEnum]

End Function

'---------------------------------------------------------------------
Public Function CountValid() As Long
'---------------------------------------------------------------------
Dim oCategory As clsDDCategory
Dim lCount As Long
    
    For Each oCategory In moCategories
        If Not oCategory.IsDeleted Then
            lCount = lCount + 1
            
        End If
        
    Next oCategory
    CountValid = lCount
End Function

'---------------------------------------------------------------------
Public Function Count() As Long
'---------------------------------------------------------------------
    
    Count = moCategories.Count

End Function

'---------------------------------------------------------------------
Public Function CountReal() As Long
'---------------------------------------------------------------------
' NCJ 13 Feb 03
' Count of "real" entries, i.e. not deleted or blank
'---------------------------------------------------------------------
Dim oCategory As clsDDCategory
Dim lCount As Long
    
    lCount = 0
    For Each oCategory In moCategories
        If (Not oCategory.IsDeleted) And (Not oCategory.IsEmpty) Then
            lCount = lCount + 1
        End If
    Next oCategory
    
    CountReal = lCount
    
End Function


