VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MIMsgDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: MIMsgDS.cls
' Toby Aldridge, InferMed, Nov 2001
' Data Services MIMessage class for MACRO
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' TA Nov 01 - Initial development
' ic 10/07/2002 added vbtypes to ConvertFromNull() function call in Load() function
' RS 30/09/2002 Added TimezoneOffset column support
' NCJ 11 Oct 02 - Fixed bug in Load (ResponseCycle)
'----------------------------------------------------

Option Explicit

Private Const msDELIMITER = "|"

' The table names
Private Const msMIMESSAGE_TABLE = "MIMessage"

' The filter columns for Discrepancies and SDVs
Private Const msOBJECT_KEY_COLS = "MIMessageObjectId|MIMessageObjectSource|MIMessageSite"

' The filter columns for nOTES
Private Const msNOTE_KEY_COLS = "MIMessageId|MIMessageSource|MIMessageSite"

'for calculating new ids
Private Const msNEWID_COL = "MIMessageId"
Private Const msNEWID_FILTER_COLS = "MIMessageSource|MIMessageSite"

'for calculating new object ids
Private Const msNEWOBJECTID_COL = "MIMessageObjectId"
Private Const msNEWOBJECTID_FILTER_COLS = "MIMessageObjectSource|MIMessageSite"

' MIMessage cols
Private Const msMIMESSAGE_PRIKEY_COLS = "MIMessageId|MIMessageSite|MIMessageSource"
Private Const msMIMESSAGE_NONKEY_COLS = "MIMessageType|MIMessageScope" _
                                    & "|MIMessageObjectId|MIMessageObjectSource|MIMessagePriority|MIMessageTrialName" _
                                    & "|MIMessagePersonId|MIMessageVisitId|MIMessageVisitCycle|MIMessageCRFPageTaskId" _
                                    & "|MIMessageResponseTaskId|MIMessageResponseValue|MIMessageOCDiscrepancyId" _
                                    & "|MIMessageCreated|MIMessageSent|MIMessageReceived" _
                                    & "|MIMessageHistory|MIMessageProcessed|MIMessageStatus" _
                                    & "|MIMessageText|MIMessageUserName|MIMessageUserNameFull" _
                                    & "|MIMessageResponseTimeStamp|MIMessageResponseCycle" _
                                    & "|MIMessageCreated_TZ|MIMessageSent_TZ|MIMessageReceived_TZ" _
                                    & "|MIMESSAGECRFPAGEID|MIMESSAGECRFPAGECYCLE|MIMESSAGEDATAITEMID"
                                    
Private Const msMIMESSAGE_COLS = msMIMESSAGE_PRIKEY_COLS & msDELIMITER & msMIMESSAGE_NONKEY_COLS

Private Const msORDER_COL1 = "MIMessageHistory"
Private Const msORDER_COL2 = "MIMessageCreated"

'--------------------------------------------------------------------
Public Function Load(ByVal sConnection As String, ByVal lId As Long, ByVal nSource As Integer, ByVal sSite As String, bNote As Boolean) As String
'--------------------------------------------------------------------
' Load the given study subject and return info as (property bag) string
' sConnection is the DB connection string.
'--------------------------------------------------------------------
' ic 10/07/2002 added vbtypes to ConvertFromNull() function call
'--------------------------------------------------------------------
Dim oState As MIMsgState
Dim oQS As QueryServer
Dim oQD As QueryDef
Dim vData As Variant
Dim i As Long
Dim colStates As Collection

    On Error GoTo ErrorLabel
    
    ' Create the Query Server and Select Defn objects
    Set oQS = New QueryServer
    Set oQD = New QueryDef

    ' Set up the Tables, the Columns, the Joins and the Filters
    If bNote Then
        'only one row required
        oQD.InitSelect msMIMESSAGE_TABLE, Split(msMIMESSAGE_COLS, msDELIMITER), Split(msNOTE_KEY_COLS, msDELIMITER), Array(lId, nSource, sSite)
    Else
        oQD.InitSelect msMIMESSAGE_TABLE, Split(msMIMESSAGE_COLS, msDELIMITER), Split(msOBJECT_KEY_COLS, msDELIMITER), Array(lId, nSource, sSite)
        oQD.QueryOrders.Add msORDER_COL1, True
        oQD.QueryOrders.Add msORDER_COL2
    End If
        
    oQS.Init sConnection
    oQS.ConnectionOpen
    
    'always follow a connection open with a begin trans
    oQS.BeginTrans
    On Error GoTo ErrorLabelTrans
    
    ' Get the array of values from the database
    vData = oQS.SelectArray(oQD)
    oQS.Commit
    On Error GoTo ErrorLabel
    
    Set oQS = Nothing       ' This closes connection
    Set oQD = Nothing
    
    Set colStates = New Collection
    
    'vData will always contain at least one row
    For i = 0 To UBound(vData, 2)
        Set oState = New MIMsgState
        ' Fill in values from Row 0 of our arrays
        'ic 10/07/2002 added vbtypes to ConvertFromNull() function call
        With oState
            .Id = vData(0, i)
            .Site = vData(1, i)
            .Source = vData(2, i)
            .MIMessageType = vData(3, i)
            .Scope = vData(4, i)
            .ObjectId = vData(5, i)
            .ObjectSource = vData(6, i)
            .Priority = vData(7, i)
            .StudyName = vData(8, i)
            .SubjectId = ConvertFromNull(vData(9, i), vbInteger)
            .VisitId = ConvertFromNull(vData(10, i), vbLong)
            .VisitCycle = ConvertFromNull(vData(11, i), vbInteger)
            .EFormTaskId = ConvertFromNull(vData(12, i), vbLong)
            .ResponseTaskId = ConvertFromNull(vData(13, i), vbLong)
            .ResponseValue = ConvertFromNull(vData(14, i), vbString)
            .ExternalId = vData(15, i)
            .TimeCreated = vData(16, i)
            .TimeSent = vData(17, i)
            .TimeReceived = vData(18, i)
            .History = vData(19, i)
            .Processed = vData(20, i)
            .Status = vData(21, i)
            .Text = ConvertFromNull(vData(22, i), vbString)
            .UserName = vData(23, i)
            .UserNameFull = vData(24, i)
            .ResponseTimeStamp = vData(25, i)
            .ResponseCycle = vData(26, i)       ' NCJ 11 Oct 02 - Changed from 1
            
            ' RS 30/09/2002: TimezoneOffset columns
            .TimeCreatedTimezoneOffset = ConvertFromNull(vData(27, i), vbInteger)
            .TimeSentTimezoneOffset = ConvertFromNull(vData(28, i), vbInteger)
            .TimeReceivedTimezoneOffset = ConvertFromNull(vData(29, i), vbInteger)
            
            .EFormId = ConvertFromNull(vData(30, i), vbLong)
            .EFormCycle = ConvertFromNull(vData(31, i), vbInteger)
            .QuestionId = ConvertFromNull(vData(32, i), vbLong)
                        
            ' Return the state as a string
            colStates.Add oState
        End With
        
    Next
    
    Load = CollectionSerialise(colStates)
    
    Exit Function
    
ErrorLabel:
    Set oQS = Nothing
    Err.Raise Err.Number, , Err.Description & "|" & "MIMsgDS.Load"
    Exit Function
    
ErrorLabelTrans:

    oQS.Rollback
    Set oQS = Nothing
    Err.Raise Err.Number, , Err.Description & "|" & "MIMsgDS.Load"
    Exit Function
    
End Function

'--------------------------------------------------------------------
Public Function Save(ByVal sConnection As String, sState As String) As String
'--------------------------------------------------------------------
' Save the given MIMessage state and return info as (property bag) string
' sConnection is the DB connection string.
'--------------------------------------------------------------------
Dim oState As MIMsgState
Dim oQS As QueryServer
Dim oQD As QueryDef
Dim oNewIDQD As QueryDef
Dim vData As Variant
Dim i As Long
Dim colStates As Collection
Dim colNewStates As Collection
Dim lNewId As Long
Dim vValues As Variant
Dim bCalculateObject As Boolean
Dim lObjectId As Long
Dim nObjectSource As Integer

        
    On Error GoTo ErrLabel
    
    Set colStates = CollectionDeSerialise(sState)
    Set colNewStates = New Collection
    
    Set oQS = New QueryServer
    
    oQS.ConnectionOpen sConnection
    
    On Error GoTo ErrLabelTrans
    oQS.BeginTrans
    
    For i = 1 To colStates.Count
        Set oState = colStates(i)
        With oState
            If bCalculateObject Then
                'we have had our object id and source calulated
                .ObjectId = lObjectId
                .ObjectSource = nObjectSource
            End If
            
            If .IsNew Then
                
                If i = 1 Then
                    'first message and is new - we need to create an object id and source
                    bCalculateObject = True
                    nObjectSource = .Source ' this tells us if we are on the site or the server
                    Set oNewIDQD = New QueryDef
                    oNewIDQD.InitNewId msMIMESSAGE_TABLE, msNEWOBJECTID_COL, Split(msNEWOBJECTID_FILTER_COLS, msDELIMITER), Array(nObjectSource, .Site)
                    lObjectId = oQS.QueryDefNewId(oNewIDQD)
                End If
            
                If bCalculateObject Then
                    'we have had our object id and source calculated
                    .ObjectId = lObjectId
                    .ObjectSource = nObjectSource
                End If
                
                'insert row
                Set oNewIDQD = New QueryDef
                Set oQD = New QueryDef
                oNewIDQD.InitNewId msMIMESSAGE_TABLE, msNEWID_COL, Split(msNEWID_FILTER_COLS, msDELIMITER), Array(.Source, .Site)
                .Id = oQS.QueryDefNewId(oNewIDQD)
                vValues = Array(.Id, .Site, .Source, _
                                .MIMessageType, .Scope, _
                                .ObjectId, .ObjectSource, .Priority, .StudyName, _
                                .SubjectId, .VisitId, .VisitCycle, .EFormTaskId, _
                                .ResponseTaskId, .ResponseValue, .ExternalId, _
                                .TimeCreated, .TimeSent, .TimeReceived, _
                                .History, .Processed, .Status, _
                                ConvertToNull(.Text, vbString), .UserName, .UserNameFull, .ResponseTimeStamp, .ResponseCycle, _
                                ConvertToNull(.TimeCreatedTimezoneOffset, vbInteger), _
                                ConvertToNull(.TimeSentTimezoneOffset, vbInteger), _
                                ConvertToNull(.TimeReceivedTimezoneOffset, vbInteger), _
                                ConvertToNull(.EFormId, vbLong), _
                                ConvertToNull(.EFormCycle, vbInteger), _
                                ConvertToNull(.QuestionId, vbLong))
                                
                oQD.InitSave msMIMESSAGE_TABLE, Split(msMIMESSAGE_COLS, msDELIMITER), vValues
                oQS.QueryInsert oQD
            Else
                If oState.IsChanged Then
                    'update row
                    Set oQD = New QueryDef
                    vValues = Array(.MIMessageType, .Scope, _
                                .ObjectId, .ObjectSource, .Priority, .StudyName, _
                                .SubjectId, .VisitId, .VisitCycle, .EFormTaskId, _
                                .ResponseTaskId, .ResponseValue, .ExternalId, _
                                .TimeCreated, .TimeSent, .TimeReceived, _
                                .History, .Processed, .Status, _
                                ConvertToNull(.Text, vbString), .UserName, .UserNameFull, .ResponseTimeStamp, .ResponseCycle, _
                                .TimeCreatedTimezoneOffset, .TimeSentTimezoneOffset, .TimeReceivedTimezoneOffset, _
                                ConvertToNull(.EFormId, vbLong), _
                                ConvertToNull(.EFormCycle, vbInteger), _
                                ConvertToNull(.QuestionId, vbLong))
                                
                    oQD.InitSave msMIMESSAGE_TABLE, Split(msMIMESSAGE_NONKEY_COLS, msDELIMITER), vValues, Split(msMIMESSAGE_PRIKEY_COLS, msDELIMITER), Array(oState.Id, oState.Site, oState.Source)
                    oQS.QueryUpdate oQD
                End If
            End If
        End With
        colNewStates.Add oState
    Next
        
    oQS.Commit
    
    On Error GoTo ErrLabel
    Set oQS = Nothing
        
    Save = CollectionSerialise(colNewStates)
    
    Exit Function
    
ErrLabel:
    Set oQS = Nothing
    Err.Raise Err.Number, , Err.Description & "|" & "MIMsgDS.Save(" & i & ")"
    Exit Function
    
ErrLabelTrans:

    oQS.Rollback
    Set oQS = Nothing
    Err.Raise Err.Number, , Err.Description & "|" & "MIMsgDS.Save(" & i & ")"
    Exit Function
        
End Function
