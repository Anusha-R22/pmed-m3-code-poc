VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       clsTextList.cls
'   Author:     Paul Norris 03/08/99
'   Purpose:    Class to wrap up the a list of items in the database. This class does
'               nothing more than map a list of ids onto a list of strings. It removes
'               the need to store the ids in the combobox control, thus simplifying the UI.
'
'               NB. Use only with combobox controls that have the style set to Dropdown list
'               When calling the Load mehtod and specifying bAddBlank as True, be sure that
'               the table does not (and will not) contain an entry with id = 0.
'
'               Sample use:
'                Set moCategories = New clsTextList
'                moCategories.Load "TrialTypes"
'               The text list can then be used to populate comboboxes etc.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   Paul Norris     23/08/99    Amended Fetch(),Item() and Key() to allow use
'                               of elements with key 0
'   Paul Norris     03/09/99    Amended Fetch() and Load() to accept database object as parameter
'                               This allows reading form the security database as well as macro
'   Paul Norris     06/09/99    Amended Key() to remove K from key using BeforeStr() rather than Val
'                               This allows for string key values
'   PN  30/09/99    Commented code more thoroughly and converted DAO to ADO
'   NCJ 11 Feb 00   Changed argument type in AddItem
'
'----------------------------------------------------------------------------------------'
Option Explicit

' this is the collection of strings
Private moList As Collection

' this is the collection of ids
Private moKeys As Collection
'----------------------------------------------------------------------------------------'
Public Function Count() As Long
'----------------------------------------------------------------------------------------'
' return the number of entries in the list
'----------------------------------------------------------------------------------------'
    Count = moList.Count
End Function

'----------------------------------------------------------------------------------------'
Public Sub AddItem(sText As String, sIndex As String)
Attribute AddItem.VB_MemberFlags = "40"
'----------------------------------------------------------------------------------------'
' add an individual item
' this allows for populating without calling the load method or adding ids and strings
' after calling the load method
' NCJ 11 Feb 00 - Changed type of Index from Integer to String
'----------------------------------------------------------------------------------------'
    moList.Add Format$(sText), Format$(sIndex & "K")
    moKeys.Add Format$(sIndex & "K"), Format$(sText)
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Fetch(sListName As String, bAddBlank As Boolean, oDatabase As ADODB.Connection)
'----------------------------------------------------------------------------------------'
' this does the actual read form the database
' it will also add a blank to the list with id = 0 if bAddBlank =  True
' sListName can be of the form
'       'Select TrialTypeID, TrialTypeName From TrialType'
' but must always contain the id field first
'----------------------------------------------------------------------------------------'
Dim rsList As ADODB.Recordset
    
  On Error GoTo ErrHandler
  
    ' add a blank item then add the db items
    Set rsList = New ADODB.Recordset
    rsList.Open sListName, oDatabase, adOpenKeyset, adLockReadOnly, adCmdText
    
    ' PN change 1
    ' use a 'K' in the key at the end to ensure that numeric keys are not
    ' used as indexes
    ' (this will cause a problem for elements with numeric key of 0)
    
    If bAddBlank Then
        ' add a blank value to the list
        ' this will assume that 0 is not used in the existing id list
        moList.Add vbNullString, "0K"
        moKeys.Add "0K", vbNullString
    End If

    ' add all records in the recordset
    ' the assumption is made that the id field is the first field
    ' and the string field is the second field
    Do While Not rsList.EOF
        ' add the strings
        moList.Add Format$(rsList(1)), Format$(rsList(0) & "K")
        
        ' add the ids
        moKeys.Add Format$(rsList(0) & "K"), Format$(rsList(1))
        rsList.MoveNext
    Loop

    rsList.Close
    Set rsList = Nothing
        
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Fetch", "clsTextList")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select


End Sub
'----------------------------------------------------------------------------------------'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
'----------------------------------------------------------------------------------------'
' routine for iteration handling ofthe for each syntax
' proc id = -4
'----------------------------------------------------------------------------------------'
    Set NewEnum = moList.[_NewEnum]
End Function
'----------------------------------------------------------------------------------------'
Public Function Item(ByVal vIndex As Variant) As String
Attribute Item.VB_UserMemId = 0
'----------------------------------------------------------------------------------------'
' return the string that corresponds to vIndex as a key
'----------------------------------------------------------------------------------------'
    ' append the 'K' for the correct key
    Item = moList.Item(vIndex & "K")
End Function
'----------------------------------------------------------------------------------------'
Public Function Key(ByVal sText As Variant) As String
'----------------------------------------------------------------------------------------'
' return the id for the string sText
'----------------------------------------------------------------------------------------'
    ' return only the numeric part of the key
    ' PN 06/09/99
    Key = BeforeStr(moKeys.Item(sText), 1)
End Function
'----------------------------------------------------------------------------------------'
Public Sub Load(sListName As String, _
                Optional bAddBlank As Boolean = False, _
                Optional oDatabase As ADODB.Connection)
'----------------------------------------------------------------------------------------'
' load data from Macro.mdb as default otherwise use oDatabase
'----------------------------------------------------------------------------------------'
    
  On Error GoTo ErrHandler
  
    ' PN 03/09/99
    ' load the data from the database
    If oDatabase Is Nothing Then
        ' use the default (which is macro)
        Set oDatabase = MacroADODBConnection
    End If
    Call Fetch(sListName, bAddBlank, oDatabase)
        
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "load", "clsTextList")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub
'----------------------------------------------------------------------------------------'
Private Sub Class_Initialize()
'----------------------------------------------------------------------------------------'
    Set moList = New Collection
    Set moKeys = New Collection
End Sub
'----------------------------------------------------------------------------------------'
Private Sub Class_Terminate()
'----------------------------------------------------------------------------------------'
    Set moList = Nothing
    Set moKeys = Nothing
End Sub
