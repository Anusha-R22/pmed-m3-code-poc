VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCodesForChanging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsCodeForChanging"
Attribute VB_Ext_KEY = "Member0" ,"clsCodeForChanging"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       clsCodesForChanging.cls
'   Author:     Mo Morris 18/2/00
'   Purpose:    Collection class for clsCodeForChanging
'               Only used by modArezzoReBuild
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'
'----------------------------------------------------------------------------------------'

'local variable to hold collection
Private mCol As Collection

'---------------------------------------------------------------------
Public Function Add(OldCode As String, NewCode As String, CodeType As String, CodeId As Long) As clsCodeForChanging
'---------------------------------------------------------------------
'create a new object
'---------------------------------------------------------------------
Dim objNewMember As clsCodeForChanging

    Set objNewMember = New clsCodeForChanging

    'set the properties passed into the method
    objNewMember.OldCode = OldCode
    objNewMember.NewCode = NewCode
    objNewMember.CodeType = CodeType
    objNewMember.CodeId = CodeId
    mCol.Add objNewMember, Str(CodeId)

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

'---------------------------------------------------------------------
Public Property Get Item(vntIndexKey As Variant) As clsCodeForChanging
Attribute Item.VB_UserMemId = 0
'---------------------------------------------------------------------
'used when referencing an element in the collection
'vntIndexKey contains either the Index or Key to the collection,
'this is why it is declared as a Variant
'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
'---------------------------------------------------------------------

    Set Item = mCol(vntIndexKey)

End Property

'---------------------------------------------------------------------
Public Property Get Count() As Long
'---------------------------------------------------------------------
'used when retrieving the number of elements in the
'collection. Syntax: Debug.Print x.Count
'---------------------------------------------------------------------

    Count = mCol.Count

End Property

'---------------------------------------------------------------------
Public Sub Remove(vntIndexKey As Variant)
'---------------------------------------------------------------------
'used when removing an element from the collection
'vntIndexKey contains either the Index or Key, which is why
'it is declared as a Variant
'Syntax: x.Remove(xyz)
'---------------------------------------------------------------------

    mCol.Remove vntIndexKey

End Sub

'---------------------------------------------------------------------
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'---------------------------------------------------------------------
'this property allows you to enumerate
'this collection with the For...Each syntax
'---------------------------------------------------------------------

    Set NewEnum = mCol.[_NewEnum]

End Property

'---------------------------------------------------------------------
Private Sub Class_Initialize()
'---------------------------------------------------------------------
'creates the collection when this class is created
'---------------------------------------------------------------------

    Set mCol = New Collection

End Sub

'---------------------------------------------------------------------
Private Sub Class_Terminate()
'---------------------------------------------------------------------
'destroys collection when this class is terminated
'---------------------------------------------------------------------

    Set mCol = Nothing

End Sub

'---------------------------------------------------------------------
Public Function ChangedCode(sOriginalCode As String, lSearchId As Long) As String
'---------------------------------------------------------------------
'Returns the New code if the original code has been changed
'or the Original Code if it has not been changed
'---------------------------------------------------------------------
Dim oCodeToChange As clsCodeForChanging

    On Error Resume Next
    
    Set oCodeToChange = mCol.Item(Str(lSearchId))
    
    If Err.Number <> 0 Then
        ChangedCode = sOriginalCode
    Else
        ChangedCode = oCodeToChange.NewCode
    End If

End Function
