VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCodeForChanging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       clsCodeForChanging.cls
'   Author:     Mo Morris 18/2/00
'   Purpose:    class for holding information about codes that needed to be changed during
'               a session of Arezzo Re-Build
'               clsCodeForChanging elements are held in collection clsCodesForChanging
'               Only used by modArezzoReBuild
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'
'----------------------------------------------------------------------------------------'
'local variable(s) to hold property value(s)
Private mvarOldCode As String 'local copy
Private mvarNewCode As String 'local copy
Private mvarCodeType As String 'local copy
'local variable(s) to hold property value(s)
Private mvarCodeId As Long 'local copy

'---------------------------------------------------------------------
Public Property Let CodeId(ByVal vData As Long)
'---------------------------------------------------------------------
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CodeId = 5
'---------------------------------------------------------------------

    mvarCodeId = vData

End Property

'---------------------------------------------------------------------
Public Property Get CodeId() As Long
'---------------------------------------------------------------------
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CodeId
'---------------------------------------------------------------------

    CodeId = mvarCodeId

End Property

'---------------------------------------------------------------------
Public Property Let CodeType(ByVal vData As String)
'---------------------------------------------------------------------
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CodeType = 5
'---------------------------------------------------------------------

    mvarCodeType = vData

End Property

'---------------------------------------------------------------------
Public Property Get CodeType() As String
'---------------------------------------------------------------------
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CodeType
'---------------------------------------------------------------------

    CodeType = mvarCodeType

End Property

'---------------------------------------------------------------------
Public Property Let NewCode(ByVal vData As String)
'---------------------------------------------------------------------
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NewCode = 5
'---------------------------------------------------------------------

    mvarNewCode = vData

End Property

'---------------------------------------------------------------------
Public Property Get NewCode() As String
'---------------------------------------------------------------------
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NewCode
'---------------------------------------------------------------------

    NewCode = mvarNewCode

End Property

'---------------------------------------------------------------------
Public Property Let OldCode(ByVal vData As String)
'---------------------------------------------------------------------
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OldCode = 5
'---------------------------------------------------------------------

    mvarOldCode = vData

End Property

'---------------------------------------------------------------------
Public Property Get OldCode() As String
'---------------------------------------------------------------------
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OldCode
'---------------------------------------------------------------------

    OldCode = mvarOldCode

End Property



