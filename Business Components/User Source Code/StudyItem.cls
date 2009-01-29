VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------'
'   File:       StudyItem.cls
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Toby Aldridge, November 2002
'   Purpose:    Study Item class
'----------------------------------------------------------------------------------------'
' Revisions:

Option Explicit


Public Enum eStudyObject
    soVisit
    soeform
    soQuestion
End Enum

Public StudyId As Long
Public Id As Long
Public Code As String
Public Name As String
Public StudyObject As eStudyObject

'----------------------------------------------------------------------------------------'
Public Property Get Key() As String
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    Key = GetKey(StudyId, StudyObject, Id)

End Property

'----------------------------------------------------------------------------------------'
Public Function GetKey(lStudyId As Long, enStudyObject As eStudyObject, lId As Long) As String
'----------------------------------------------------------------------------------------'

    GetKey = "k" & lStudyId & "|" & enStudyObject & "|" & lId
    
End Function

'----------------------------------------------------------------------------------------'
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'----------------------------------------------------------------------------------------'

    StudyId = PropBag.ReadProperty("s")
    Id = PropBag.ReadProperty("i")
    Code = PropBag.ReadProperty("c")
    Name = PropBag.ReadProperty("n")
    StudyObject = PropBag.ReadProperty("so")
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------------------------------------------'

    PropBag.WriteProperty "s", StudyId
    PropBag.WriteProperty "i", Id
    PropBag.WriteProperty "c", Code
    PropBag.WriteProperty "n", Name
    PropBag.WriteProperty "so", StudyObject
    
End Sub
