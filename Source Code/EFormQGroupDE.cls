VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EFormQGroupDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       EFormQGroupDE.cls
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   Author:     Nicky Johns, November 2001
'   Purpose:    eFormQGroup object for eForm data entry for Data Management in MACRO 3.0.
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'
' REVISIONS:
'   NCJ 12-14 Nov 01 - Initial development
'   MLM 09/09/02: Added public collection so that we can remember the group's EFormItems.
'   NCJ 17 Sept 02 - Store Ids instead of DEBS objects
'   NCJ 24 Sept 02 - Deal with there being no QGroup instance
'   NCJ 10 Oct 02 - Added ColXs to store X coords of each group column
'   NCJ 27 Nov 02 - Added NextTabIndex to store tab index for new group rows
'   NCJ 3 Dec 02 - That doesn't work, so instead, store last eFormItem built
'----------------------------------------------------------------------------------------'

Option Explicit

Private moSubject As StudySubject

'Public QGI As QGroupInstance
Private mlEFormTaskId As Long
Private moQGroup As QGroupRO

Public ControlIndex As Integer
Public RowHeight As Single
Public RowWidth As Single
Public HeaderHeight As Single
Public ColWidths As Collection
' NCJ 10 Oct 02
Public ColXs As Collection

'MLM 09/09/02:
Public EFormItems As Collection

' NCJ 27 Nov 02
Public NextTabIndex As Integer

' NCJ 3 Dec 02 - Last item created (so we can get its tab index)
Public LastItem As EFormItem
Public LastPushButtonIndex As Integer
Public GroupIsBuilt As Boolean  ' Has group been built?

'------------------------------------------------------------------------'
Public Sub Init(oSubject As StudySubject, oEFI As EFormInstance, oQGroup As QGroupRO)
'------------------------------------------------------------------------'
' NCJ 17 Sept 02
' Store what we need to know about this Group
'------------------------------------------------------------------------'

    Set moSubject = oSubject
    mlEFormTaskId = oEFI.EFormTaskId
    Set moQGroup = oQGroup
    GroupIsBuilt = False
    LastPushButtonIndex = 0
    Set LastItem = Nothing
    
End Sub

'------------------------------------------------------------------------'
Public Property Get QGI() As QGroupInstance
'------------------------------------------------------------------------'
' The Group Instance, retrieved by its ID
' NB May be Nothing if it's a read-only Requested eForm
'------------------------------------------------------------------------'

    Set QGI = EFI.QGroupInstanceById(moQGroup.QGroupId)

End Property

'------------------------------------------------------------------------'
Public Property Get QGroup() As QGroupRO
'------------------------------------------------------------------------'
' The Group Definition
'------------------------------------------------------------------------'

    Set QGroup = moQGroup

End Property

'------------------------------------------------------------------------'
Public Property Get QGroupId() As Long
'------------------------------------------------------------------------'
' The Group ID
'------------------------------------------------------------------------'

    QGroupId = moQGroup.QGroupId

End Property

'------------------------------------------------------------------------'
Public Property Get CurrentRows() As Integer
'------------------------------------------------------------------------'
' The no. of rows in this group instance
' NB QG Instance may not exist yet...
'------------------------------------------------------------------------'

    If Not QGI Is Nothing Then
        CurrentRows = QGI.Rows
    Else
        CurrentRows = moQGroup.InitialRows
    End If
    
End Property

'--------------------------------------------------------------------------
Public Property Get EFI() As EFormInstance
'--------------------------------------------------------------------------
' NCJ 17 Sept 02
' The eFormInstance for this eForm item,
' retrieved by EFormTaskId
'--------------------------------------------------------------------------

    Set EFI = moSubject.eFIByTaskId(mlEFormTaskId)

End Property

'------------------------------------------------------------------------'
Private Sub Class_Initialize()
'------------------------------------------------------------------------'

    Set ColWidths = New Collection
    Set ColXs = New Collection      ' NCJ 10 Oct 02
    RowHeight = 0
    RowWidth = 0
    HeaderHeight = 0
    'MLM 09/09/02:
    Set EFormItems = New Collection
    ' NCJ 3 Dec 02
    Set LastItem = Nothing
    GroupIsBuilt = False
    
End Sub

'------------------------------------------------------------------------'
Private Sub Class_Terminate()
'------------------------------------------------------------------------'

    Set ColWidths = Nothing
    Set ColXs = Nothing      ' NCJ 10 Oct 02
    'MLM 09/09/02:
    Set EFormItems = Nothing
    
    Set moQGroup = Nothing
    Set moSubject = Nothing
    Set LastItem = Nothing

End Sub
