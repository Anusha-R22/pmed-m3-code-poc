VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QGroupRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2001-2004. All Rights Reserved
'   File:       QGroupRO.cls
'   Author:     Nicky Johns, October 2001
'   Purpose:    Question Group object to be used
'               for Repeating Question Groups
'               in MACRO DE 3.0 and later versions.
'-----------------------------------------------------------------------------------
' REVISIONS
'   NCJ 31 Oct-5 Nov 01 - Initial development
'   NCJ 23 Nov 01 - SetElement
'   NCJ 24 Mar 04 - Clear out collections in Terminate
'-----------------------------------------------------------------------------------

Option Explicit

Private moElements As Collection
Private moQGroupState As QGroupState
Private moEForm As eFormRO
Private moElement As eFormElementRO

'-------------------------------------------------------------------
Friend Sub SetState(oGroupState As QGroupState, oEForm As eFormRO)
'-------------------------------------------------------------------
' Set the state
'-------------------------------------------------------------------

    Set moQGroupState = oGroupState
    Set moEForm = oEForm
    
End Sub

'-------------------------------------------------------------------
Friend Sub SetElement(oElement As eFormElementRO)
'-------------------------------------------------------------------
' The eFormElement that this group belongs to
'-------------------------------------------------------------------

    Set moElement = oElement
    
End Sub

'-------------------------------------------------------------------
Public Property Get QGroupID() As Long
'-------------------------------------------------------------------
' Unique integer ID of group
'-------------------------------------------------------------------

    QGroupID = moQGroupState.QGroupID

End Property

'-------------------------------------------------------------------
Public Property Get ElementID() As Long
'-------------------------------------------------------------------
' eForm Element ID of group on eForm
'-------------------------------------------------------------------

    ElementID = moElement.ElementID

End Property

'-------------------------------------------------------------------
Public Property Get Element() As eFormElementRO
'-------------------------------------------------------------------
' eForm Element ID of group on eForm
'-------------------------------------------------------------------

    Set Element = moElement

End Property

'-------------------------------------------------------------------
Public Property Get Code() As String
'-------------------------------------------------------------------
' Unique text code of group
'-------------------------------------------------------------------

    Code = moQGroupState.Code

End Property

'-------------------------------------------------------------------
Public Property Get Name() As String
'-------------------------------------------------------------------
' Name of group
'-------------------------------------------------------------------

    Name = moQGroupState.Name

End Property

'-------------------------------------------------------------------
Public Property Get DisplayType() As Integer
'-------------------------------------------------------------------
' Display type of group (autofit, user-defined, etc.)
'-------------------------------------------------------------------

    DisplayType = moQGroupState.DisplayType

End Property

'-------------------------------------------------------------------
Public Property Get DisplayRows() As Integer
'-------------------------------------------------------------------
' Number of rows to display at a time
'-------------------------------------------------------------------

    DisplayRows = moQGroupState.DisplayRows

End Property

'-------------------------------------------------------------------
Public Property Get InitialRows() As Integer
'-------------------------------------------------------------------
' Number of rows to display initially
'-------------------------------------------------------------------

    InitialRows = moQGroupState.InitialRows

End Property

'-------------------------------------------------------------------
Public Property Get MinRepeats() As Integer
'-------------------------------------------------------------------
' Minimum number of repeats required for group success
'-------------------------------------------------------------------

    MinRepeats = moQGroupState.MinRepeats

End Property

'-------------------------------------------------------------------
Public Property Get MaxRepeats() As Integer
'-------------------------------------------------------------------
' Maximum number of repeats allowed for group
'-------------------------------------------------------------------

    MaxRepeats = moQGroupState.MaxRepeats

End Property

'-------------------------------------------------------------------
Public Property Get Border() As Boolean
'-------------------------------------------------------------------
' Whether border should be displayed
'-------------------------------------------------------------------

    Border = moQGroupState.Border

End Property

'-------------------------------------------------------------------
Public Property Get Elements() As Collection
'-------------------------------------------------------------------
' Collection of eFormElementRO objects which belong to the group
'-------------------------------------------------------------------

    Set Elements = moElements

End Property

'-------------------------------------------------------------------
Public Property Get eForm() As eFormRO
'-------------------------------------------------------------------
' The eForm to which this group belongs
'-------------------------------------------------------------------

    Set eForm = moEForm

End Property


'-------------------------------------------------------------------
Private Sub Class_Initialize()
'-------------------------------------------------------------------

    Set moElements = New Collection
    
End Sub

'-------------------------------------------------------------------
Private Sub Class_Terminate()
'-------------------------------------------------------------------
' Tidy up when we leave
'-------------------------------------------------------------------

    Call Terminate

End Sub

'-------------------------------------------------------------------
Friend Sub Terminate()
'-------------------------------------------------------------------
' Tidy up when we leave
'-------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    Call CollectionRemoveAll(moElements)
    Set moElements = Nothing
    
    Set moQGroupState = Nothing
    
    ' Don't terminate eForm or element because we're being called from the eForm
    Set moEForm = Nothing
    
    Set moElement = Nothing

    
Exit Sub
ErrLabel:
'    Debug.Print "Error in Terminating QGroup "

End Sub
