VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EFormItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   File:       EFormItem.cls
'   Author:     Nicky Johns, November 2001
'   Purpose:    EFormItem object to represent a data entry field
'               on an eForm in MACRO 3.0 and later versions.
'-----------------------------------------------------------------------------------
' REVISIONS
'
' MLM 06/09/02: To cope with displaying items from different forms, each item should
'   remember its Response.
' NCJ 17 Sept 02 - Do not store Response/EFI but store their IDs
' NCJ 24 Sept 02 - Added Init routine. Remember Response may be nothing
' TA 6/1/03: added tooltip text property for an item
'-----------------------------------------------------------------------------------

Option Explicit

' Note that for non-group items, RowNo will be 0
Public RowNo As Integer
Public ControlIndex As Integer

Private moElement As eFormElementRO
Private moSubject As StudySubject

' NCJ 17 Sept 02 - Store Ids rather than objects
Private mlResponseId As Long
Private mnRepeatNo As Integer
Private mlEFormTaskId As Long

'--------------------------------------------------------------------------
Public Sub Init(oSubject As StudySubject, oEFI As EFormInstance, oElement As eFormElementRO)
'--------------------------------------------------------------------------
' Initialise this eForm Item with given Subject, EFI and Element
'--------------------------------------------------------------------------
    
    Set moSubject = oSubject
    mlEFormTaskId = oEFI.EFormTaskId
    Set moElement = oElement

End Sub

'--------------------------------------------------------------------------
Public Property Get Element() As eFormElementRO
'--------------------------------------------------------------------------
' The eForm Element for this eForm item
'--------------------------------------------------------------------------

    Set Element = moElement

End Property

'--------------------------------------------------------------------------
Public Property Get EFI() As EFormInstance
'--------------------------------------------------------------------------
' The eFormInstance for this eForm item
'--------------------------------------------------------------------------

    Set EFI = moSubject.eFIByTaskId(mlEFormTaskId)

End Property

'--------------------------------------------------------------------------
Public Property Get Response() As Response
'--------------------------------------------------------------------------
' The response for this eForm item, retrieved by ID
' NCJ 24 Sept 02 - There may not be a response
'--------------------------------------------------------------------------

    If mlResponseId > 0 Then
        Set Response = EFI.Responses.ResponseByResponseId(mlResponseId, mnRepeatNo)
    End If

End Property

'--------------------------------------------------------------------------
Public Property Set Response(oResponse As Response)
'--------------------------------------------------------------------------
' The response for this eForm item
' We don't store the response but we store the Ids of the response and its EFI
' (because the object pointers may change while we have an eForm open)
'--------------------------------------------------------------------------

    If Not oResponse Is Nothing Then
        mlResponseId = oResponse.ResponseId
        mnRepeatNo = oResponse.RepeatNumber
    End If

End Property


'---------------------------------------------------------------------
Public Property Get TooltipText() As String
'---------------------------------------------------------------------
' Get the tooltip text for the item
'---------------------------------------------------------------------
Dim sText As String

    sText = " Question: " & moElement.Name
    
    If Not Response Is Nothing Then
        With Response
            'MLM 09/06/03: Only show comments if user is allowed to see them.
            If .Comments <> "" And goUser.CheckPermission(gsFnViewIComments) Then
                'if ther are comments show them
                sText = sText & " - " & "Comments: " & Replace(Replace(.Comments, vbCrLf & vbCrLf, ".  "), vbCrLf, " ")
            Else
                'show help text
                If moElement.HelpText <> "" Then
                    sText = sText & " - " & moElement.HelpText
                End If
            End If
        End With
    Else
        'show help text
        If moElement.HelpText <> "" Then
            sText = sText & " - " & moElement.HelpText
        End If
    End If

    TooltipText = sText
    
End Property

'---------------------------------------------------------------------
Public Property Get StatusTooltipText() As String
'---------------------------------------------------------------------
' Get the status tooltip text for the item
'---------------------------------------------------------------------
Dim sText As String


    sText = " Question: " & moElement.Name

    If Not Response Is Nothing Then
        With Response
            If .ValidationMessage <> "" Then
                sText = sText & " - " & .StatusString & ": " & .ValidationMessage
            End If
            If .Status = eStatus.OKWarning Then
                sText = sText & ". " & "Overrule reason: " & .OverruleReason & "."
            End If
        End With
        
    End If
    StatusTooltipText = sText
    
End Property

'--------------------------------------------------------------------------
Private Sub Class_Initialize()
'--------------------------------------------------------------------------
        
        mlResponseId = 0
        mnRepeatNo = 1

End Sub

'--------------------------------------------------------------------------
Private Sub Class_Terminate()
'--------------------------------------------------------------------------

    Set moElement = Nothing
    Set moSubject = Nothing

End Sub
