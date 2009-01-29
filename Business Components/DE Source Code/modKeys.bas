Attribute VB_Name = "modKeys"

'----------------------------------------------------
' File: modKeys.bas
' Nicky Johns, InferMed, May 2001
' Key generation routines MACRO 2.2
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 21-29 May 01 - Initial development
'----------------------------------------------------

Private Const msDELIMITER = "|"

Option Explicit

''--------------------------------------------------------------------
'Public Function VisitInstanceKey(ByVal lVisitId As Long, ByVal nCycleNo As Integer) As String
''--------------------------------------------------------------------
'' Create key comprising VisitId and CycleNo
''--------------------------------------------------------------------
'
'    VisitInstanceKey = Str(lVisitId) & msDELIMITER & Str(nCycleNo)
'
'End Function

'--------------------------------------------------------------------
Public Function VisitEFormKey(ByVal lVisitId As Long, _
                            ByVal lEFormId As Long) As String
'--------------------------------------------------------------------
' Create key comprising VisitId and eFormID
'--------------------------------------------------------------------

    VisitEFormKey = Str(lVisitId) & msDELIMITER & Str(lEFormId)
    
End Function

'--------------------------------------------------------------------
Public Function VisitEFormInstanceKey(ByVal lEFormId As Long, _
                            ByVal lEFormTaskId As Long) As String
'--------------------------------------------------------------------
' Create key comprising eFormID and eFormTaskId (which may be 0)
'--------------------------------------------------------------------

    VisitEFormInstanceKey = Str(lEFormId) & msDELIMITER & Str(lEFormTaskId)
    
End Function

'--------------------------------------------------------------------
Public Function ScheduleVisitKey(ByVal lVisitId As Long, _
                            ByVal lVisitTaskId As Long) As String
'--------------------------------------------------------------------
' Create key comprising lVisitId and lVisitTaskId (which may be 0)
'--------------------------------------------------------------------

    ScheduleVisitKey = Str(lVisitId) & msDELIMITER & Str(lVisitTaskId)
    
End Function


