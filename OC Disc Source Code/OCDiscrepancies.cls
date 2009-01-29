VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OCDiscrepancies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------'
'   File:       OCDiscpreancies.cls
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Toby Aldridge, March 2002
'   Purpose:    Interface class for automating cutting and pasting of
'                   discrepancies from OC
'----------------------------------------------------------------------------------------'
' Revisions:
'ta 27/05/2002 CBB 2.2.13.47: if form is already displayed when just restore it when disaply is called
'----------------------------------------------------------------------------------------'

Option Explicit

'they want to open  a subject
Public Event OpenSubject(sStudy As String, sSite As String, sSubjectLabel As String)
 
'the details for the currently selected OC discrepancy
Private mlOCId As Long
Private msDText As String

'is the OC List open?
Private mbFormOpen As Boolean

'link to OC list instance
Private WithEvents mofrmOC As frmOC
Attribute mofrmOC.VB_VarHelpID = -1


'----------------------------------------------------------------------------------------'
Private Sub Class_Initialize()
'----------------------------------------------------------------------------------------'
'initialise reference to OC list form
    Set mofrmOC = New frmOC
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_Terminate()
'----------------------------------------------------------------------------------------'
   If mbFormOpen Then
        Unload mofrmOC
        Set mofrmOC = Nothing
    End If
End Sub

'----------------------------------------------------------------------------------------'
Private Sub mofrmOC_Closed()
'----------------------------------------------------------------------------------------'
    mbFormOpen = False
End Sub

'----------------------------------------------------------------------------------------'
Public Property Get FormOpen() As Boolean
'----------------------------------------------------------------------------------------'
'is the discrepancy list displayed?
    FormOpen = mbFormOpen
End Property

'----------------------------------------------------------------------------------------'
Private Sub mofrmOC_OCChoose(lOCId As Long, sDiscText As String)
'----------------------------------------------------------------------------------------'
    mlOCId = lOCId
    msDText = sDiscText
End Sub

'----------------------------------------------------------------------------------------'
Private Sub mofrmOC_OpenSubject(sStudy As String, sSite As String, sSubjectLabel As String, _
                                lOCId As Long, sDiscText As String)
'----------------------------------------------------------------------------------------'
    mlOCId = lOCId
    msDText = sDiscText
    RaiseEvent OpenSubject(sStudy, sSite, sSubjectLabel)
End Sub

'----------------------------------------------------------------------------------------'
Public Sub Display(sCon As String, oUser As MACROUser)
'----------------------------------------------------------------------------------------'

    'ta 27/05/2002 CBB 2.2.13.47: if already displayed just restore it
    If mbFormOpen Then
        mofrmOC.WindowState = vbNormal
    Else
        mofrmOC.Display sCon, oUser
    End If
    mbFormOpen = True
End Sub

'----------------------------------------------------------------------------------------'
Public Property Get OCId() As Long
'----------------------------------------------------------------------------------------'
    OCId = mlOCId
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get DiscrepancyText() As String
'----------------------------------------------------------------------------------------'
    DiscrepancyText = msDText

End Property

'----------------------------------------------------------------------------------------'
Public Sub MarkAsRaised(lOCId As Long)
'----------------------------------------------------------------------------------------'

Dim oItem As ListItem

    If mbFormOpen Then
        mofrmOC.MarkAsRaised lOCId
    End If
End Sub
