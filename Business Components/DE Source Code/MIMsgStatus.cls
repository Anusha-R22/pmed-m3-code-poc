VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MIMsgStatus"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'----------------------------------------------------
' File: MIMsgStatus.cls
' Author: Toby Aldridge
' Copyright: InferMed, August 2002, All Rights Reserved
' MIMsgStatus Enumerations MACRO Data Entry
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'TA 02/09/2002: NoteStatus changed so as not to include comments
'--------------------------------------


'NB VALUE ORDER IN ENUMERATION IS USED FOR CALCULATIONS
'eg if an object has two discepancies one 'clsoed' and one 'raised', 'raised is used as has the higher value


'discrepancy status for icons
Public Enum eDiscrepancyStatus
    dsNone = 0
    dsClosed = 10
    dsResponded = 20
    dsRaised = 30
End Enum

'SDV status for icons
Public Enum eSDVStatus
    ssNone = 0
    ssCancelled = 10
    ssComplete = 20
    ssPlanned = 30
    ssQueried = 40
End Enum

'note (comment also) status for icons
Public Enum eNoteStatus
    nsNoNote = 0
    nsHasNote = 1
End Enum
