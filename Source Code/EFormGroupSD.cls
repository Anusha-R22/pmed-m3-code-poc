VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EFormGroupSD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------
'   File:       EFormGroupSD.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Richard Meinesz, November 2001
'   Purpose:    EForm Group Object for MACRO SD
'--------------------------------------------------------------------------------------------------
' Revisions:
' NCJ 28 Nov 01 - Fixed bug in InitialRows Let
' NCJ 3 Dec 01 - Added Rows property
' NCJ 1 Mar 02 - Rewrote validation code for Min, Max and Init
'           Introduced max of 999 for integer values
' NCJ 1 Nov 02 - Restrict Initial rows and Display rows to max 99
'--------------------------------------------------------------------------------------------------

Option Explicit

Public Event IsValid(bValid As Boolean)

Public Event DisplayRowsValid(bValid As Boolean)
Public Event InitialRowsValid(bValid As Boolean)
Public Event MinRepeatsValid(bValid As Boolean)
Public Event MaxRepeatsValid(bValid As Boolean)

' These are used by the SD user interface when drawing/maintaining the group
Public ControlIndex As Integer
Public ElementID As Integer
Public Rows As Integer
Public RowHeight As Single
Public RowWidth As Single
Public HeaderHeight As Single
Public ColWidths As Collection

' Private copies of variables
Private mlStudyID As Long
Private mnVersionId As Integer
Private mlCRFPageId As Long
Private mlQGroupID As Long
Private mbBorder As Boolean
Private mnDisplayRows As Integer
Private mnInitialRows As Integer
Private mnMinRepeats As Integer
Private mnMaxRepeats As Integer
Private mcolElements As Collection
Private mbIsChanged As Boolean
Private mbIsNew As Boolean

Private mbBorderStored As Boolean
Private mnDisplayRowsStored As Integer
Private mnInitialRowsStored As Integer
Private mnMinRepeatsStored As Integer
Private mnMaxRepeatsStored As Integer

Private WithEvents moBrokenRules As clsBrokenRules
Attribute moBrokenRules.VB_VarHelpID = -1

'--------------------------------------------------------------------------------------------------
Public Sub Store()
'--------------------------------------------------------------------------------------------------
' REM 28/11/01
' Stores the values in the object when the sub is called
'--------------------------------------------------------------------------------------------------

    mbBorderStored = mbBorder
    mnDisplayRowsStored = mnDisplayRows
    mnInitialRowsStored = mnInitialRows
    mnMinRepeatsStored = mnMinRepeats
    mnMaxRepeatsStored = mnMaxRepeats
    
End Sub

'--------------------------------------------------------------------------------------------------
Public Sub Restore()
'--------------------------------------------------------------------------------------------------
' REM 28/11/01
' Restores the values in the object to those that were stored
'--------------------------------------------------------------------------------------------------'

    mbBorder = mbBorderStored
    mnDisplayRows = mnDisplayRowsStored
    mnInitialRows = mnInitialRowsStored
    mnMinRepeats = mnMinRepeatsStored
    mnMaxRepeats = mnMaxRepeatsStored

End Sub

'--------------------------------------------------------------------------------------------------
Public Property Get StudyID() As Long
'--------------------------------------------------------------------------------------------------

    StudyID = mlStudyID

End Property

'--------------------------------------------------------------------------------------------------
Public Property Get VersionId() As Integer
'--------------------------------------------------------------------------------------------------

    VersionId = mnVersionId

End Property

'--------------------------------------------------------------------------------------------------
Public Property Get CRFPageId() As Long
'--------------------------------------------------------------------------------------------------

    CRFPageId = mlCRFPageId

End Property

'--------------------------------------------------------------------------------------------------
Public Property Get QGroupID() As Long
'--------------------------------------------------------------------------------------------------

    QGroupID = mlQGroupID
    
End Property

'--------------------------------------------------------------------------------------------------
Public Property Get Border() As Boolean
'--------------------------------------------------------------------------------------------------

    Border = mbBorder

End Property

'--------------------------------------------------------------------------------------------------
Public Property Let Border(bBorder As Boolean)
'--------------------------------------------------------------------------------------------------

    If bBorder <> mbBorder Then
        mbBorder = bBorder
        mbIsChanged = True
    End If

End Property

'--------------------------------------------------------------------------------------------------
Public Property Get DisplayRows() As Integer
'--------------------------------------------------------------------------------------------------

    DisplayRows = mnDisplayRows
    
End Property

'--------------------------------------------------------------------------------------------------
Public Property Let DisplayRows(nDisplayRows As Integer)
'--------------------------------------------------------------------------------------------------
' REM 28/11/01
' Validation for return values
' NCJ 1 Mar 02 - Changed validation
'--------------------------------------------------------------------------------------------------

    ' We save it even if it violates the rules
    If nDisplayRows <> mnDisplayRows Then
        mnDisplayRows = nDisplayRows
        mbIsChanged = True
        Call RaiseDispRowsEvent(Not ValidRowValue(nDisplayRows))
    End If

End Property

'--------------------------------------------------------------------------------------------------
Public Property Get InitialRows() As Integer
'--------------------------------------------------------------------------------------------------
' The number of rows to be displayed initially for a new group instance
'--------------------------------------------------------------------------------------------------

    InitialRows = mnInitialRows

End Property

'--------------------------------------------------------------------------------------------------
Public Property Let InitialRows(nInitialRows As Integer)
'--------------------------------------------------------------------------------------------------'
' REM 28/11/01
' Set the value for InitialRows
' NCJ 1 Mar 02 - Changed validation
'--------------------------------------------------------------------------------------------------
    
    ' We save it even if it violates the rules
    If nInitialRows <> mnInitialRows Then
        mnInitialRows = nInitialRows
        mbIsChanged = True
        ' Check the cross-validations
        Call CheckValidations
    End If

End Property

'--------------------------------------------------------------------------------------------------
Public Property Get MinRepeats() As Integer
'--------------------------------------------------------------------------------------------------

    MinRepeats = mnMinRepeats

End Property

'--------------------------------------------------------------------------------------------------
Public Property Let MinRepeats(nMinRepeats As Integer)
'--------------------------------------------------------------------------------------------------
' REM 28/11/01
' Set the value for the min repeats
' NCJ 1 Mar 02 - Changed validation
'--------------------------------------------------------------------------------------------------

    ' We save it even if it violates the rules
    If nMinRepeats <> mnMinRepeats Then
        mnMinRepeats = nMinRepeats
        mbIsChanged = True
        ' Check the cross-validations
        Call CheckValidations
    End If

End Property

'--------------------------------------------------------------------------------------------------
Public Property Get MaxRepeats() As Integer
'--------------------------------------------------------------------------------------------------

    MaxRepeats = mnMaxRepeats

End Property

'--------------------------------------------------------------------------------------------------
Public Property Let MaxRepeats(ByVal nMaxRepeats As Integer)
'--------------------------------------------------------------------------------------------------
' REM 28/11/01
' Validation of Max Repeats value
' NCJ 1 Mar 02 - Changed validation
'--------------------------------------------------------------------------------------------------
    
    ' We save it even if it violates any rules
    If nMaxRepeats <> mnMaxRepeats Then
        mnMaxRepeats = nMaxRepeats
        mbIsChanged = True
        ' Check whether it's invalid
        Call CheckValidations
    End If

End Property


'--------------------------------------------------------------------------------------------------
Private Sub RaiseMinRptEvent(bRuleBroken As Boolean)
'--------------------------------------------------------------------------------------------------
' Raise the events that say whether the MinRepeats is valid.
'--------------------------------------------------------------------------------------------------
        
        Call moBrokenRules.RuleBroken("MinRepeats", bRuleBroken)
        RaiseEvent MinRepeatsValid(Not bRuleBroken)

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub RaiseMaxRptEvent(bRuleBroken As Boolean)
'--------------------------------------------------------------------------------------------------
' Raise the events that say whether the MaxRepeats is valid.
'--------------------------------------------------------------------------------------------------
        
        Call moBrokenRules.RuleBroken("MaxRepeats", bRuleBroken)
        RaiseEvent MaxRepeatsValid(Not bRuleBroken)

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub RaiseInitRowsEvent(bRuleBroken As Boolean)
'--------------------------------------------------------------------------------------------------
' Raise the events that say whether the InitialRows is valid
'--------------------------------------------------------------------------------------------------
        
        Call moBrokenRules.RuleBroken("InitialRows", bRuleBroken)
        RaiseEvent InitialRowsValid(Not bRuleBroken)

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub RaiseDispRowsEvent(bRuleBroken As Boolean)
'--------------------------------------------------------------------------------------------------
' Raise the events that say whether the DisplayRows is valid
'--------------------------------------------------------------------------------------------------
        
        Call moBrokenRules.RuleBroken("DisplayRows", bRuleBroken)
        RaiseEvent DisplayRowsValid(Not bRuleBroken)

End Sub

'--------------------------------------------------------------------------------------------------
Public Property Get Elements() As Collection
'--------------------------------------------------------------------------------------------------

    Set Elements = mcolElements

End Property

'--------------------------------------------------------------------------------------------------
Public Property Get IsChanged() As Boolean
'--------------------------------------------------------------------------------------------------

    IsChanged = mbIsChanged

End Property

'--------------------------------------------------------------------------------------------------
Public Sub Init(lStudyID As Long, nVersionId As Integer, lCRFPageId As Long, lQGroupId As Long, bIsNew As Boolean)
'--------------------------------------------------------------------------------------------------
' REM 16/11/01
' Initialises the modular level variables and sets new collections
'--------------------------------------------------------------------------------------------------

    mlStudyID = lStudyID
    mnVersionId = nVersionId
    mlQGroupID = lQGroupId
    mlCRFPageId = lCRFPageId
    
    Set mcolElements = New Collection

    Set ColWidths = New Collection
    
    Set moBrokenRules = New clsBrokenRules
    
    mbIsChanged = False
    mbIsNew = bIsNew
    
    'if is new then sets initial values
    If bIsNew Then
        mbBorder = True
        mnDisplayRows = 1
        mnInitialRows = 1
        mnMinRepeats = 1
        mnMaxRepeats = 1
    End If
    
End Sub

'--------------------------------------------------------------------------------------------------
Public Sub Save()
'--------------------------------------------------------------------------------------------------
' REM 16/11/01
' Saves a new EFormQGroup or updates a current one
'--------------------------------------------------------------------------------------------------
Dim sSQL As String
Dim nBorder As Integer

    On Error GoTo ErrLabel
    
    'Change Border property from a Boolean to an Integer
    If mbBorder = True Then
        nBorder = 1
    Else
        nBorder = 0
    End If
    
    If mbIsNew Then
        'Insert a new record into the EFormQGroup table
        sSQL = "INSERT INTO EFormQGroup" & _
            "(ClinicalTrialID, VersionID,CRFPageID,QGroupID,Border,DisplayRows,InitialRows,MinRepeats,MaxRepeats)" & _
            " VALUES (" & mlStudyID & "," & mnVersionId & "," & mlCRFPageId & "," & mlQGroupID & "," & nBorder & "," & mnDisplayRows & "," & mnInitialRows & "," & mnMinRepeats & "," & mnMaxRepeats & ")"
        MacroADODBConnection.Execute sSQL
        
        mbIsNew = False
    
    ElseIf mbIsChanged Then
        'Update the EFormQGroup table with the new values if they changed
        sSQL = "UPDATE EFormQGroup " & _
               " SET Border = " & nBorder & _
               ", DisplayRows = " & mnDisplayRows & _
               ", InitialRows = " & mnInitialRows & _
               ", MinRepeats = " & mnMinRepeats & _
               ", MaxRepeats = " & mnMaxRepeats & _
               " WHERE ClinicalTrialID = " & Me.StudyID & _
               " AND VersionID = " & Me.VersionId & _
               " AND CRFPageID = " & mlCRFPageId & _
               " AND QGroupID = " & mlQGroupID
        MacroADODBConnection.Execute sSQL
    End If

    mbIsChanged = False

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormGroupSD.Save"
    
End Sub

'--------------------------------------------------------------------------------------------------
Public Sub AddElement(lCRFElementID As Long)
'--------------------------------------------------------------------------------------------------
' REM 16/11/01
' Add a CRFElement to this eFormGroup's collection
'--------------------------------------------------------------------------------------------------
    
    'Ignore errors if CRFElementId already exists
    On Error Resume Next
    mcolElements.Add lCRFElementID, Str(lCRFElementID)

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
'--------------------------------------------------------------------------------------------------

    Set mcolElements = Nothing
    Set ColWidths = Nothing

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub moBrokenRules_BrokenRule()
'--------------------------------------------------------------------------------------------------

    RaiseEvent IsValid(False)

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub moBrokenRules_NoBrokenRules()
'--------------------------------------------------------------------------------------------------

    RaiseEvent IsValid(True)

End Sub

'--------------------------------------------------------------------------------------------------
Private Function ValidRepeatValue(nValue As Integer) As Boolean
'--------------------------------------------------------------------------------------------------
' Is this a valid value for Min or Max repeats?
' Must be greater than 0 and no more than 999
'--------------------------------------------------------------------------------------------------

    ValidRepeatValue = (nValue > 0) And (nValue < 1000)

End Function

'--------------------------------------------------------------------------------------------------
Private Function ValidRowValue(nValue As Integer) As Boolean
'--------------------------------------------------------------------------------------------------
' Is this a valid row value?
' Must be greater than 0 and no more than 99
' NCJ 1 Nov 02 - restrict rows to 0-99 (previously 999)
'--------------------------------------------------------------------------------------------------

    ValidRowValue = (nValue > 0) And (nValue < 100)

End Function

'-------------------------------------------------------------------------------
Private Sub CheckValidations()
'-------------------------------------------------------------------------------
' NCJ 1 Mar 02
' Check the validations on each of the three interdependent fields
'   Max Repeats
'   Min Repeats
'   Initial Rows
' individually and raise appropriate events
'-------------------------------------------------------------------------------

    Call RaiseMaxRptEvent(MaxRepeatsInvalid)
    Call RaiseMinRptEvent(MinRepeatsInvalid)
    Call RaiseInitRowsEvent(InitRowsInvalid)
    
End Sub

'--------------------------------------------------------------------------------------------------
Private Function MaxRepeatsInvalid() As Boolean
'--------------------------------------------------------------------------------------------------
' NCJ 1 Mar 02
' Returns TRUE if MaxRepeats is invalid
'--------------------------------------------------------------------------------------------------

    ' Assume invalid unless proved otherwise
    MaxRepeatsInvalid = True
    
    If Not ValidRepeatValue(mnMaxRepeats) Then Exit Function
    
    ' Check against Min repeats (if Min repeats a valid integer)
    If ValidRepeatValue(mnMinRepeats) Then
        If mnMaxRepeats < mnMinRepeats Then
            Exit Function
        End If
    End If
    
    ' Check against Init rows (if Init rows a valid integer)
    If ValidRowValue(mnInitialRows) Then
        If mnMaxRepeats < mnInitialRows Then
            Exit Function
        End If
    End If
   
    ' If we get here then it wasn't invalid
    MaxRepeatsInvalid = False

End Function

'--------------------------------------------------------------------------------------------------
Private Function MinRepeatsInvalid() As Boolean
'--------------------------------------------------------------------------------------------------
' NCJ 1 Mar 02
' Returns TRUE if MinRepeats is invalid
'--------------------------------------------------------------------------------------------------

    ' Assume invalid unless proved otherwise
    MinRepeatsInvalid = True
    
    If Not ValidRepeatValue(mnMinRepeats) Then Exit Function
    
    ' Check against Max repeats (if Max repeats a valid integer)
    If ValidRepeatValue(mnMaxRepeats) Then
        If mnMaxRepeats < mnMinRepeats Then
            Exit Function
        End If
    End If
    
    ' If we get here then it wasn't invalid
    MinRepeatsInvalid = False

End Function

'--------------------------------------------------------------------------------------------------
Private Function InitRowsInvalid() As Boolean
'--------------------------------------------------------------------------------------------------
' NCJ 1 Mar 02
' Returns TRUE if InitRows is invalid
'--------------------------------------------------------------------------------------------------

    ' Assume invalid unless proved otherwise
    InitRowsInvalid = True
    
    If Not ValidRowValue(mnInitialRows) Then Exit Function
    
    ' Check against Max repeats (if Max repeats a valid integer)
    If ValidRepeatValue(mnMaxRepeats) Then
        If mnMaxRepeats < mnInitialRows Then
            Exit Function
        End If
    End If
    
    ' If we get here then it wasn't invalid
    InitRowsInvalid = False

End Function


