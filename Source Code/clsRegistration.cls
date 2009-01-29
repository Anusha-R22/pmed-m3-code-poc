VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsRegistration.cls
'   Copyright:  InferMed Ltd. 2000-2003. All Rights Reserved
'   Author:     Toby Aldridge, November 2000
'   Purpose:    Registration class for MACRO SD
'                   This class should be used to update clsSubjectnumbering
'----------------------------------------------------------------------------------------'
' REVISIONS:
'   NCJ 16 Jun 03 - Mad sure the IsDirty property works correctly (MACRO 3.0 Bug 1004)
'----------------------------------------------------------------------------------------'

Option Explicit

Public Event ValidStartNumber(bIsValid As Boolean)
Public Event ValidNumberWidth(bIsValid As Boolean)
Public Event ValidPrefix(bIsValid As Boolean)
Public Event ValidSuffix(bIsValid As Boolean)

Public Event IsValid(bValid As Boolean)
Public Event HasChanges(bIsDirty As Boolean)

Private moSubjectNumbering As clsSubjectNumbering


'changes have been made
Private mbIsDirty As Boolean

Private WithEvents moBrokenRules As clsBrokenRules
Attribute moBrokenRules.VB_VarHelpID = -1


'----------------------------------------------------------------------------------------'
Public Sub Load(lClinicalTrialId As Long, nVersionId As Integer)
'----------------------------------------------------------------------------------------'

    Call moSubjectNumbering.Load(lClinicalTrialId, nVersionId)
 
    IsDirty = False
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub Save()
'----------------------------------------------------------------------------------------'
    
    Call moSubjectNumbering.Save
                    
    IsDirty = False

End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_Initialize()
'----------------------------------------------------------------------------------------'
'set up default values
'----------------------------------------------------------------------------------------'

    
    Set moSubjectNumbering = New clsSubjectNumbering
    
    Set moBrokenRules = New clsBrokenRules
    
    IsDirty = False     ' NCJ 16 Jun 03 - Changed from TRUE
    
End Sub

'----------------------------------------------------------------------------------------'
Public Property Get StartNumber() As String
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    StartNumber = moSubjectNumbering.StartNumber

End Property

'----------------------------------------------------------------------------------------'
Public Property Let StartNumber(sNumber As String)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    If Val(sNumber) <= LONG_MAX And Val(sNumber) >= LONG_MIN Then
        'ensure no overflow
        If moSubjectNumbering.StartNumber <> CLng(Val(sNumber)) Then
            moSubjectNumbering.StartNumber = CLng(Val(sNumber))
            IsDirty = True
        End If
    End If
    
    If gblnValidString(sNumber, valNumeric) And Val(sNumber) > 0 And Val(sNumber) < LONG_MAX Then
        moBrokenRules.RuleBroken "StartNumber", False
        'raise this event to inform the display that it is valid
        RaiseEvent ValidStartNumber(True)
    Else
        'raise this event to inform the display that it is invalid
        moBrokenRules.RuleBroken "StartNumber", True
        RaiseEvent ValidStartNumber(False)
    End If

    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get NumberWidth() As String
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    NumberWidth = moSubjectNumbering.NumberWidth
    

End Property

'----------------------------------------------------------------------------------------'
Public Property Let NumberWidth(sNumber As String)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    If Val(sNumber) >= INTEGER_MIN And Val(sNumber) <= INTEGER_MAX Then
        'ensure no overflow
        If moSubjectNumbering.NumberWidth <> CInt(Val(sNumber)) Then
            moSubjectNumbering.NumberWidth = CInt(Val(sNumber))
            IsDirty = True
        End If
    End If
    
    If gblnValidString(sNumber, valNumeric) And Val(sNumber) >= 0 And Val(sNumber) <= 255 Then
        moBrokenRules.RuleBroken "NumberWidth", False
        'raise this event to inform the display that it is valid
        RaiseEvent ValidNumberWidth(True)
    Else
        'raise this event to inform the display that it is invalid
        moBrokenRules.RuleBroken "NumberWidth", True
        RaiseEvent ValidNumberWidth(False)
    End If

    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get Prefix() As String
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    Prefix = moSubjectNumbering.Prefix
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Let Prefix(sPrefix As String)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    sPrefix = Trim(sPrefix)
    If moSubjectNumbering.Prefix <> sPrefix Then
        moSubjectNumbering.Prefix = sPrefix
        
        Call PrefixValidate
    
        IsDirty = True
    End If
    
End Property

'---------------------------------------------------------------------
Private Sub PrefixValidate()
'---------------------------------------------------------------------

    With moSubjectNumbering
        If gclmGuideline.IsValidExpression(.Prefix) Or (.Prefix = "" And Not .UsePrefix) Then
            moBrokenRules.RuleBroken "Prefix", False
            'raise this event to inform the display that it is valid
            RaiseEvent ValidPrefix(True)
            
        Else
            'raise this event to inform the display that it is invalid
            moBrokenRules.RuleBroken "Prefix", True
            RaiseEvent ValidPrefix(False)
        End If
    End With
    
End Sub


'----------------------------------------------------------------------------------------'
Public Property Get UsePrefix() As Boolean
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    UsePrefix = moSubjectNumbering.UsePrefix
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Let UsePrefix(bUse As Boolean)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    If moSubjectNumbering.UsePrefix <> bUse Then
        moSubjectNumbering.UsePrefix = bUse
        Call PrefixValidate
        
        IsDirty = True
    End If
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get Suffix() As String
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    Suffix = moSubjectNumbering.Suffix
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Let Suffix(sSuffix As String)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    sSuffix = Trim(sSuffix)
    If moSubjectNumbering.Suffix <> sSuffix Then
        moSubjectNumbering.Suffix = sSuffix
    
        Call SuffixValidate
    
        IsDirty = True
    End If
    
End Property

'---------------------------------------------------------------------
Private Sub SuffixValidate()
'---------------------------------------------------------------------

    With moSubjectNumbering
        If gclmGuideline.IsValidExpression(.Suffix) Or (.Suffix = "" And Not .UseSuffix) Then
            moBrokenRules.RuleBroken "Suffix", False
            'raise this event to inform the display that it is valid
            RaiseEvent ValidSuffix(True)
            
        Else
            'raise this event to inform the display that it is invalid
            moBrokenRules.RuleBroken "Suffix", True
            RaiseEvent ValidSuffix(False)
        End If
    End With
    
End Sub

'----------------------------------------------------------------------------------------'
Public Property Get UseSuffix() As Boolean
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    UseSuffix = moSubjectNumbering.UseSuffix
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Let UseSuffix(bUse As Boolean)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    If moSubjectNumbering.UseSuffix <> bUse Then
        moSubjectNumbering.UseSuffix = bUse
        Call SuffixValidate
    
        IsDirty = True
    End If

End Property

'----------------------------------------------------------------------------------------'
Public Property Get TriggerFormId() As Long
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    TriggerFormId = moSubjectNumbering.TriggerFormId
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Let TriggerFormId(lId As Long)
'----------------------------------------------------------------------------------------'

    If moSubjectNumbering.TriggerFormId <> lId Then
        moSubjectNumbering.TriggerFormId = lId
        IsDirty = True
    End If
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get TriggerVisitId() As Long
'----------------------------------------------------------------------------------------'

    TriggerVisitId = moSubjectNumbering.TriggerVisitId
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Let TriggerVisitId(lId As Long)
'----------------------------------------------------------------------------------------'

    If moSubjectNumbering.TriggerVisitId <> lId Then
        moSubjectNumbering.TriggerVisitId = lId
        
        'use registration if visit has been selected
        moSubjectNumbering.UseRegistration = (lId <> 0)
        
        IsDirty = True
    End If
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get UseRegistration() As Boolean
'----------------------------------------------------------------------------------------'

    UseRegistration = moSubjectNumbering.UseRegistration
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get RegistrationConditions() As clsDataTable
'----------------------------------------------------------------------------------------'

    Set RegistrationConditions = moSubjectNumbering.RegistrationConditions
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Set RegistrationConditions(tblRegistrationConditions As clsDataTable)
'----------------------------------------------------------------------------------------'

    Set moSubjectNumbering.RegistrationConditions = tblRegistrationConditions
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get UniquenessChecks() As clsDataTable
'----------------------------------------------------------------------------------------'

    Set UniquenessChecks = moSubjectNumbering.UniquenessChecks
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Set UniquenessChecks(tblUniquenessChecks As clsDataTable)
'----------------------------------------------------------------------------------------'

    Set moSubjectNumbering.UniquenessChecks = tblUniquenessChecks

End Property

'----------------------------------------------------------------------------------------'
Private Sub moBrokenRules_BrokenRule()
'----------------------------------------------------------------------------------------'

    RaiseEvent IsValid(False)
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub moBrokenRules_NoBrokenRules()
'----------------------------------------------------------------------------------------'

    RaiseEvent IsValid(True)

End Sub

'---------------------------------------------------------------------
Public Property Get IsValid() As Boolean
'---------------------------------------------------------------------

    IsValid = (moBrokenRules.Count = 0)

End Property

'---------------------------------------------------------------------
Public Property Let IsDirty(bIsDirty As Boolean)
'---------------------------------------------------------------------
'public because reg cond and unique expr handled outside class
    
    mbIsDirty = bIsDirty
    RaiseEvent HasChanges(mbIsDirty)
    
End Property

'---------------------------------------------------------------------
Public Property Get IsDirty() As Boolean
'---------------------------------------------------------------------

    IsDirty = mbIsDirty

End Property


