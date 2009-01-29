VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValueStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = ""
Attribute VB_Description = ""
Attribute VB_Description = ""
Attribute VB_Description = ""
Attribute VB_Description = ""
'----------------------------------------------------
' File: ValueStore.cls
' Copyright: InferMed 2001, All Rights Reserved
' Author: Nicky Johns, InferMed, June 2001
' Purpose: Handles changing class properties (in Response class)
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 25 June 01 - Initial development
'----------------------------------------------------

Option Explicit

' Current, Previous and Original values
Private mvCurrent As Variant
Private mvPrevious As Variant
Private mvOriginal As Variant

'----------------------------------------------------
Public Sub Saved()
Attribute Saved.VB_Description = "Call Saved when values have been saved, so move Current value to Original and Previous values"
'----------------------------------------------------
' Call Saved when values have been saved,
' so move Current value to Original and Previous values
'----------------------------------------------------

    mvOriginal = mvCurrent
    mvPrevious = mvCurrent
    
End Sub

'----------------------------------------------------
Public Property Get Dirty() As Boolean
Attribute Dirty.VB_Description = "TRUE if Current value differs from Original value"
'----------------------------------------------------
' TRUE if Current value differs from Original value
'----------------------------------------------------

    Dirty = (mvCurrent <> mvOriginal)

End Property

'----------------------------------------------------
Public Property Let OriginalValue(vValue As Variant)
'----------------------------------------------------
' Setting the OriginalValue also sets the Current and Previous
'----------------------------------------------------
    
    mvOriginal = vValue
    mvCurrent = vValue
    mvPrevious = vValue

End Property

'----------------------------------------------------
Public Property Get OriginalValue() As Variant
Attribute OriginalValue.VB_Description = "Original value"
'----------------------------------------------------
' Original value
'----------------------------------------------------

    OriginalValue = mvOriginal

End Property

'----------------------------------------------------
Public Property Get CurrentValue() As Variant
Attribute CurrentValue.VB_Description = "The current value"
'----------------------------------------------------
' The current value
'----------------------------------------------------

    CurrentValue = mvCurrent

End Property

'----------------------------------------------------
Public Property Let CurrentValue(vValue As Variant)
'----------------------------------------------------
' Set the Current value if it's different from existing Current
' Automatically moves existing Current into Previous
'----------------------------------------------------

    If vValue <> mvCurrent Then
        mvPrevious = mvCurrent
        mvCurrent = vValue
    End If
    
End Property

'----------------------------------------------------
Public Property Get PreviousValue() As Variant
Attribute PreviousValue.VB_Description = "Get the previous value (read-only)"
'----------------------------------------------------
' Get the previous value (read-only)
'----------------------------------------------------

    PreviousValue = mvPrevious

End Property

 
 
 
 
 
