VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CategoryItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: CategoryItem.cls
' Author: Nicky Johns
' Copyright: InferMed, August 2001, All Rights Reserved
' Category Item Read-only Object for MACRO 2.2
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 8 Aug 01 - Initial development
' DPH 24/10/2001 - Added in Active value
'----------------------------------------------------

Option Explicit

Private msCatValue As String
Private msCatCode As String
Private mbCatActive As Boolean

'----------------------------------------------------
Friend Sub SetItem(sCatValue As String, sCatCode As String, bCatActive As Boolean)
'----------------------------------------------------
' Set the item's Value and Code
' (only to be done within MACRO Schedule business object)
'----------------------------------------------------
' DPH 24/10/2001 - Added in active value
'----------------------------------------------------

    msCatValue = sCatValue
    msCatCode = sCatCode
    mbCatActive = bCatActive
    
End Sub

'----------------------------------------------------
Public Property Get Value() As String
Attribute Value.VB_Description = "The category value (read-only)"
'----------------------------------------------------
' The category value (read-only)
'----------------------------------------------------

    Value = msCatValue
    
End Property

'----------------------------------------------------
Public Property Get Code() As String
Attribute Code.VB_Description = "The category code (read-only)"
'----------------------------------------------------
' The category code (read-only)
'----------------------------------------------------

    Code = msCatCode
    
End Property

'----------------------------------------------------
Public Property Get Active() As Boolean
'----------------------------------------------------
' The category active flag (read-only)
'----------------------------------------------------
' DPH 24/10/2001 - Added in active flag
'----------------------------------------------------

    Active = mbCatActive
    
End Property
