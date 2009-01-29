VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MultiSelectElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   File:       MultiSelectElement.cls
'   Author:     Ronald Schravendeel, August 2002
'   Purpose:    Elements used to refer to selected CRFelements used by the
'               MultiSelect (Lasso) function.
'-----------------------------------------------------------------------------------
' REVISIONS
'   RS  10/9/2002   Renamed member CRFElementIndex (was CRFElelemenentIndex)
'                   Added CRFelementID member
'-----------------------------------------------------------------------------------
Option Explicit

Public CRFElementIndex As Integer     ' Index of selected CRFelement
Public CRFelementID As Long             ' CRFElement ID
Public CRFElementDataItemID As Integer  ' Corresponding DataItemID
Public Element As CRFElement            ' Reference, only used for Copy/Paste
Public CaptionSelected As Boolean       ' Indicates if Caption is part of selection
Public ElementSelected As Boolean       ' Indicates if CRFcontrol is part of selection
Public ElementType As Long              ' CRF Element Type

Public X As Single                      ' X coordinate when selected
Public Y As Single                      ' Y coordinate when selected

Public MaxX As Single                   ' Maximum allowed X coordinate for this control
Public MaxY As Single                   ' Maximum allowed Y coordinate for this control

Public CaptionX As Single               ' X coordinate of Caption when selected
Public CaptionY As Single               ' Y coordinate of Caption when selected




