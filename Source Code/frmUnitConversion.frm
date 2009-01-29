VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUnitConversion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unit Conversion"
   ClientHeight    =   3870
   ClientLeft      =   2415
   ClientTop       =   2520
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3870
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   4455
      TabIndex        =   8
      Top             =   3375
      Width           =   4455
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   315
         Left            =   2280
         TabIndex        =   10
         Top             =   120
         Width           =   1155
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Accept"
         Height          =   315
         Left            =   420
         TabIndex        =   9
         Top             =   120
         Width           =   1155
      End
   End
   Begin VB.Frame frmConversion 
      Height          =   2715
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   4215
      Begin VB.CommandButton cmdConvert 
         Caption         =   "Convert"
         Height          =   315
         Left            =   3000
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox cboUnits 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   300
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskAlternativeResult 
         Height          =   375
         Left            =   1380
         TabIndex        =   3
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Value to be converted:"
         Height          =   435
         Left            =   180
         TabIndex        =   13
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Convert from:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label lblResult 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1380
         TabIndex        =   7
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Converted value:"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   240
         Index           =   1
         Left            =   3000
         TabIndex        =   5
         Top             =   1920
         Width           =   345
      End
      Begin VB.Line Line1 
         X1              =   60
         X2              =   4140
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label lblConversionFactor 
         BackStyle       =   0  'Transparent
         Height          =   315
         Left            =   360
         TabIndex        =   4
         Top             =   1365
         Width           =   1635
      End
   End
   Begin VB.Label lblUnitClass 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unit Class"
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1395
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Class"
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1155
   End
End
Attribute VB_Name = "frmUnitConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998 - 2005. All Rights Reserved
'   File:       frmUnitConversion.frm
'   Author:     Andrew Newbigging  March 1998
'   Purpose:    Helps user to enter data in different units and calculate the conversion
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'  1           Andrew Newbigging       24/03/98
'  2       Andrew Newbigging       02/04/98
'   PN  20/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   PN  26/09/99    Amended cboUnits_Click() for mtm1.6 changes
' WillC 11/11/99    Added the error handler
'   NCJ 16/11/99    Store CRFElement rather than DataFormat (to enable correct formatting)
'   Mo Morris   17/11/99    DAO to ADO conversion
'  WillC    Changed the Following where present from Integer to Long  ClinicalTrialId
'           CRFPageId,VisitId,CRFElementID
'   NCJ 23/5/00 SR3491 Don't Unload in the Load routine
' MACRO 2.2
'   NCJ 2 Oct 01 - Updated for MACRO 2.2
'   REM 12/10/01 - Fix FromUnits/ToUnits confusion bug
' MACRO 3.0
'   NCJ 24 Mar 05 - Bug 2469 - Use CDbl instead of Val (to solve reg. settings probs.)
'   TA 16/05/2005: Only enable OK button if units have been converted
'------------------------------------------------------------------------------------'

'---------------------------------------------------------------------
Option Explicit
Option Base 0
Option Compare Binary

Private mbResultOK As Boolean
Private mnConversionFactors() As Double
Private mdblResult As Double

Private moElement As eFormElementRO

'---------------------------------------------------------------------
Private Sub cmdCancel_Click()
'---------------------------------------------------------------------

    mbResultOK = False
    Unload Me
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdOK_Click()
'---------------------------------------------------------------------

    mbResultOK = True
    ' NCJ 24 Mar 05 - Bug 2469 - Use CDbl instead of Val (to solve reg. settings probs.)
    mdblResult = CDbl(lblResult.Caption)
    Unload Me

End Sub

'---------------------------------------------------------------------
Public Function Display(oElement As eFormElementRO, ByRef dblResult As Double) As Boolean
'---------------------------------------------------------------------
' Refresh and display ourselves.
' Returns Null if no conversion done, otherwise returns converted value
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsUnits As ADODB.Recordset
Dim bShowMe As Boolean

    On Error GoTo ErrHandler

    bShowMe = False
    mbResultOK = False
    
    Set moElement = oElement
    
    cboUnits.Clear
    Me.BackColor = glFormColour
    ' REM 12/10/01 Changed SELECT ToUnit to FromUnit and WHERE from FromUnit to ToUnit
    sSQL = "SELECT UnitClass, FromUnit, ConversionFactor " _
            & " FROM  UnitConversionFactors " _
            & " WHERE UnitConversionFactors.ToUnit = '" & oElement.Unit & "'"
    
    Set rsUnits = New ADODB.Recordset
    rsUnits.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsUnits.RecordCount = 0 Then
        Call DialogInformation("There are no alternative units defined for: " & oElement.Unit)
    Else
        ' There's something to display
        bShowMe = True
        lblUnitClass.Caption = rsUnits!UnitClass
        lblUnit(1).Caption = oElement.Unit
        mskAlternativeResult.Enabled = False
    
        While Not rsUnits.EOF
        
            cboUnits.AddItem rsUnits!FromUnit 'Changed from ToUnits by REM 12/10/01
            ReDim Preserve mnConversionFactors(cboUnits.ListCount - 1)
            mnConversionFactors(cboUnits.ListCount - 1) = rsUnits!ConversionFactor
            
            rsUnits.MoveNext
        Wend
    End If
    
    ' NCJ 23/5/00 SR3491
    If bShowMe Then
        ' Select first unit
        cboUnits.ListIndex = 0
        'TA 16/05/2005: Only enable OK button if units have been converted
        cmdOK.Enabled = (lblResult.Caption <> "")
        Me.Show vbModal
    Else
        Unload Me
    End If
    
    ' Return the result if it was OK
    If mbResultOK Then
        dblResult = mdblResult
    End If
    Display = mbResultOK
    
Exit Function
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Display")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function

'---------------------------------------------------------------------
Private Sub cboUnits_Click()
'---------------------------------------------------------------------
' They selected a conversion factor
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    ' PN 26/09/99
    ' check that an item has actually been selected
    If cboUnits.ListIndex >= 0 Then
        lblConversionFactor.Caption = cboUnits.ItemData(cboUnits.ListIndex)
        lblConversionFactor.Visible = False
        mskAlternativeResult.Enabled = True
        cmdConvert.Enabled = True
        mskAlternativeResult.Text = ""
        lblResult.Caption = ""
    Else
        cmdConvert.Enabled = False
    End If

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboUnits_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenu.Icon
End Sub

'---------------------------------------------------------------------
Private Sub mskAlternativeResult_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
' Treat RETURN as equivalent to hitting Convert button
'---------------------------------------------------------------------

    If KeyAscii = Asc(vbCr) Then
        cmdConvert_click
    End If

End Sub

'---------------------------------------------------------------------
Private Sub cmdConvert_click()
'---------------------------------------------------------------------
' Try converting the value
'---------------------------------------------------------------------
Dim sNum As String
Dim dblConvertedValue As Double

    On Error GoTo ErrHandler

    sNum = mskAlternativeResult.ClipText

    If IsNumeric(sNum) Then
        'REM 12/10/01 changed the conversion equation from dividing by the conversion factor to multiplying
        ' NCJ 24 Mar 05 - Use CDbl instead of Val (to solve reg. settings probs.)
        dblConvertedValue = CDbl(sNum) * mnConversionFactors(cboUnits.ListIndex)
        lblResult.Caption = Format(dblConvertedValue, moElement.VBFormat)
    Else
        lblResult.Caption = ""
        Call DialogError("This is not a numeric value.")
    End If
    
    'TA 16/05/2005: Only enable OK button if units have been converted
    cmdOK.Enabled = (lblResult.Caption <> "")
    Exit Sub

ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdconvert_click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub
