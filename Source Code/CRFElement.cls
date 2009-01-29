VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRFElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       CRFElement.cls (used to be clsCRFElement.cls)
'   Copyright:  InferMed Ltd. 1998-2003. All Rights Reserved
'   Author:     Andrew Newbigging, June 1997
'   Purpose:    Represents a CRF element in MACRO SD
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'  Andrew Newbigging       4/06/97 - 13/10/98
'   TA/NCJ 25/4/00 - 3/4/01 Various amendments
' MACRO 3.0 - CRFElement now only used in Study Design
'   REM 15/11/01 - Removed all code related to data entry
' NCJ 26-28 Nov 01 - Addition of several new properties & methods for SD
' NCJ 4 Dec 01 - More stuff moved here from frmCRFDesign
' NCJ 3 Jan 02 - Bug fix to Relocate
' ZA 19/07/2002 - added functionality for font/colour for caption & comment
' ZA 19/09/2002 - added display length property
' NCJ 7 Nov 02 - Added Hotlink property
' NCJ 1 Jul 03 - Added Expandable property
' ic 14/06/2005 added clinical coding
'------------------------------------------------------------------------------------

Option Explicit

Public ClinicalTrialId As Long
Public VersionId As Integer
Public CRFPageId As Long
Public CRFelementID As Integer
Public Caption As String
Public DataItemId As Long
Public DataItemCode As String
Public DataItemName As String
Public ControlType As Integer
Public DataItemType As Integer
Public DataItemLength As Integer
Public DataItemFormat As String

' NCJ 26 Nov 01 - New properties added
Public FieldOrder As Integer
Public CaptionX As Long
Public CaptionY As Long
Public ElementX As Long
Public ElementY As Long
Public Unit As String
Public ShowStatusFlag As Boolean
Public Hidden As Boolean
Public FontName As String
Public FontSize As Integer
Public FontBold As Boolean
Public FontItalic As Boolean
Public FontColour As Long
Public OwnerQGroupID As Long
Public QGroupID As Long
Public QGroupFieldOrder As Integer

Private moCategories As clsDDCategories
' The VB form with which it's associated (assume frmCRFDesign)
Private moVBForm As Form
Private msSQLWhere As String

Private Const ms_UPDATE_CRFELEMENT = "UPDATE CRFElement SET "

' ZA 19/07/2002 - added font and colour properties of caption
Public CaptionFontName As String
Public CaptionFontSize As Integer
Public CaptionFontBold As Boolean
Public CaptionFontItalic As Boolean
Public CaptionFontColour As Long

'ZA 21/08/2002
Public ElementUse As eElementUse

'ZA 19/09/2002
Public DisplayLength As Integer

' NCJ 7 Nov 02
Public Hotlink As String

'TA 18/06/2003
Public Description As String


'------------------------------------------------------
Public Property Get Categories() As clsDDCategories
'------------------------------------------------------
' The category values for this element
' Always freshly loaded when this property is read
'------------------------------------------------------

    Set moCategories = Nothing
    Set moCategories = New clsDDCategories
    Call moCategories.Load(DataItemId, ClinicalTrialId, VersionId)
    Set Categories = moCategories
    
End Property

'------------------------------------------------------
Private Sub Class_Initialize()
'------------------------------------------------------
' Initialise some values
'------------------------------------------------------

    DataItemId = 0
    Caption = ""
    DataItemCode = ""
    DataItemName = ""
    QGroupID = 0
    OwnerQGroupID = 0
    QGroupFieldOrder = 0
    FieldOrder = 0

End Sub

'------------------------------------------------------
Private Function SQLWhereClause() As String
'------------------------------------------------------
' Get the WHERE clause for updating this CRFElement
'------------------------------------------------------
Dim sSQL As String

    sSQL = " WHERE ClinicalTrialId = " & Me.ClinicalTrialId _
         & " AND   VersionId       = " & Me.VersionId _
         & " AND   CRFPageId       = " & Me.CRFPageId _
         & " AND   CRFElementId    = " & Me.CRFelementID

    SQLWhereClause = sSQL

End Function

'------------------------------------------------------
Public Sub SetOwnerQGroup(lOwnerGroupID As Long, nGroupFieldOrder As Integer, nFieldOrder As Integer)
'------------------------------------------------------
' Set this to be a group member by setting owner group and group field order
' nFieldOrder is the field order of its owning group
' Default status flag to 0, and make sure ControlType is allowable
'------------------------------------------------------
Dim sSQL As String

    OwnerQGroupID = lOwnerGroupID
    QGroupFieldOrder = nGroupFieldOrder
    FieldOrder = nFieldOrder
    ShowStatusFlag = False
    
    ' For group members, we can only have certain control types in AutoFit
    ' (Rethink this when we have UserDefined groups)
    If DataItemId > 0 Then
        Select Case ControlType
        Case gn_OPTION_BUTTONS, gn_PUSH_BUTTONS
            ControlType = gn_POPUP_LIST
        Case Else
            'Leave as is
        End Select
    End If
    
    ' Do the database update
    sSQL = ms_UPDATE_CRFELEMENT _
            & " OwnerQGroupID = " & lOwnerGroupID _
            & " , QGroupFieldOrder = " & nGroupFieldOrder _
            & " , FieldOrder = " & nFieldOrder _
            & " , ControlType = " & ControlType _
            & " , ShowStatusFlag =  0 " _
            & SQLWhereClause
    
    MacroADODBConnection.Execute sSQL

End Sub

'-----------------------------------------------------------------------------------
Public Sub Relocate(ByVal lNewX As Long, ByVal lNewY As Long, _
                    ByVal bCaptionDragged As Boolean, ByVal bCombinedMovement As Boolean)
'-----------------------------------------------------------------------------------
' Move the caption and/or field of this CRF element to a new position
' bCaptionDragged is TRUE if element's caption was dragged,
' or FALSE if the field itself was dragged
' bCombinedMovement says whether both the caption and the field should be moved
'-----------------------------------------------------------------------------------
Dim sSQL As String

    If bCaptionDragged And (DataItemId > 0 Or QGroupID > 0) Then
        ' They dragged the caption of a question or a group
        If bCombinedMovement Then
            ' Move the X,Y by the same offset
            ElementX = ElementX + (lNewX - CaptionX)
            ElementY = ElementY + (lNewY - CaptionY)
        End If
        ' Move the caption itself
        CaptionX = lNewX
        CaptionY = lNewY
    Else
        ' They dragged a non-question, or a data entry field
        If bCombinedMovement Then
            ' Update the caption by the same offset
            CaptionX = CaptionX + (lNewX - ElementX)
            CaptionY = CaptionY + (lNewY - ElementY)
        End If
        ' Update the X,Y
        ElementX = lNewX
        ElementY = lNewY
    End If
    
    sSQL = ms_UPDATE_CRFELEMENT _
        & " X = " & ElementX & ", " _
        & " Y = " & ElementY & ", " _
        & " CaptionX = " & CaptionX & ", " _
        & " CaptionY = " & CaptionY _
        & SQLWhereClause

    MacroADODBConnection.Execute sSQL, dbFailOnError

End Sub

'-----------------------------------------------------------------------------------
Public Sub SetControlType(ByVal nControlType As Integer)
'-----------------------------------------------------------------------------------
' Reset the control type of a data entry field
'-----------------------------------------------------------------------------------
Dim sSQL As String

    ' Update the element object
    ControlType = nControlType

    ' Update the DB
    sSQL = ms_UPDATE_CRFELEMENT _
        & "  ControlType = " & nControlType _
        & SQLWhereClause
    
    MacroADODBConnection.Execute sSQL

End Sub

'-----------------------------------------------------------------------------------
Public Sub SetFontDetails(ByVal sName As String, ByVal nSize As Integer, _
                        ByVal bBold As Boolean, ByVal bItalic As Boolean)
'-----------------------------------------------------------------------------------
' Reset the font details of the element
'-----------------------------------------------------------------------------------
Dim sSQL As String
        
    ' Update the element object
    FontName = sName
    FontSize = nSize
    FontBold = bBold
    FontItalic = bItalic
    
    ' Update the DB
    sSQL = ms_UPDATE_CRFELEMENT _
            & " FontName = '" & sName _
            & "', FontSize = " & nSize _
            & ", FontBold = " & Switch(bBold, 1, Not bBold, 0) _
            & ", FontItalic = " & Switch(bItalic, 1, Not bItalic, 0) _
            & SQLWhereClause

    MacroADODBConnection.Execute sSQL

End Sub

'-----------------------------------------------------------------------------------
Public Sub SetCaptionFontDetails(ByVal sName As String, ByVal nSize As Integer, _
                        ByVal bBold As Boolean, ByVal bItalic As Boolean)
'-----------------------------------------------------------------------------------
' Reset the font details of the caption
'-----------------------------------------------------------------------------------
Dim sSQL As String
        
    ' Update the element object
    CaptionFontName = sName
    CaptionFontSize = nSize
    CaptionFontBold = bBold
    CaptionFontItalic = bItalic
    
    ' Update the DB
    sSQL = ms_UPDATE_CRFELEMENT _
            & " CaptionFontName = '" & sName _
            & "', CaptionFontSize = " & nSize _
            & ", CaptionFontBold = " & Switch(bBold, 1, Not bBold, 0) _
            & ", CaptionFontItalic = " & Switch(bItalic, 1, Not bItalic, 0) _
            & SQLWhereClause

    MacroADODBConnection.Execute sSQL

End Sub

'-----------------------------------------------------------------------------------
Public Sub SetFontColour(ByVal lColour As Long)
'-----------------------------------------------------------------------------------
' Reset the font details of the element
'-----------------------------------------------------------------------------------
Dim sSQL As String
        
    ' Update the element object
    FontColour = lColour
    
    ' Update the DB
    sSQL = ms_UPDATE_CRFELEMENT _
            & " FontColour = " & lColour _
            & SQLWhereClause

    MacroADODBConnection.Execute sSQL

End Sub


'-----------------------------------------------------------------------------------
Public Sub SetCaptionFontColour(ByVal lColour As Long)
'-----------------------------------------------------------------------------------
' Reset the font details of the element
'-----------------------------------------------------------------------------------
Dim sSQL As String
        
    ' Update the element object
    CaptionFontColour = lColour
    
    ' Update the DB
    sSQL = ms_UPDATE_CRFELEMENT _
            & " CaptionFontColour = " & lColour _
            & SQLWhereClause

    MacroADODBConnection.Execute sSQL

End Sub

'-----------------------------------------------------------------------------------
Public Sub SetCaption(ByVal sNewCaption As String)
'-----------------------------------------------------------------------------------
' Reset the caption of the element
'-----------------------------------------------------------------------------------
Dim sSQL As String
Dim rsCRFElement As ADODB.Recordset

    ' Set the property
    Me.Caption = sNewCaption
    
    'JL 07/09/98. Recordset edit method replaces Update table method to
    'prevent SQLServer crashing when saving a long text string to the db.
    sSQL = "SELECT Caption FROM CRFElement " & SQLWhereClause
    
    Set rsCRFElement = New ADODB.Recordset
    rsCRFElement.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Len(sNewCaption) > 0 Then
        rsCRFElement!Caption = sNewCaption
    Else
        rsCRFElement!Caption = Null
    End If
    
    rsCRFElement.Update
    
    rsCRFElement.Close
    Set rsCRFElement = Nothing

End Sub

'-----------------------------------------------------------------------------------
Public Sub SetHotlink(ByVal sCaption As String, ByVal sNewHotlink As String)
'-----------------------------------------------------------------------------------
' Reset the caption and link of a hotlink field
'-----------------------------------------------------------------------------------
Dim sSQL As String

    ' Update the element object properties
    Me.Caption = sCaption
    Me.Hotlink = sNewHotlink

    ' Update the DB
    sSQL = ms_UPDATE_CRFELEMENT _
            & " Caption = '" & ReplaceQuotes(sCaption) & "'"
    sSQL = sSQL & ", Hotlink = '" & sNewHotlink & "' "
    sSQL = sSQL & SQLWhereClause
    
    MacroADODBConnection.Execute sSQL

End Sub

'-----------------------------------------------------------------------------------
Public Property Get Expandable() As Boolean
'-----------------------------------------------------------------------------------
' NCJ 1 Jul 03 - TRUE if it's a text box with DisplayLength < QuestionLength
' ic 14/06/2005 added clinical coding
'-----------------------------------------------------------------------------------

    Expandable = (Me.ControlType = gn_TEXT_BOX) _
                And (Me.DisplayLength > 0) _
                And (Me.DisplayLength < Me.DataItemLength) _
                And (Me.DataItemType <> eDataType.Thesaurus)

End Property

