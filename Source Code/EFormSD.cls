VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EFormSD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       EFormSD.cls
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   Author:     Nicky Johns, November 2001
'   Purpose:    Eform object for eForms in Study Definition in MACRO 3.0.
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'
' REVISIONS:
'   NCJ 28-30 Nov 01 - Initial development, based on code from BuildCRF
'   NCJ 7 Dec 01 - Change to way Groups are created
'   NCJ 13 Dec 01 - Added GetControlTopLeft
'   ZA 19/07/2002 - added functionality for font/colour for caption & comment
'   NCJ 15 Aug 02 - Sorted out null captions in NewCRFElement
'   ZA 19/09/2002 - added display length property in PopulateCRFElement
'   ASH 5/11/2002 - Modified loadeformProperties and added property get EFormWidth
'   NCJ 14 Feb 03 - Ensure correct updates when rebuilding groups and group questions
'----------------------------------------------------------------------------------------'

Option Explicit

Public FontName As String
Public FontSize As Integer
Public FontBold As Boolean
Public FontItalic As Boolean

Private mlTrialID As Long
Private mnVersionId As Integer
Private mlCRFPageId As Long

Private mlEFormWidth As Long ' ASH 5/11/2002

Private mlBackgroundColour As Long
Private mbDisplayNumbers As Boolean

Private mcolElements As Collection
Private moEFGroups As EFormGroupsSD

Private moVBForm As Form

Private moElBuilder As ElementBuilderSD

Private Const mn_CALENDAR_FUDGE = 200

'--------------------------------------------------------------------------'
Public Sub InitAndBuild(oVBForm As Form, _
            lTrialID As Long, nVersionId As Integer, lCRFPageId As Long)
'--------------------------------------------------------------------------'
' Initialise ourselves and build all the elements
'--------------------------------------------------------------------------'

    Set moVBForm = oVBForm
    mlTrialID = lTrialID
    mnVersionId = nVersionId
    mlCRFPageId = lCRFPageId
    
    Set mcolElements = New Collection

    Call LoadEFormProperties
    
    ' Get the eFormGroups
    Set moEFGroups = New EFormGroupsSD
    Call moEFGroups.Load(mlTrialID, mnVersionId, mlCRFPageId)
    
    ' Get the CRFElements
    Call LoadCRFElements
    
    ' Initialise the Element builder class
    Set moElBuilder = New ElementBuilderSD
    Call moElBuilder.Init(oVBForm, Me)
    Call moElBuilder.BuildElements

End Sub

'---------------------------------------------------------------------
Private Sub RebuildEForm()
'---------------------------------------------------------------------
' Rebuild the entire eForm
'---------------------------------------------------------------------

    Set mcolElements = New Collection
    
    ' Get the eFormGroups
    Set moEFGroups = New EFormGroupsSD
    Call moEFGroups.Load(mlTrialID, mnVersionId, mlCRFPageId)
    
    ' Get the CRFElements
    Call LoadCRFElements
    
    Call moElBuilder.BuildElements

End Sub

'---------------------------------------------------------------------
Public Sub ReloadEFormGroup(ByVal lQGroupId As Long)
'---------------------------------------------------------------------
' Reload the eFormGroups collection
' This only happens if this group is used on this eForm
'---------------------------------------------------------------------
Dim oEFGroup As EFormGroupSD
Dim bDoRefresh As Boolean
Dim nElementID As Integer

    ' Decide whether to do the refresh
    bDoRefresh = False
    For Each oEFGroup In moEFGroups
        If oEFGroup.QGroupID = lQGroupId Then
            ' The group is on the eForm
            bDoRefresh = True
            Exit For
        End If
    Next
    
    If bDoRefresh Then
    
        ' The sledgehammer approach!!!
        Call RebuildEForm
        
'        ' Remove the previous one before reloading the group defn.
'        ' Delete CRFElement (but not from DB)
'        Set oElement = CRFElementByGroupId(lQGroupID)
'        nElementId = oElement.CRFElementId
'        Call DeleteCRFElement(oElement, False)
'
'        Set moEFGroups = Nothing
'        Set moEFGroups = New EFormGroupsSD
'        Call moEFGroups.Load(mlTrialId, mnVersionId, mlCRFPageId)
'
'        ' Now reload all the CRFElements
'        Set oElement = CreateNewCRFElement(0, nElementId)
'        For Each nElementId In moEFGroups.EFormGroupById(lQGroupID)
'
'        Next


    End If
   
    Set oEFGroup = Nothing

End Sub

'--------------------------------------------------------------------------'
Public Property Get ElBuilder() As ElementBuilderSD
'--------------------------------------------------------------------------'
' The element builder class
'--------------------------------------------------------------------------'

    Set ElBuilder = moElBuilder
    
End Property

'--------------------------------------------------------------------------'
Public Property Get CRFElements() As Collection
'--------------------------------------------------------------------------'
' The collection of CRFElements on the form
'--------------------------------------------------------------------------'

    Set CRFElements = mcolElements

End Property

'--------------------------------------------------------------------------'
Public Property Get EFormGroups() As EFormGroupsSD
'--------------------------------------------------------------------------'
' The collection of EForm Groups on the form
'--------------------------------------------------------------------------'

    Set EFormGroups = moEFGroups

End Property

'--------------------------------------------------------------------------'
Public Property Get VBForm() As Form
'--------------------------------------------------------------------------'
' The VB form on which this eForm is displayed
' (an instance of frmCRFDesign)
'--------------------------------------------------------------------------'

    Set VBForm = moVBForm

End Property

'--------------------------------------------------------------------------'
Public Property Get CRFPageId() As Long
'--------------------------------------------------------------------------'
' The ID of the eForm we're displaying
'--------------------------------------------------------------------------'

    CRFPageId = mlCRFPageId

End Property

'--------------------------------------------------------------------------'
Public Property Get EFormWidth() As Long
'--------------------------------------------------------------------------'
'ASH 5/11/2002
'--------------------------------------------------------------------------'

    EFormWidth = mlEFormWidth

End Property

'--------------------------------------------------------------------------'
Public Property Get ClinicalTrialId() As Long
'--------------------------------------------------------------------------'
' The ID of the trial of the eForm we're displaying
'--------------------------------------------------------------------------'

    ClinicalTrialId = mlTrialID

End Property

'--------------------------------------------------------------------------'
Public Property Get VersionId() As Long
'--------------------------------------------------------------------------'
' The version of the trial of the eForm we're displaying
'--------------------------------------------------------------------------'

    VersionId = mnVersionId

End Property

'--------------------------------------------------------------------------'
Public Function CRFElementById(nElementID As Integer) As CRFElement
'--------------------------------------------------------------------------'
' Return the CRFElement corresponding to this ID
'--------------------------------------------------------------------------'

    Set CRFElementById = mcolElements.Item(Str(nElementID))
    
End Function

'--------------------------------------------------------------------------'
Public Function CRFElementByGroupId(ByVal lQGroupId As Long) As CRFElement
'--------------------------------------------------------------------------'
' Return the CRFElement corresponding to this Group ID
' Returns Nothing if there isn't one
'--------------------------------------------------------------------------'
Dim oElement As CRFElement

    ' Search through elements for matching QGroupID
    For Each oElement In mcolElements
        If oElement.QGroupID = lQGroupId Then
            Set CRFElementByGroupId = oElement
            Exit For
        End If
    Next
    Set oElement = Nothing
    
End Function

'--------------------------------------------------------------------------'
Public Function CRFElementByDataItemId(ByVal lDataItemId As Long) As CRFElement
'--------------------------------------------------------------------------'
' Return CRFElement corresponding to this DataItemId (if any)
' NB May be a group member
' Returns Nothing if question does not appear on eForm
'--------------------------------------------------------------------------'

Dim oElement As CRFElement

    ' Search through elements for matching DataItemID
    For Each oElement In mcolElements
        If oElement.DataItemId = lDataItemId Then
            Set CRFElementByDataItemId = oElement
            Exit For
        End If
    Next
    Set oElement = Nothing

End Function

'--------------------------------------------------------------------------'
Public Function CRFElementByIndex(nControlIndex As Integer) As CRFElement
'--------------------------------------------------------------------------'
' Return the CRFElement corresponding to this control array index
'--------------------------------------------------------------------------'

    Set CRFElementByIndex = moElBuilder.EFormItemByIndex(nControlIndex).Element
    
End Function

'--------------------------------------------------------------------------'
Public Sub BuildNewCRFElement(oElement As CRFElement)
'--------------------------------------------------------------------------'
' Build an individual CRFElement
'(assume new so the old one doesn't have to be removed)
'--------------------------------------------------------------------------'

    Call moElBuilder.BuildElement(oElement)

End Sub

'--------------------------------------------------------------------------'
Public Sub ReBuildCRFElement(oElement As CRFElement, _
                            Optional bReloadDefn As Boolean = True, _
                            Optional bRemovePrevious As Boolean = True)
'--------------------------------------------------------------------------'
' ReBuild an individual CRFElement
' If bReloadDefn = TRUE, reload element definition from database
' If bRemovePrevious = TRUE, remove previous element from eForm first
' Automatically handles groups & group members
'--------------------------------------------------------------------------'
'Dim oEFQG As EFormGroupSD
'Dim oGElement As CRFElement

    On Error GoTo ErrLabel
    
    ' NCJ 14 Feb 03 - Adopt the (proven successful) sledgehammer approach
    ' for groups or group questions...
    If oElement.OwnerQGroupID > 0 Or oElement.QGroupID > 0 Then
        Call RebuildEForm
'        ' We must rebuild its Owner group element
'        Set oEFQG = moEFGroups.EFormGroupById(oElement.OwnerQGroupID)
'        Set oGElement = EFormItemByIndex(oEFQG.ControlIndex).Element
'        Call ReBuildCRFElement(oGElement, bReloadDefn, bRemovePrevious)
'        Set oEFQG = Nothing
'        Set oGElement = Nothing
    Else
        If bRemovePrevious Then
            ' Remove the old one from the eForm
            Call RemoveCRFElement(oElement)
        End If
        
        If bReloadDefn Then
            ' Reload with values from the database
            Call LoadCRFElement(oElement)
        End If
        
        ' Now rebuild it on the eForm
        Call moElBuilder.BuildElement(oElement)
    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormSD.RebuildCRFElement"

End Sub

'--------------------------------------------------------------------------'
Private Sub LoadCRFElement(oElement As CRFElement)
'--------------------------------------------------------------------------'
' Populate a CRFElement with values from the database
' Requires the DataItemId and CRFElementID properties of the oElement
'--------------------------------------------------------------------------'
Dim rsCRFElement As ADODB.Recordset
    
    Set rsCRFElement = New ADODB.Recordset
    ' Get all the info for this element
    If oElement.DataItemId > 0 Then
        Set rsCRFElement = gdsCRFDataElement(mlTrialID, mnVersionId, mlCRFPageId, _
                        oElement.CRFelementID)
    Else
        Set rsCRFElement = gdsCRFElement(mlTrialID, mnVersionId, mlCRFPageId, _
                        oElement.CRFelementID)
    End If
    
    Call PopulateCRFElement(oElement, rsCRFElement)
    
    rsCRFElement.Close
    Set rsCRFElement = Nothing

End Sub

'--------------------------------------------------------------------------'
Private Sub Class_Initialize()
'--------------------------------------------------------------------------'
' Set things up on creation
'--------------------------------------------------------------------------'

    Set mcolElements = New Collection
    
End Sub

'--------------------------------------------------------------------------'
Private Sub Class_Terminate()
'--------------------------------------------------------------------------'
' Tidy up when we close
'--------------------------------------------------------------------------'

    Call Terminate
    
End Sub

'--------------------------------------------------------------------------'
Public Sub Terminate()
'--------------------------------------------------------------------------'
' Close things down when we're closing
'--------------------------------------------------------------------------'
    
    If Not moElBuilder Is Nothing Then
        Call moElBuilder.Terminate
        Set moElBuilder = Nothing
    End If
    Set mcolElements = Nothing
    Set moVBForm = Nothing
    Set moEFGroups = Nothing

End Sub

'--------------------------------------------------------------------------'
Public Function NewCRFGroup(oEFQGroup As EFormGroupSD, _
                            ByVal lX As Long, ByVal lY As Long, _
                            ByVal lCaptionX As Long, ByVal lCaptionY As Long, _
                            ByVal vCaption As Variant) As CRFElement
'--------------------------------------------------------------------------'
' Create a new CRF group and return the top level CRFElement
' oEFQGroup is a new eFormGroup which doesn't have any CRFElements yet
'--------------------------------------------------------------------------'
Dim oElement As CRFElement
Dim oGElement As CRFElement
Dim nCRFElementID As Integer
Dim oQGroup As QuestionGroup
Dim i As Integer
Dim lDataItemId As Long
Dim lGroupId As Long

    On Error GoTo ErrLabel
    
    lGroupId = oEFQGroup.QGroupID
    ' First we create the Group element
    ' and then create all its member elements
    Set oElement = NewCRFElement(0, lX, lY, lCaptionX, lCaptionY, vCaption, , lGroupId)
    
    Set oQGroup = frmMenu.QuestionGroups.GroupById(lGroupId)
    ' For each question in the group create a new CRF element
    ' Note that the X,Y and CaptionX,CaptionY are irrelevant for Autofit groups
    ' (since they'll be calculated automatically)
    For i = 1 To oQGroup.Questions.Count
        ' Get the question ID
        lDataItemId = oQGroup.Questions(i)
        Set oGElement = NewCRFGroupMemberElement(lDataItemId, lGroupId, _
                            oElement.FieldOrder, i)
'        Set oGElement = NewCRFElement(lDataItemId, 0, 0, 0, 0, 0, _
'                                DataItemNameFromId(mlTrialId, lDataItemId))
'        ' Set its owner group and field order
'        Call oGElement.SetOwnerQGroup(lGroupID, i, oElement.FieldOrder)
        ' Add it to the eFormGroup
        Call oEFQGroup.AddElement(oGElement.CRFelementID)
    Next
    
    ' Now save the EFormGroup with its new elements
    Call oEFQGroup.Save
    
    ' Return the "top level" group CRFelement
    Set NewCRFGroup = oElement
    
    ' Tidy up before we leave
    Set oElement = Nothing
    Set oGElement = Nothing

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormSD.NewCRFGroup"

End Function

'--------------------------------------------------------------------------'
Private Function NewCRFGroupMemberElement(ByVal lDataItemId As Long, ByVal lOwnerGroupID As Long, _
                    ByVal nFieldOrder As Integer, ByVal nGroupFieldOrder As Integer) As CRFElement
'--------------------------------------------------------------------------'
' Create a new group member CRFElement
' which belongs to the given lOwnerGroupId.
' This is called from NewCRFGroup
'--------------------------------------------------------------------------'
Dim nCRFElementID As Integer

    ' Insert a CRFElement into the database
    nCRFElementID = DBInsertCRFGroupMemberElement(mlTrialID, mnVersionId, mlCRFPageId, _
                        lDataItemId, lOwnerGroupID, nFieldOrder, nGroupFieldOrder)
    
    ' Now create a new CRFElement object
    Set NewCRFGroupMemberElement = CreateNewCRFElement(lDataItemId, nCRFElementID)

End Function

'--------------------------------------------------------------------------'
Private Function CreateNewCRFElement(ByVal lDataItemId As Long, _
                                    ByVal nCRFElementID As Integer) As CRFElement
'--------------------------------------------------------------------------'
' Create a new CRF element object for this DataItemId and nCRFElementId
'--------------------------------------------------------------------------'
Dim oElement As CRFElement

    Set oElement = New CRFElement
    oElement.DataItemId = lDataItemId
    oElement.CRFelementID = nCRFElementID
    ' Load it with values from the database
    Call LoadCRFElement(oElement)
    mcolElements.Add oElement, Str(oElement.CRFelementID)
       
    Set CreateNewCRFElement = oElement
    
    Set oElement = Nothing

End Function

'--------------------------------------------------------------------------'
Public Function NewCRFElement(ByVal lDataItemId As Long, _
                            ByVal lX As Long, ByVal lY As Long, _
                            ByVal lCaptionX As Long, ByVal lCaptionY As Long, _
                            ByVal sCaption As String, _
                            Optional nControlType As Integer = 0, _
                            Optional lGroupId As Long = 0) As CRFElement
'--------------------------------------------------------------------------'
' Create a new CRFElement (NOT a group member)
' See NewCRFGroup for creating a Group crf element
' NCJ 15 Aug 02 - Changed Caption from Variant to String
'--------------------------------------------------------------------------'
Dim nCRFElementID As Integer

    ' Insert a CRFElement into the database
    nCRFElementID = gnInsertCRFElement(mlTrialID, mnVersionId, mlCRFPageId, _
                        lDataItemId, lGroupId, _
                        lX, lY, lCaptionX, lCaptionY, sCaption, nControlType)
    
    ' Now create a new CRFElement object
    Set NewCRFElement = CreateNewCRFElement(lDataItemId, nCRFElementID)

End Function

'--------------------------------------------------------------------------'
Private Function NextFieldOrder() As Integer
'--------------------------------------------------------------------------'
' Return the next available field order for this eForm
'--------------------------------------------------------------------------'

    NextFieldOrder = mnNextFieldOrder(mlTrialID, mnVersionId, mlCRFPageId)
    
End Function

'--------------------------------------------------------------------------'
Private Function NextCRFElementID() As Integer
'--------------------------------------------------------------------------'
' Return the next available CRFElement ID for this eForm
'--------------------------------------------------------------------------'

    NextCRFElementID = mnNextCRFElementId(mlTrialID, mnVersionId, mlCRFPageId)
    
End Function

'---------------------------------------------------------------------
Private Sub LoadCRFElements()
'---------------------------------------------------------------------
' Load the CRFElements collection prior to building an eForm
'---------------------------------------------------------------------
Dim rsCRFElement As ADODB.Recordset
Dim oCRFElement As CRFElement

    Set mcolElements = New Collection
    
    Set rsCRFElement = New ADODB.Recordset
    ' Get all the elements on this page
    Set rsCRFElement = gdsCRFPageDataItems(mlTrialID, mnVersionId, mlCRFPageId)
    
    While Not rsCRFElement.EOF
        ' Build and populate the CRFElement object
        Set oCRFElement = New CRFElement
        Call PopulateCRFElement(oCRFElement, rsCRFElement)
        mcolElements.Add oCRFElement, Str(oCRFElement.CRFelementID)
        rsCRFElement.MoveNext   'get next record
    Wend
    
    rsCRFElement.Close
    Set rsCRFElement = Nothing
    Set oCRFElement = Nothing
    
End Sub

'---------------------------------------------------------------------
Private Sub PopulateCRFElement(oCRFElement As CRFElement, rsCRFElement As ADODB.Recordset)
'---------------------------------------------------------------------
' Populate a CRFElement based on the values in the current row
' of the given recordset
' Assume rsCRFElement contains necessary info from the CRFElement and DataItem tables
'---------------------------------------------------------------------
    
    With oCRFElement
        .CRFelementID = rsCRFElement!CRFelementID
        .ControlType = rsCRFElement!ControlType
        .ClinicalTrialId = rsCRFElement!ClinicalTrialId
        .VersionId = rsCRFElement!VersionId
        .CRFPageId = rsCRFElement!CRFPageId
        .Caption = RemoveNull(rsCRFElement!Caption)
        .DataItemId = rsCRFElement!DataItemId
        .FieldOrder = Val(RemoveNull(rsCRFElement!FieldOrder))
        
        .OwnerQGroupID = rsCRFElement!OwnerQGroupID
        .QGroupFieldOrder = Val(RemoveNull(rsCRFElement!QGroupFieldOrder))
        .QGroupID = rsCRFElement!QGroupID
        .CaptionX = Val(RemoveNull(rsCRFElement!CaptionX))
        .CaptionY = Val(RemoveNull(rsCRFElement!CaptionY))
        .ElementX = Val(RemoveNull(rsCRFElement!X))
        .ElementY = Val(RemoveNull(rsCRFElement!Y))
        .Hidden = (rsCRFElement!Hidden = 1)
        .FontName = RemoveNull(rsCRFElement!FontName)
        ' Use defaults if font name not specified
        If .FontName > "" Then
            .FontSize = Val(RemoveNull(rsCRFElement!FontSize))
            .FontBold = (rsCRFElement!FontBold = 1)
            .FontItalic = (rsCRFElement!FontItalic = 1)
        Else
            .FontName = moVBForm.DefaultFontName
            .FontSize = moVBForm.DefaultFontSize
            .FontBold = moVBForm.DefaultFontBold
            .FontItalic = moVBForm.DefaultFontItalic
        End If
        .FontColour = Val(RemoveNull(rsCRFElement!FontColour))
        If .FontColour = 0 Then
            .FontColour = moVBForm.DefaultFontColour
        End If
        
        'ZA 22/07/2002 - added defaults fonts for question caption
        .CaptionFontName = RemoveNull(rsCRFElement!CaptionFontName)
        If .CaptionFontName > "" Then
            .CaptionFontSize = Val(RemoveNull(rsCRFElement!CaptionFontSize))
            .CaptionFontBold = (rsCRFElement!CaptionFontBold = 1)
            .CaptionFontItalic = (rsCRFElement!CaptionFontItalic = 1)
        Else
            .CaptionFontName = moVBForm.DefaultFontName
            .CaptionFontSize = moVBForm.DefaultFontSize
            .CaptionFontBold = moVBForm.DefaultFontBold
            .CaptionFontItalic = moVBForm.DefaultFontItalic
        End If
        .CaptionFontColour = Val(RemoveNull(rsCRFElement!CaptionFontColour))
        If .CaptionFontColour = 0 Then
            .CaptionFontColour = moVBForm.DefaultFontColour
        End If
        
        ' NCJ 9 Oct 02 - Groups can have StatusFlags too
        If .DataItemId > 0 Or .QGroupID > 0 Then
            .ShowStatusFlag = (rsCRFElement!ShowStatusFlag = 1)
        End If
        
        ' Is it a question?
        If .DataItemId > 0 Then
            .DataItemType = rsCRFElement!DataType
            .DataItemCode = rsCRFElement!DataItemCode
            .DataItemName = rsCRFElement!DataItemName
            .Unit = RemoveNull(rsCRFElement!UnitOfMeasurement)
            .DataItemLength = Val(RemoveNull(rsCRFElement!DataItemLength))
            
            'get Displaylength property
            .DisplayLength = Val(RemoveNull(rsCRFElement!DisplayLength))
            
        End If
        
        ' NCJ 8 Nov 02 - Hotlink property
        .Hotlink = RemoveNull(rsCRFElement!Hotlink)
        'TA 16/10/2003: crf element descritpion
        .Description = RemoveNull(rsCRFElement!Description)
    End With

End Sub

'---------------------------------------------------------------------
Public Function OwnerCRFElement(oElement As CRFElement) As CRFElement
'---------------------------------------------------------------------
' Return the "owner" CRFElement if this is a group member
' and return oElement unchanged if not
'---------------------------------------------------------------------

    If oElement.OwnerQGroupID > 0 Then
        Set OwnerCRFElement = CRFElementById(EFormGroups.EFormGroupById(oElement.OwnerQGroupID).ElementID)
    Else
        Set OwnerCRFElement = oElement
    End If
    
End Function

'---------------------------------------------------------------------
Public Sub RemoveCRFElement(oElement As CRFElement)
'---------------------------------------------------------------------
' This removes a top level CRFElement (i.e. not a group member) from the eForm
' but does NOT delete it from the collection or from the database
' Group members are automatically removed when the group is removed
'---------------------------------------------------------------------
Dim oEFGroup As EFormGroupSD

    On Error GoTo ErrLabel
    
    If oElement.QGroupID > 0 Then
        Set oEFGroup = moEFGroups.EFormGroupById(oElement.QGroupID)
    Else
        Set oEFGroup = Nothing
    End If
    ' Remove it from the VB form
    Call moElBuilder.RemoveElement(oElement, 1, oEFGroup)

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormSD.RemoveCRFElement"
  
End Sub

'---------------------------------------------------------------------
Public Sub DeleteCRFElement(oElement As CRFElement, Optional bDeleteFromDB As Boolean = True)
'---------------------------------------------------------------------
' This removes a CRFElement AND deletes it from the collection
' AND deletes it from the database if bDeleteFromDB = TRUE (the default)
' If a group, deletes all its component elements
' If a group member, deletes the owning group and all its component elements
'---------------------------------------------------------------------
Dim oEFGroup As EFormGroupSD
Dim i As Integer
Dim oElementToDelete As CRFElement

    On Error GoTo ErrLabel
    
    Set oElementToDelete = oElement
    If oElement.QGroupID > 0 Then
        Set oEFGroup = moEFGroups.EFormGroupById(oElement.QGroupID)
    ElseIf oElement.OwnerQGroupID > 0 Then
        Set oEFGroup = moEFGroups.EFormGroupById(oElement.OwnerQGroupID)
        ' Change the element to be the owning group
        Set oElementToDelete = CRFElementById(oEFGroup.ElementID)
    Else
        Set oEFGroup = Nothing
    End If
    
    ' First remove it (and any group elements) from the VB form
    Call moElBuilder.RemoveElement(oElementToDelete, 1, oEFGroup)
       
    If bDeleteFromDB Then
        ' Delete it from the database
        Call DeleteCRFElementFromDB(oElementToDelete)
    End If
    
    ' Now remove it (and group elements) from our collection
    If oElementToDelete.QGroupID > 0 Then
        ' Remove all the group elements
        For i = 1 To oEFGroup.Elements.Count
            mcolElements.Remove Str(oEFGroup.Elements(i))
        Next
        If bDeleteFromDB Then
            ' Delete from the EFormGroups collection
            Call EFormGroups.Delete(oElementToDelete.QGroupID)
        End If
    End If
    
    ' Finally remove the element itself
    mcolElements.Remove Str(oElementToDelete.CRFelementID)
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormSD.DeleteCRFElement"

End Sub

'---------------------------------------------------------------------
Private Sub DeleteCRFElementFromDB(oElement As CRFElement)
'---------------------------------------------------------------------
' Delete CRFElement records from DB
' If a group, also delete all group elements
'---------------------------------------------------------------------
Dim sSQL As String
Dim nErr As Integer
Dim sDesc As String

    On Error GoTo ErrLabel
    
    'Begin transaction
    TransBegin
    
    sSQL = "DELETE FROM CRFElement " _
        & " WHERE  ClinicalTrialId = " & mlTrialID _
        & " AND    VersionId     = " & mnVersionId _
        & " AND    CRFPageId     = " & mlCRFPageId _
        & " AND    CRFElementId  = " & oElement.CRFelementID
                        
    MacroADODBConnection.Execute sSQL
    
    ' Delete all associated group elements
    If oElement.QGroupID > 0 Then
        sSQL = "DELETE FROM CRFElement " _
            & " WHERE  ClinicalTrialId  = " & mlTrialID _
            & " AND    VersionId        = " & mnVersionId _
            & " AND    CRFPageId        = " & mlCRFPageId _
            & " AND    OwnerQGroupId    = " & oElement.QGroupID
                            
        MacroADODBConnection.Execute sSQL
    End If
    
    TransCommit
          
Exit Sub
ErrLabel:
    nErr = Err.Number
    sDesc = Err.Description
    'RollBack transaction
    TransRollBack
    
    Err.Raise nErr, , sDesc & "|EFormSD.DeleteCRFElementFromDB"

End Sub

'---------------------------------------------------------------------
Public Sub ClearEForm()
'---------------------------------------------------------------------
'---------------------------------------------------------------------

    Call moElBuilder.UnloadControls
    Set mcolElements = Nothing

End Sub

'---------------------------------------------------------------------
Public Property Get DisplayNumbers() As Boolean
'---------------------------------------------------------------------
' Whether or not to display question numbering
'---------------------------------------------------------------------

    DisplayNumbers = mbDisplayNumbers
    
End Property

'---------------------------------------------------------------------
Public Property Get BackgroundColour() As Long
'---------------------------------------------------------------------
' The eForm's background colour
'---------------------------------------------------------------------

    BackgroundColour = mlBackgroundColour
    
End Property

'---------------------------------------------------------------------
Public Property Let BackgroundColour(lColour As Long)
'---------------------------------------------------------------------
' Set the eForm's background colour
' (this doesn't actually refresh the eForm)
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrLabel
    
    If lColour <> mlBackgroundColour Then
        sSQL = "UPDATE CRFPage SET " _
            & " BackgroundColour = " & lColour _
            & " WHERE ClinicalTrialId = " & mlTrialID _
            & " AND   VersionId = " & mnVersionId _
            & " AND   CRFPageId = " & mlCRFPageId
        
        MacroADODBConnection.Execute sSQL
        mlBackgroundColour = lColour
    End If
    
Exit Property
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormSD.BackgroundColour"

End Property

'---------------------------------------------------------------------
Private Sub LoadEFormProperties()
'---------------------------------------------------------------------
' Get the background colour and DisplayNumbers from DB
'---------------------------------------------------------------------
Dim rsCRFPage As ADODB.Recordset

    Set rsCRFPage = New ADODB.Recordset
    Set rsCRFPage = gdsCRFPage(mlTrialID, mnVersionId, mlCRFPageId)
    If rsCRFPage!BackgroundColour = 0 Then   ' use default
        mlBackgroundColour = moVBForm.DefaultCRFColour
    Else                                ' use specified colour
        mlBackgroundColour = rsCRFPage!BackgroundColour
    End If
    
    mbDisplayNumbers = (rsCRFPage!DisplayNumbers = 1)
    ' Do we still need this...????
    gnDisplayNumbers = rsCRFPage!DisplayNumbers
    
    'get eform width
    If Not IsNull(rsCRFPage!EFormWidth) Then
        mlEFormWidth = rsCRFPage!EFormWidth
    Else
        mlEFormWidth = glPORTRAIT_WIDTH
    End If
    
    rsCRFPage.Close
    Set rsCRFPage = Nothing

End Sub

'---------------------------------------------------------------------
Public Function EFormItemByIndex(nIndex As Integer) As EFormItemSD
'---------------------------------------------------------------------
' Returns the element/row represented by this control array index
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    Set EFormItemByIndex = moElBuilder.EFormItemByIndex(nIndex)

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormSD.EFormItemByIndex(" & nIndex & ")"

End Function

'---------------------------------------------------------------------
Public Function IndexForElementRow(oElement As CRFElement, nRow As Integer) As Integer
'---------------------------------------------------------------------
' Returns control array index used for this element/row
' (simply pass call on to ElementBuilder class)
'---------------------------------------------------------------------

    IndexForElementRow = moElBuilder.IndexForElementRow(oElement, nRow)

End Function

'---------------------------------------------------------------------
Public Function QuestionIDExists(ByVal lDataItemId As Long) As Boolean
'---------------------------------------------------------------------
' Does question exist on this eForm already?
'---------------------------------------------------------------------
Dim oElement As CRFElement

    QuestionIDExists = False
    For Each oElement In mcolElements
        If oElement.DataItemId = lDataItemId Then
            QuestionIDExists = True
            Exit Function
        End If
    Next

End Function

'---------------------------------------------------------------------
Public Function QuestionCodeExists(ByVal sDataItemCode As String) As Boolean
'---------------------------------------------------------------------
' Does question exist on this eForm already?
'---------------------------------------------------------------------
Dim oElement As CRFElement
Dim sLCCode As String

    sLCCode = lCase(sDataItemCode)
    QuestionCodeExists = False
    For Each oElement In mcolElements
        If lCase(oElement.DataItemCode) = sLCCode Then
            QuestionCodeExists = True
            Exit Function
        End If
    Next

End Function

'---------------------------------------------------------------------
Public Sub DisplayCalendar(ByVal nIndex As Integer)
'---------------------------------------------------------------------
' Display calendar for the text control with the given index
'---------------------------------------------------------------------
Dim sglTop As Single
Dim sglLeft As Single

    On Error GoTo ErrLabel

    ' Get the position of the calendar
    ' (taking into account if element belongs to a group)
    Call moElBuilder.GetCalPopupTopLeft(nIndex, sglTop, sglLeft)
    
    With moVBForm.calCRFElement
    
        'changed Mo Morris 6/6/00 SR 3543, replace moVBForm.Width with moVBForm.picCRFPage.width
        'so that the calendar control is positioned correctly (i.e. always in view
        'and never beyond the right margin)
        'The mn_CALENDAR_FUDGE corrects the display under a 800*600 display when
        'the vertical scroll bar occupies some of the form space.
        If moVBForm.picCRFPage.Width - sglLeft - mn_CALENDAR_FUDGE < .Width Then
            .Left = moVBForm.picCRFPage.Width - .Width - mn_CALENDAR_FUDGE
        Else
            .Left = sglLeft
        End If
        
        .Top = sglTop
        .BackColor = moVBForm.picCRFPage.BackColor
        .Visible = True

        .ZOrder
    
    End With

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormSD.DisplayCalendar(" & nIndex & ")"
End Sub

'---------------------------------------------------------------------
Public Sub DisplayPopupList(nIndex As Integer)
'---------------------------------------------------------------------
' Display PopupList for a "top level" category item,
' i.e. not belonging to a group
'---------------------------------------------------------------------
Dim sglTop As Single
Dim sglLeft As Single
Dim oElement As CRFElement

    On Error GoTo ErrLabel
    
    Set oElement = EFormItemByIndex(nIndex).Element
    
    ' Get the coords of where to display the popup list
    ' NCJ 4 Nov 02 - No longer need oElement
    Call moElBuilder.GetCalPopupTopLeft(nIndex, sglTop, sglLeft)
    
    Call CreatePopupList(oElement, sglTop, sglLeft)
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormSD.DisplayPopupList(" & nIndex & ")"
End Sub

'---------------------------------------------------------------------
Private Sub CreatePopupList(oElement As CRFElement, _
                            sglTop As Single, sglLeft As Single)
'---------------------------------------------------------------------
' Craete and display a popup category list for the given Element,
' at the given coordinates (NCJ 7 Nov 01, extracted from original DisplayPopupList)
'Re-Written by Mo Morris 8/2/00, SR2930,2841
'Size the popuplist based on the specific maximum textwidth of a popuplist's
'value codes and value items.
'The popuplist is now allowed to go beyond the width of the Textbox and commandbox
'that it belongs to.
'The overall width might get restricted by the right side border of the form.
'Changes made by Mo Morris 4/1/01
'Instead of using vForm.Textheight when calculating the height of a popuplist's
'listview .ListItems(1).Height is used.
'Test for a popuplist's listview going beyond the bottom of a form added.
'TA 4/9/01: listview col sizes now down using listview library
'DPH 24/10/2001 - Added Category Item Is Active Checking
'   TA  07/07/2005: seeting of Font.Charset = 1 to allow eastern european characters. CBD2591.
'---------------------------------------------------------------------
Dim olistItem As ListItem
Dim oCatItem As clsDDCategory
Dim i As Long

    ' Set font attributes of PopupList to be same as its text box
    With moVBForm.lvwPopupList
        .Font.Name = oElement.FontName
        .Font.Charset = 1
        .Font.SIZE = oElement.FontSize
        .Font.Bold = oElement.FontBold
        .Font.Italic = oElement.FontItalic
        'Clear the list
        .ListItems.Clear
    End With
   
    ' Set font attributes of form so that form.textwidth can be used for different size fonts
    With moVBForm
        .Font.Name = oElement.FontName
        .Font.Charset = 1
        .Font.SIZE = oElement.FontSize
        .Font.Bold = oElement.FontBold
        .Font.Italic = oElement.FontItalic
    End With

    ' Populate the list view
    For Each oCatItem In oElement.Categories
        ' DPH 24/10/2001 - Active Category Item Checking
        If oCatItem.Active Then
            Set olistItem = moVBForm.lvwPopupList.ListItems.Add(, "K" & oCatItem.ValueOrder, _
                                                oCatItem.ValueCode)
            olistItem.SubItems(1) = oCatItem.Value
        End If
    Next

    With moVBForm.lvwPopupList
        .Left = sglLeft
        .Top = sglTop
        'reset width
        .Width = 20
        
        ' Set Height based on ListItems.Count + 1.6 to allow space for the headers
        ' the number of lines is then multipied against the height of the first row
        ' Check that there is at least one category value first
        If .ListItems.Count > 0 Then
            .Height = (.ListItems.Count + 1.6) * .ListItems(1).Height
        Else
            .Height = moVBForm.TextHeight("W") * 2
        End If
            
        ' Check that width of the popup list
        ' does not put it outside the right border of the form
        ' If so reduce the width and increase the height to allow for scroll bars
        If .Left + .Width > moVBForm.picCRFPage.ScaleWidth Then
            .Width = moVBForm.picCRFPage.ScaleWidth - .Left
            'add 170 twips  to height for horizontal scrollbar
            .Height = .Height + 170
        End If
    
        ' Make the popup list enabled and visible
        .Visible = True
        .Enabled = True

        lvw_SetAllColWidths moVBForm.lvwPopupList, LVSCW_AUTOSIZE_USEHEADER

        
        'TA 09/10/2001: moved this code to after the lvw_SetColwidth code to avoid the lisview expanding each time
        ' Set the overall width of the popup list to be the sum of the headers plus
        ' a fudge factor of 60 twips
        ' (because listview header widths is not an exact science)
        .Width = .ColumnHeaders(1).Width + .ColumnHeaders(2).Width + 60
        
        'Mo Morris 4/1/01
        'Check that the height of the popuplist does not go beyond bottom border of the form
        If .Height > moVBForm.picCRFPage.ScaleHeight - .Top Then
            .Height = moVBForm.picCRFPage.ScaleHeight - .Top
            'the above reduction of the listviews height will cause a vertical scrollbar to be displayed
            'if the listviews width is not increased to allow for this a horizontal scrollbar will also appear
            'check to see if there is enough space to increase the width
            If .Left + .Width + 170 < moVBForm.picCRFPage.ScaleWidth Then
                .Width = .Width + 170
            End If
        End If
        
        .ZOrder
        'ZA 25/07/2002 - added colour property for this control
        .ForeColor = oElement.FontColour
        .Refresh

    End With    ' moVBForm.lvwPopupList

End Sub

'-------------------------------------------------------------------
Public Sub GetControlTopLeft(oControl As Control, _
                        ByRef sglTop As Single, ByRef sglLeft As Single)
'-------------------------------------------------------------------
' Get the Top and Left for a control on the CRFPage,
' taking into account the offsets for its container
'-------------------------------------------------------------------

    Call moElBuilder.GetControlTopLeft(oControl, sglTop, sglLeft)
    
End Sub

