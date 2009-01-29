VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataList 
   Caption         =   "Question List"
   ClientHeight    =   5700
   ClientLeft      =   1545
   ClientTop       =   1650
   ClientWidth     =   2655
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5700
   ScaleWidth      =   2655
   Begin VB.Timer tmrTreeViewClick 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2040
      Top             =   5160
   End
   Begin MSComctlLib.TreeView trvDataList 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   9975
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   18
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgDataList 
      Left            =   240
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "frmDataList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998 - 2006. All Rights Reserved
'   File:       frmDataList.frm
'   Author:     Andrew Newbigging June 1997
'   Purpose:    Displays list of data items within a trial, sorted alphabetically or
'   by CRF page.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'  1 - 21   Andrew Newbigging, Mo Morris        4/06/97 - 29/07/98
'   22  Joanne Lau               02/10/98   SPR 506
'                                           Removed the Loop and replaced it with a call to
'                                           RefreshDataList in the routine DeleteDataItem.
'                                           See Sub for explanation.
'   23  Andrew Newbigging       24/11/98    SR 616
'   Added validation to check that code starts with alphabetic character (required by Prolog)
'   on InsertDataItem
'   24  Andrew Newbigging       3/12/1998   SR 650
'       Modified ChangeDataItemName so that it doesn't rename data items in the tree
'       based on a partial match of the data item id (eg 1 matches 10, etc)
'       Mo Morris               6/1/99      SR 668
'       InsertDataItem now makes an additional call to gblnNotAReservedWord
'   Andrew Newbigging           29/4/99     SR 830
'       Modified trvDataList_DragDrop to force a refresh of the CRF window (when a form
'       is copied) and select the copied form in the CRF window
'   Andrew Newbigging           29/4/99     SR 840
'       Right click on tree view removed
'   41  Paul Norris             29/7/99     SR 648
'       Amended InsertDataItem() for MTM v2.0
'   42 NCJ  12/8/99
'       Updated to use CLM calls for storing data defns.
'   43  Paul Norris             16/08/99     SR 236, 1712, 1501, 1540
'   PN  17/09/99    Amended trvDataList_MouseDown() to prevent menu popup display in library
'   PN  20/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   PN  28/09/99    Amended trvDataList_Click() to set system label if data definition
'                   is not visible and InsertDataItem()
'   NCJ 4 Oct 99    When cancelling drag, reset CRFElementID in frmCRFDesign
'   NCJ 18 Oct 99   CanBeDropped to determine whether data item can be dropped onto CRF page
'                   Changed name of SelectedItemID to SelectedItemPageID
'   NCJ 9 Nov 99    Delete data item from Arezzo in Library Mode (SR 2060)
'  WillC 10/11/99   Added the Error handlers
'   Mo Morris   12/11/99    DAO to ADO conversion
'   WillC 12/11/99  changed data item to question on msgboxes input boxes etc
'   ATN     1/12/99 Added function SelectedCRFElementId
'   NCJ 10/12/99    Added user rights check to menus & clicks etc.
'   NCJ 13/12/99    Ids to Long
'   NCJ 16/12/99    More checks on user's authorisation levels
'   TA 29/04/2000   subclassing removed
'   ZA 03/08/01     SR2731 - Display eForm code when question code is being displayed
'   ASH 21/11/2001  DoUnusedQuestionsExist routine added to fix current buglist no.23
'   ASH 17/12/2001  Cosmetic changes to DoUnUsedQuestionsExist
'   ASH 19/12/2001  Cosmetic changes to DoUnUsedQuestionsExist:added versionid as
'                   additional parameter to allow for deletions in library
'   ASH 21/06/2002 - Added EnableUnusedQuestionsMenu routine to deal with unused questions
' MACRO 3.0
'   NCJ 29 Nov 01 - Use frmCRFDesign to check existence of question in CanBeDropped
'   NCJ 5 Dec 01 - Use frmCRFDesign.RemoveCRFElement
'   REM 17 Dec 01 - Added new routines to build the Question Groups into the tree view
'   NCJ 3 Jan 02 - Made changes to CanBeDropped (from 2.2)
'           Added CanBeDragged to prevent dragging that's not allowed
'   NCJ 7 Jan 02 - Corrected right mouse popup menu (see ShowPopupMenu)
'   REM 31/01/02 - added check for CRFPage in DeleteItem Routine
'   ZA 09/08/2002 - added MACROOnly column in InsertDataItem routine
' NCJ 30 Jan 03 - In DeleteDataItem, check whether Data Definition form is visible
' MLM 01/04/04: Added "Duplicate eForm" menu item.
' NCJ 11 Aug 03 - Ensure question correctly removed from eForm in DeleteDataItem (Roche bug 1944)
' NCJ 15 Jan 04 - Prevent dragging of QGroups from other studies
' NCJ 3 Mar 04 - Prevent RTE 91 in trvDataList_DblClick by checking SelectedItem
' NCJ 9 Mar 05 - Issue 2540 - Default to "MACROOnly", i.e. do not send to OC, for new questions
' NCJ May/Jun 06 - MUSD - Consider access mode
' NCJ 7-13 Sept 06 - Tidying up refreshments etc. for MUSD
'------------------------------------------------------------------------------------'

Option Explicit
Option Base 0
Option Compare Binary

' PN change 43
' Cache the last selected item in case the previously selected item had
' invalid changes in frmDataDefinition
Private msSelectedItemKey As String

' constants used to tag each node in datalist with node type
Private Const msFORM_NODE = "F"
Private Const msDATAITEM_NODE = "D"
Private Const msQGROUP_NODE = "G"
Private Const msUNUSED_NODE = "U"

Private Const msSEPARATOR As String = "|"

'when true, indicates that an item is being dragged
Private mblnDragInProgress As Boolean

Private msUpdateMode As String  'read, update

' Trial key
Private mlClinicalTrialId As Long
Private mnVersionId As Integer
Private msClinicalTrialName As String

Private mbQGroupsOnCRFPages As Boolean
Private mbQuestionsOnCRFPage As Boolean

'REM 17/07/02 -
Public Enum eIdType
    CRFPageId = 0
    QGroupID = 1
    DataItemId = 2
    CRFelementID = 3
End Enum
'Private Const mnCRFPageIdKeyPosition = 0
'Private Const mnQGroupIdKeyPosition = 1
'Private Const mnDataItemIdKeyPosition = 2
'Private Const mnCRFElementIdKeyPosition = 3


Private Const msNOT_ASSIGNED_LABEL = "Not assigned Questions"
Private Const msNOT_ASSIGNED_TEXT = "Unused questions"
'REM 17/12/01 -  new constants for question groups
Private Const msNOT_ASSIGNED_QGROUP_LABEL = "Not assigned QGroup"
Private Const msNOT_ASSIGNED_QGROUP_TEXT = "Unused question groups"

Private mbRefreshing As Boolean

'---------------------------------------------------------------------
Public Property Get ClinicalTrialId() As Long
'---------------------------------------------------------------------

    ClinicalTrialId = mlClinicalTrialId

End Property

'---------------------------------------------------------------------
Public Property Let ClinicalTrialId(ByVal lClinicalTrialId As Long)
'---------------------------------------------------------------------

    mlClinicalTrialId = lClinicalTrialId

End Property

'---------------------------------------------------------------------
Public Property Get VersionId() As Integer
'---------------------------------------------------------------------

    VersionId = mnVersionId

End Property

'---------------------------------------------------------------------
Public Property Let VersionId(ByVal nVersionId As Integer)
'---------------------------------------------------------------------

    mnVersionId = nVersionId

End Property

'---------------------------------------------------------------------
Public Property Get ClinicalTrialName() As String
'---------------------------------------------------------------------

    ClinicalTrialName = msClinicalTrialName

End Property

'---------------------------------------------------------------------
Public Property Let ClinicalTrialName(ByVal sClinicalTrialName As String)
'---------------------------------------------------------------------

    msClinicalTrialName = sClinicalTrialName

End Property

'---------------------------------------------------------------------
Public Property Get UpdateMode() As String
'---------------------------------------------------------------------

    UpdateMode = msUpdateMode
    
End Property

'---------------------------------------------------------------------
Public Property Let UpdateMode(tmpMode As String)
'---------------------------------------------------------------------
    msUpdateMode = tmpMode

End Property

'---------------------------------------------------------------------
Public Function SelectedCRFFormId() As Long
'---------------------------------------------------------------------
' NCJ 13 Dec 99 - Now returns Long
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

'    SelectedCRFFormId = Val(AfterStr(trvDataList.SelectedItem.Key, Len(gsCRF_PAGE_LABEL)))

    SelectedCRFFormId = GetIdFromSelectedItemKey(CRFPageId)

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDataList.SelectedCRFFormId"
    
End Function

'---------------------------------------------------------------------
Public Function SelectedDataItemId() As Long
'---------------------------------------------------------------------
' Get the Data Item Id of the currently selected item
' Tidied up - NCJ 18 Oct 99
' Now returns Long - NCJ 13 Dec 99
' Deal with questions in unused groups - NCJ 7 Jan 02
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    SelectedDataItemId = GetIdFromSelectedItemKey(DataItemId)

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDataList.SelectedDataItemId"
    
End Function

'---------------------------------------------------------------------
Public Function SelectedNodeTag() As eNodeTag
'---------------------------------------------------------------------
' REM 12/07/02
' returns the Tag for a selected node in the question list tree view.
'---------------------------------------------------------------------
Dim sTag As String

    On Error GoTo ErrHandler

    sTag = trvDataList.SelectedItem.Tag
    
    Select Case sTag
    Case "F"
        SelectedNodeTag = eNodeTag.CRFpageTag
    Case "D"
        SelectedNodeTag = eNodeTag.QuestionTag
    Case "G"
        SelectedNodeTag = eNodeTag.QGroupTag
    Case "U"
        SelectedNodeTag = eNodeTag.UnUsedQuestandRQGTag
    Case Else
        GoTo ErrHandler
    End Select

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDatalist.SelectedNodeTag"
End Function


'---------------------------------------------------------------------
Public Function SelectedQGroupID() As Long
'---------------------------------------------------------------------
' REM 19/12/01
' A function that returns the Question Group ID
' of the currently selected item in the Tree View (assuming it's a group)
'---------------------------------------------------------------------
'Dim sKey As String
'Dim nCharsLeftOfLabel As Integer

    On Error GoTo ErrHandler

    SelectedQGroupID = GetIdFromSelectedItemKey(QGroupID)

'    ' Get the key of the selected item
'    sKey = trvDataList.SelectedItem.Key
'    ' Find the QGroupLabel
'    nCharsLeftOfLabel = InStr(sKey, gsQGROUP_LABEL) - 1
'    ' Find what's to the right of it
'    sKey = Right(sKey, Len(sKey) - nCharsLeftOfLabel)
'    ' Now peel off the group label itself
'    SelectedQGroupID = CLng(Right(sKey, Len(sKey) - Len(gsQGROUP_LABEL)))

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDatalist.SelectedQGroupId"
    
End Function

'---------------------------------------------------------------------
Public Function SelectedCRFElementId() As Integer
'---------------------------------------------------------------------
' Get the CRFElementId of the currently selected item
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'REM 17/07/02 - get id from function
    SelectedCRFElementId = GetIdFromSelectedItemKey(CRFelementID)

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDatalist.SelectedCRFElementId"
End Function

'---------------------------------------------------------------------
Public Sub SelectDataItem(lDataItemId As Long, bShowDefinition As Boolean)
'---------------------------------------------------------------------
'Select the given dataitem assuming its in the unused questions node
'---------------------------------------------------------------------
On Error GoTo ErrHandler

    'trvDataList.Nodes(gsDATA_ITEM_LABEL & " " & lDataItemId).Selected = True
    
    'REM 22/07/02 - changed because of new Node Key
    trvDataList.Nodes(msSEPARATOR & msSEPARATOR & lDataItemId & msSEPARATOR).Selected = True
    If bShowDefinition Then
        Call trvDataList_DblClick
    End If
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDatalist.SelectDataItem"
End Sub

'---------------------------------------------------------------------
Public Function SelectedItemName() As String
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler

    If trvDataList.SelectedItem Is Nothing Then
         SelectedItemName = ""
    
    Else
         SelectedItemName = trvDataList.SelectedItem
    
    End If
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDatalist.SelectedItemName"
End Function

'---------------------------------------------------------------------
Public Function SelectedQGroupItemName() As String
'---------------------------------------------------------------------
'returns the name of the selected Question Group
'---------------------------------------------------------------------
Dim lQGroupId As Long

    lQGroupId = GetIdFromSelectedItemKey(QGroupID)
    
    If lQGroupId > 0 Then
        SelectedQGroupItemName = trvDataList.SelectedItem
    End If
    
End Function


'---------------------------------------------------------------------
Public Function SelectedDataItemName() As String
'---------------------------------------------------------------------
'returns the name of the selected question
'---------------------------------------------------------------------
Dim lDataItemId As Long

    On Error GoTo ErrHandler
    
    lDataItemId = GetIdFromSelectedItemKey(DataItemId)

    If lDataItemId > 0 Then
        SelectedDataItemName = trvDataList.SelectedItem
    End If

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDatalist.SelectedDataItemName"

End Function

'---------------------------------------------------------------------
Public Sub AddDataItemToList(lDataItemId As Long, sDataItemName As String, sDataItemCode As String)
'---------------------------------------------------------------------
' this function adds one item to the tree view question list,
' thus there is no need to refresh the whole list
'---------------------------------------------------------------------
Dim nodX As Node
Dim sDisplayText As String
    
    On Error GoTo ErrHandler
    
    ' display the name OR the code in the datalist window?
    If frmMenu.mnuVDisplayDataByName.Checked Then
        sDisplayText = sDataItemName
    Else
        sDisplayText = sDataItemCode
    End If
    
    'If the Display by CRFPage is checked then do
    If frmMenu.mnuVDataCRFPage.Checked = True Then
        Set nodX = trvDataList.Nodes.Add(msNOT_ASSIGNED_LABEL, tvwChild, _
            GetDataItemNodeKey(lDataItemId), _
            sDisplayText, gsDATA_ITEM_LABEL)
    
    Else ' else display by Question
        Set nodX = trvDataList.Nodes.Add(, tvwLast, _
            GetDataItemNodeKey(lDataItemId), _
            sDisplayText, gsDATA_ITEM_LABEL)
    
    End If

    nodX.Tag = msDATAITEM_NODE
    'ensure that the new node is selected
    nodX.Selected = True
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDatalist.AddDataItemToList"
End Sub

'---------------------------------------------------------------------
Public Sub InsertDataItem()
'---------------------------------------------------------------------
' NCJ 16/12/99 - Check user's authorisation level
'---------------------------------------------------------------------
Dim sDataItemCode As String
Dim lDataItemId As Long
Dim nodX As Node
Dim sSQL As String
Dim sUniqueExportName As String
Dim oForm As Form

    On Error GoTo ErrHandler
    
    If UpdateMode = gsREAD Or (Not goUser.CheckPermission(gsFnCreateQuestion)) Then
        Call DialogInformation("You cannot create a new question in this study definition or library")
    Else
    
        ' first check if an item can be inserted
        If frmDataDefinition.SaveChangedData(True) Then
            ' Get valid code from user
            'TA 28/03/2000 - call new function to get code
            sDataItemCode = GetItemCode(gsITEM_TYPE_QUESTION, "New " & gsITEM_TYPE_QUESTION & " code:")
            If sDataItemCode = "" Then    ' if cancel, then return control to user
                Exit Sub
            
            End If
        
            TransBegin
        
            ' Use Arezzo data item ID - NCJ 12/8/99
            lDataItemId = gnNewCLMDataItem(sDataItemCode)
            sUniqueExportName = GenerateNextUniqueExportName(sDataItemCode, ClinicalTrialId, VersionId)
        
            ' NCJ 9 Mar 05 - Issue 2540 - Default to "MACROOnly", i.e. do not send to OC
            sSQL = "INSERT INTO DataItem (ClinicalTrialId,VersionId,DataItemId," _
                & "DataItemCode,DataItemName,DataType,DataItemLength,DataItemFormat,UnitOfMeasurement," _
                & "Derivation, DataItemHelpText, " _
                & "CopiedFromClinicalTrialId,CopiedFromVersionId, CopiedFromDataItemId,MACROOnly, ExportName) " _
                & "VALUES (" & ClinicalTrialId & "," & VersionId & "," _
                & lDataItemId & ",'" & sDataItemCode & "','" & sDataItemCode & "'," _
                & DataType.Text & ",100,'','','','',NULL,NULL,NULL," & eMACROOnly.MACROOnly & ",'" & sUniqueExportName & "')"
            
            MacroADODBConnection.Execute sSQL
                        
            'End transaction
            TransCommit
        
            ' NCJ 19 Jun 06 - Mark as changed
            Call frmMenu.MarkStudyAsChanged
            
            ' NCJ 10 Jan 02 - Add node to ALL question lists currently displayed
            For Each oForm In Forms
                If oForm.Name = "frmDataList" Then
                    If oForm.ClinicalTrialId = ClinicalTrialId Then
                        Call oForm.AddDataItemToList(lDataItemId, sDataItemCode, sDataItemCode)
                    End If
                End If
            Next
            
            ' save the selected node key
            Call CacheSelectedItemKey
            
            ' PN change 43
            If frmDataDefinition.ShowDataDefinition(Me.ClinicalTrialId, _
                                                   Me.VersionId, _
                                                   Me.ClinicalTrialName, _
                                                   lDataItemId, _
                                                   frmMenu.StudyAccessMode, True) Then
                'frmDataDefinition.cboDataType.Enabled = True
                
                ' PN 28/09/99 comment this line out because
                ' when inserting an item the item is set to a text item which does not allow
                ' unit of measurement input
                'frmDataDefinition.cboUnitOfMeasurement.Enabled = True
            
            Else
                ' changes made to the data definition are not valid
                ' and the user wants to keep them
                ' so go back to the previous selected item
                tmrTreeViewClick.Enabled = True
                
            End If
        End If
    End If
        
    'ASH 21/06/2002 Bug 2.2.16 no.11
    Call frmMenu.EnableUnusedQuestionsMenu(mlClinicalTrialId, mnVersionId)

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDatalist.InsertDataItem"
    
End Sub

'---------------------------------------------------------------------
Public Function GetIdFromSelectedItemKey(lGetId As eIdType) As Long
'---------------------------------------------------------------------
'REM 17/07/02
'returns specified Id, either CFRPAgeId, QGroupId, DataItemId or CRFElementId depending on which eIdType is specifid
'---------------------------------------------------------------------
Dim sKey As String
Dim vKey As Variant
Dim lId As Long

    On Error GoTo ErrHandler

    'If there are no items selected in the tree-view then return 0
    If trvDataList.SelectedItem Is Nothing Then
        GetIdFromSelectedItemKey = 0
    Else 'get teh id from the key depending on eIdType
        'the key of the node selected
        sKey = trvDataList.SelectedItem.Key
        ''read the id's into an array
        vKey = Split(sKey, msSEPARATOR)
        'get specific Id
        GetIdFromSelectedItemKey = CLng(Val(vKey(lGetId)))
    End If

Exit Function

ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDatalist.GetIdFromSelectedItemKey"
End Function

'---------------------------------------------------------------------
Public Sub DeleteDataItem()
'---------------------------------------------------------------------
' JL 30/09/98. SPR 506
' Removed the Loop and replaced it with a call to RefreshDataList in the routine
' DeleteDataItem. When attempting to delete a dataitem which appears more than once
' on a CRFPage within the Loop, an error msg: 'Control’s collection has been modified.'
' This occurred because the nodes in the Collection were simultaneously changing whilst
' being processed in the For Loop.
' NCJ 9/11/99 - Do delete data items from Arezzo (SR 2060)
' NCJ 16/12/99 - Check user's authorisation level
' NCJ 5/12/01 - Replaced RemoveCRFElement
' REM 31/01/02 - added check for CRFPage
' NCJ 11 Aug 03 - Roche bug 1944 - Changed RemoveCRFElement to DeleteQuestion
'---------------------------------------------------------------------
Dim lDataItemId As Long
Dim rsCRFElement As ADODB.Recordset
Dim oForm As Form
Dim sSQL As String
Dim sMSG As String
Dim oElement As CRFElement

    On Error GoTo ErrHandler
    
    If UpdateMode = gsREAD Or (Not goUser.CheckPermission(gsFnDelQuestion)) Then
        DialogInformation "You cannot delete a question from this study definition or library"
        Exit Sub
    End If
    
    If frmMenu.TrialStatus > 1 Then
        DialogInformation "You cannot delete a question once a study has been opened"
        Exit Sub
    End If
    
    'Get dataItem Id
    lDataItemId = GetIdFromSelectedItemKey(DataItemId)
    
    If lDataItemId > 0 Then
        sMSG = "Are you sure you want to delete question: " _
                    & trvDataList.SelectedItem & " ?"
        If DialogQuestion(sMSG) = vbYes Then
            ' This deletes it from ALL database tables
            ' (including CRFElement and QGroupQuestion)
            Call gdsDeleteDataItem(Me.ClinicalTrialId, _
                              Me.VersionId, _
                              lDataItemId)
            ' NCJ 19 Jun 06 - Mark study as changed
            Call frmMenu.MarkStudyAsChanged
            
            ' Rebuild our collection of Question Group objects
            ' in case any contained the deleted question
            Call frmMenu.RefreshQuestionGroups(lDataItemId)
            
            ' Now see if we need to redraw stuff on the current eForm
            Set oForm = frmMenu.gFindForm(Me.ClinicalTrialId, _
                                           Me.VersionId, _
                                           "frmCRFDesign")
            If Not oForm Is Nothing Then
                'REM 31/01/02 - added check for CRFPage
                If frmCRFDesign.CRFPageId > 0 Then
                    ' Remove from current eForm if it exists
                    Set oElement = frmCRFDesign.CRFElementByDataItemId(lDataItemId)
                    If Not oElement Is Nothing Then
                        If oElement.OwnerQGroupID = 0 Then
                            ' Delete it if a non-group member
                            ' NCJ 11 Aug 03 - This should be Delete not Remove (Roche bug 1944)
'                            Call frmCRFDesign.RemoveCRFElement(oElement.CRFelementID)
                            Call frmCRFDesign.DeleteQuestion(oElement)
                        Else
                            ' Rebuild the eForm group
                            Call frmCRFDesign.RefreshEFormGroup(oElement.OwnerQGroupID)
                        End If
                    End If
                End If
            End If
            frmMenu.ChangeSelectedItem "", ""
                              
            ' NCJ 10 Jan 02 - Must refresh ALL question lists for this study
            Call frmMenu.RefreshQuestionLists(Me.ClinicalTrialId)

            Set oForm = frmMenu.gFindForm(ClinicalTrialId, VersionId, "frmDataDefinition")
            If Not oForm Is Nothing Then
                ' NCJ 30 Jan 03 - Check whether it's actually showing before doing anything
                If oForm.Visible Then
                    If oForm.DataItemId = lDataItemId Then
                        oForm.CloseWindow
                    End If
                End If
            End If
            
            DeleteProformaDataItem lDataItemId

        End If
    End If
    
Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDatalist.DeleteDataItem"
    
End Sub

'---------------------------------------------------------------------
Public Sub RefreshDataList()
'---------------------------------------------------------------------
' Checks whether data list is grouped into CRF pages
' and calls appropriate sub to refresh the list
'---------------------------------------------------------------------
Dim sDisplayQuestion As String
Dim sDisplayGroup As String
Dim bFormAlphabeticOrder As Boolean

    On Error GoTo ErrLabel

    Screen.MousePointer = vbHourglass
    
    ' PN change 43 - add bFormAlphabeticOrder,sDisplayQuestion parameters
    'If GetSetting(App.Title, "Settings", "ViewDataListName") = "-1" Then
    If frmMenu.mnuVDisplayDataByName.Checked Then
        sDisplayQuestion = "DataItemName"
        sDisplayGroup = "QGroupName"
    Else
        sDisplayQuestion = "DataItemCode"
        sDisplayGroup = "QGroupCode"
    End If
    
    bFormAlphabeticOrder = IsFormDisplayOrderAlphabetic
    
    If frmMenu.mnuVDataCRFPage.Checked = True Then
        Call RefreshDataListGroupedByCRFPage(sDisplayQuestion, sDisplayGroup, bFormAlphabeticOrder)
    Else
        Call RefreshAlphabeticDataList(sDisplayQuestion, sDisplayGroup, bFormAlphabeticOrder)
    End If
    
    Screen.MousePointer = vbDefault

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataList.RefreshDataList"
    
End Sub

'---------------------------------------------------------------------
Private Sub RefreshAlphabeticDataList(sDisplayQuestion As String, sDisplayGroup As String, bFormAlphabeticOrder As Boolean)
'---------------------------------------------------------------------
' REM 19/12/01
' Builds tree view of the questions in alphabetical order, with the nodes expanding
' to show the CRFPage the question is on, or the question group and CRFPage
'---------------------------------------------------------------------

    On Error GoTo ErrLabel

    Call LockWindow(trvDataList)
    
    Call BuildAllQuestionNodes(sDisplayQuestion, bFormAlphabeticOrder)
    
    Call AddCRFPageNodes(bFormAlphabeticOrder)
    
    Call AddQGroupNodes(sDisplayGroup, bFormAlphabeticOrder)
    
    Call AddCRFPagesToQGroupNodes(bFormAlphabeticOrder)
    
    Call UnlockWindow

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataList.RefreshAlphabeticDataList"
End Sub

'---------------------------------------------------------------------
Private Sub BuildAllQuestionNodes(sDisplayQuestion As String, bFormAlphabeticOrder As Boolean)
'---------------------------------------------------------------------
' REM 19/12/01
' Builds all the question nodes for the tree view which shows the Questions in Alphabetic order
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim nodX As Node    ' Create a tree.
Dim tmpDataItemId As Long

    On Error GoTo ErrLabel
    
    With trvDataList
        .Nodes.Clear
        .Sorted = bFormAlphabeticOrder
    End With
    
    Set rsTemp = New ADODB.Recordset
    ' Get the list of all the questions in the study
    Set rsTemp = AllQuestionList(Me.ClinicalTrialId, _
                                 Me.VersionId)
    
    ' Adds all the questions in the study to the tree view
    Do While Not rsTemp.EOF
        ' add each data item
        Set nodX = trvDataList.Nodes.Add(, tvwLast, _
            GetDataItemNodeKey(rsTemp!DataItemId), _
            rsTemp.Fields(sDisplayQuestion).Value, gsDATA_ITEM_LABEL)
        tmpDataItemId = rsTemp!DataItemId
        nodX.Tag = msDATAITEM_NODE
        rsTemp.MoveNext
    Loop

    rsTemp.Close
    Set rsTemp = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataList.BuildAllQuestionNodes"
End Sub

'---------------------------------------------------------------------
Private Sub AddCRFPageNodes(bFormAlphabeticOrder As Boolean)
'---------------------------------------------------------------------
' REM 19/12/01
' Builds the CRFPage nodes onto the Question nodes
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim nodX As Node    ' Create a tree.
Dim sPageText As String

    On Error GoTo ErrLabel

    Set rsTemp = New ADODB.Recordset
    ' Get the list of all CRFpages in the study
    Set rsTemp = CRFPagesList(Me.ClinicalTrialId, _
                                 Me.VersionId, bFormAlphabeticOrder)

    'Check if display by Title or Code
    If frmMenu.mnuVDisplayDataByName.Checked Then
        sPageText = "CRFTitle"
    Else
        sPageText = "CRFPageCode"
    End If

    ' add each CRFPage as a sub node to the Question nodes
    Do While Not rsTemp.EOF
        Set nodX = trvDataList.Nodes.Add(GetDataItemNodeKey(rsTemp!DataItemId), _
            tvwChild, _
            GetCRForDataItemNodeKey(rsTemp!CRFPageId, rsTemp!DataItemId, rsTemp!CRFelementID), _
            rsTemp.Fields(sPageText).Value, gsCRF_PAGE_LABEL)
        nodX.Tag = msFORM_NODE
        rsTemp.MoveNext
    Loop
    
    trvDataList.Sorted = True
   
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataList.AddCRFPageNodes"
End Sub

'---------------------------------------------------------------------
Private Sub AddQGroupNodes(sDisplayGroup As String, bFormAlphabeticOrder As Boolean)
'---------------------------------------------------------------------
' REM 19/12/01
' Adds all the Question Group nodes to the relevant Question nodes i.e questions that are in groups
'---------------------------------------------------------------------
Dim rsQGroups As ADODB.Recordset
Dim nodX As Node    ' Create a tree.
Dim tmpDataItemId As Long

    On Error GoTo ErrLabel
    
    Set rsQGroups = New ADODB.Recordset
    ' Gets the list of Question Groups in the study
    Set rsQGroups = QuestionGroupList(Me.ClinicalTrialId, _
                                 Me.VersionId)

    ' adds each Question Group as a sub node to the relevant Question
    Do While Not rsQGroups.EOF
        Set nodX = trvDataList.Nodes.Add(GetDataItemNodeKey(rsQGroups!DataItemId), _
            tvwChild, _
            GetGroupQuestionNodeKey(0, rsQGroups!DataItemId, rsQGroups!QGroupID, 0), _
            rsQGroups.Fields(sDisplayGroup).Value, gsQGROUP_LABEL)
            'tmpDataItemId = rsQGroups!DataItemId
            nodX.Tag = msQGROUP_NODE
            rsQGroups.MoveNext
    Loop

    rsQGroups.Close
    Set rsQGroups = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataList.AddQGroupNodes"
End Sub

'---------------------------------------------------------------------
Private Sub AddCRFPagesToQGroupNodes(bFormAlphabeticOrder As Boolean)
'---------------------------------------------------------------------
' REM 19/12/01
' Adds all the CRFPages to the Question Group nodes
'---------------------------------------------------------------------
Dim rsCRFPages As ADODB.Recordset
Dim nodX As Node    ' Create a tree.
Dim tmpDataItemId As Long
Dim sPageText As String

    On Error GoTo ErrLabel
    
    Set rsCRFPages = New ADODB.Recordset
    ' Get the list of CRFPages that have question Groups on them
    Set rsCRFPages = CRFPageQGroupList(Me.ClinicalTrialId, _
                                 Me.VersionId)

    'Check for display Title or Code
    If frmMenu.mnuVDisplayDataByName.Checked Then
        sPageText = "CRFTitle"
    Else
        sPageText = "CRFPageCode"
    End If

    ' adds each Question Group as a sub node to the relevant Question
    Do While Not rsCRFPages.EOF
        Set nodX = trvDataList.Nodes.Add(GetGroupQuestionNodeKey(0, rsCRFPages!DataItemId, rsCRFPages!QGroupID, 0), _
            tvwChild, _
            GetGroupQuestionNodeKey(rsCRFPages!CRFPageId, rsCRFPages!DataItemId, rsCRFPages!QGroupID, 0), _
            rsCRFPages.Fields(sPageText).Value, gsCRF_PAGE_LABEL)
        nodX.Tag = msFORM_NODE
        rsCRFPages.MoveNext
    Loop

    rsCRFPages.Close
    Set rsCRFPages = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataList.AddCRFPagesToQGroupNodes"
End Sub

'---------------------------------------------------------------------
Private Sub RefreshDataListGroupedByCRFPage(sDisplayQuestion As String, sDisplayGroup As String, _
                                            bFormAlphabeticOrder As Boolean)
'---------------------------------------------------------------------
' REM 18/12/01
' Sub routine to build the tree view ordered by CRFPage in Study Definition.
' Shows questions and group questions by eform and adds
' unused Questions and Question Groups to the end of the tree view.
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    Call LockWindow(trvDataList)

    Call BuildCRFPageNodes(sDisplayQuestion, bFormAlphabeticOrder)
    
    Call BuildQuestionsGroupsAndGroupQuestionNodes(sDisplayQuestion, sDisplayGroup)
    
    Call BuildUnUsedQuestionNodes(sDisplayQuestion)
    
    Call BuildUnusedQuestionGroupNodes(sDisplayQuestion, sDisplayGroup)
    
    Call UnlockWindow
   
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataList.RefreshDataListGroupedByCRFPage"
End Sub

'---------------------------------------------------------------------
Private Sub BuildCRFPageNodes(sDisplayQuestion As String, bFormAlphabeticOrder As Boolean)
'---------------------------------------------------------------------
' REM 18/12/01
' Builds the CRFPage nodes in the tree view, either in the order they appear
' in the study or in Alphabetical order
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim nodX As Node    ' Create a tree.
Dim tmpCRFPageId As Long
Dim sPageText As String

    On Error GoTo ErrLabel
    
    trvDataList.Nodes.Clear
                              
    ' PN change 43 - add bFormAlphabeticOrder parameter
    ' Adds eforms nodes to the tree view, either by title or code
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gdsCRFPageList(Me.ClinicalTrialId, _
                              Me.VersionId, bFormAlphabeticOrder) ' Gets the list of eforms
                              
    'ZA 03/08/01 SR2731- Check if menu "display question by Name" is checked or not
    'if it is not checked load the CRFTitle otherwise use CRFPageCode
    If frmMenu.mnuVDisplayDataByName.Checked Then
        sPageText = "CRFTitle"
    Else
        sPageText = "CRFPageCode"
    End If
                              
    ' PN change 43
    trvDataList.Sorted = bFormAlphabeticOrder
    'loop through recordset adding EForms by Title or Code to the tree view
    While Not rsTemp.EOF
        Set nodX = trvDataList.Nodes.Add(, tvwLast, _
            GetCRFPageNodeKey(rsTemp!CRFPageId), _
            rsTemp.Fields(sPageText).Value, gsCRF_PAGE_LABEL)
        nodX.Tag = msFORM_NODE
        tmpCRFPageId = rsTemp!CRFPageId
        rsTemp.MoveNext
    Wend
    
    rsTemp.Close
    Set rsTemp = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataList.BuildCRFPageNodes"
End Sub

'---------------------------------------------------------------------
Private Sub BuildQuestionsGroupsAndGroupQuestionNodes(ByVal sDisplayQuestion As String, ByVal sDisplayGroup As String)
'---------------------------------------------------------------------
'REM 09/07/02
'Builds the question, group question and question group nodes
'---------------------------------------------------------------------
Dim rsBuildCRFPages As ADODB.Recordset
Dim rsQuestionNames As ADODB.Recordset
Dim rsGroupNames As ADODB.Recordset
Dim nodX As Node

    'Set QGroups and Questions on CRFPages to false
    mbQGroupsOnCRFPages = False
    mbQuestionsOnCRFPage = False
    
    'Return recordset of all questions, question groups and group questions in a study
    Set rsBuildCRFPages = New ADODB.Recordset
    Set rsBuildCRFPages = BuildCRFPages(Me.ClinicalTrialId, _
                          Me.VersionId)
    
    'Return recordset of all question names and codes
    Set rsQuestionNames = New ADODB.Recordset
    Set rsQuestionNames = QuestionNames(Me.ClinicalTrialId, _
                          Me.VersionId)
    
    'return a recordset of all Qgroup names and codes
    Set rsGroupNames = New ADODB.Recordset
    Set rsGroupNames = GroupNames(Me.ClinicalTrialId, _
                          Me.VersionId)
    
    'Loop through all questions, question groups and group questions in a study
    While Not rsBuildCRFPages.EOF
    
        'if QGroupId > 0 then it is a Question Group, add it to appropriate CRFPage
        If rsBuildCRFPages!QGroupID > 0 Then
            mbQGroupsOnCRFPages = True
            'Filter QGroup Name/Code on QGroupId
            rsGroupNames.Filter = "QGroupId = " & rsBuildCRFPages!QGroupID
            
            Set nodX = trvDataList.Nodes.Add(GetCRFPageNodeKey(rsBuildCRFPages!CRFPageId), _
                tvwChild, _
                GetQuestionGroupNodeKey(rsBuildCRFPages!CRFPageId, rsBuildCRFPages!QGroupID), _
                rsGroupNames.Fields(sDisplayGroup).Value, gsQGROUP_LABEL)
                nodX.Tag = msQGROUP_NODE
                rsBuildCRFPages.MoveNext
    
        'If OwnerQgroupId > 0 then it is a group question, add it to appropriate Question group
        ElseIf rsBuildCRFPages!OwnerQGroupID > 0 Then
            'Filter Question Name/Code on DataItemId
            rsQuestionNames.Filter = "DataItemId = " & rsBuildCRFPages!DataItemId
            
            Set nodX = trvDataList.Nodes.Add(GetQuestionGroupNodeKey(rsBuildCRFPages!CRFPageId, rsBuildCRFPages!OwnerQGroupID), _
                tvwChild, _
                GetGroupQuestionNodeKey(rsBuildCRFPages!CRFPageId, rsBuildCRFPages!DataItemId, rsBuildCRFPages!OwnerQGroupID, _
                                        rsBuildCRFPages!CRFelementID), _
                rsQuestionNames.Fields(sDisplayQuestion).Value, gsDATA_ITEM_LABEL)
                nodX.Tag = msDATAITEM_NODE
                rsBuildCRFPages.MoveNext
                
        Else ' Is a normal question, add it to appropriate CRFPage
            mbQuestionsOnCRFPage = True
            'Filter Question Name/Code on DataItemId
            rsQuestionNames.Filter = "DataItemId = " & rsBuildCRFPages!DataItemId
            
            Set nodX = trvDataList.Nodes.Add(GetCRFPageNodeKey(rsBuildCRFPages!CRFPageId), _
                tvwChild, _
                GetCRForDataItemNodeKey(rsBuildCRFPages!CRFPageId, rsBuildCRFPages!DataItemId, rsBuildCRFPages!CRFelementID), _
                rsQuestionNames.Fields(sDisplayQuestion).Value, gsDATA_ITEM_LABEL)
                nodX.Tag = msDATAITEM_NODE
                rsBuildCRFPages.MoveNext
        End If
        
    Wend

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataList.BuildQuestionsGroupsAndGroupQuestionNodes"
End Sub


'---------------------------------------------------------------------
Private Sub BuildUnUsedQuestionNodes(ByVal sDisplayQuestion As String)
'---------------------------------------------------------------------
'REM 09/07/02
'Builds the unused question nodes
'---------------------------------------------------------------------
Dim nodX As Node
Dim rsUnusedQuestions As ADODB.Recordset

    'Adds unassigned questions to the end of the tree view, i.e. questions not on Eforms or in Unused question Groups
    'First check if there are any questions on EForms in the study
    If mbQuestionsOnCRFPage Then

        Set rsUnusedQuestions = New ADODB.Recordset
        ' Get a list of unused questions that also are not in Question Groups
        Set rsUnusedQuestions = UnusedQuestionList(Me.ClinicalTrialId, _
                                        Me.VersionId)

    Else  ' No questions on eforms in study
        Set rsUnusedQuestions = New ADODB.Recordset
        Set rsUnusedQuestions = QuestionList(Me.ClinicalTrialId, _
                               Me.VersionId) ' therefore get the list of all questions
    End If

    'Create "Unused Question" node at he end of the Tree View
    Set nodX = trvDataList.Nodes.Add(, tvwLast, msNOT_ASSIGNED_LABEL, msNOT_ASSIGNED_TEXT, gsDATA_ITEM_LABEL)
        nodX.Tag = msUNUSED_NODE
    
    ' Add all the unused questions to the "Unused Question" node
    While Not rsUnusedQuestions.EOF
        Set nodX = trvDataList.Nodes.Add(msNOT_ASSIGNED_LABEL, tvwChild, _
            GetDataItemNodeKey(rsUnusedQuestions!DataItemId), _
            rsUnusedQuestions.Fields(sDisplayQuestion).Value, gsDATA_ITEM_LABEL)
        nodX.Tag = msDATAITEM_NODE
        rsUnusedQuestions.MoveNext   ' Move to next record.
    Wend

    rsUnusedQuestions.Close
    Set rsUnusedQuestions = Nothing

    ' Ensure the unassigned items (questions) are expanded
    nodX.EnsureVisible
    ' Scroll to top
    Set nodX = trvDataList.Nodes.Item(1)
    nodX.EnsureVisible

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataList.BuildUnUsedQuestionNodes"
End Sub

'---------------------------------------------------------------------
Private Sub BuildUnusedQuestionGroupNodes(sDisplayQuestion As String, sDisplayGroup As String)
'---------------------------------------------------------------------
' REM 18/12/01
' Builds Unused Question Group and the groups question nodes
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim nodX As Node    ' Create a tree.

    On Error GoTo ErrLabel
    
    'Adds unassigned Questions Groups to the end of the tree view
    'First checks to see if there are any Question Groups on EForms
    If mbQGroupsOnCRFPages Then

        Set rsTemp = New ADODB.Recordset
        Set rsTemp = QGroupNotOnEForm(Me.ClinicalTrialId, _
                                        Me.VersionId) ' Get the list of unassigned Question Groups

    Else ' No Question Groups on eforms in the study
        Set rsTemp = New ADODB.Recordset
        Set rsTemp = QGroups(Me.ClinicalTrialId, _
                               Me.VersionId) ' Get the list of all question groups
    End If

    'Create the "Unused Question Group" Node at the end of the tree view
    Set nodX = trvDataList.Nodes.Add(, tvwLast, msNOT_ASSIGNED_QGROUP_LABEL, msNOT_ASSIGNED_QGROUP_TEXT, gsQGROUP_LABEL)
        nodX.Tag = msUNUSED_NODE
        
    ' Loop through the recordset adding the Group Questions nodes
    While Not rsTemp.EOF
        Set nodX = trvDataList.Nodes.Add(msNOT_ASSIGNED_QGROUP_LABEL, tvwChild, _
            GetQuestionGroupNodeKey(0, rsTemp!QGroupID), rsTemp.Fields(sDisplayGroup).Value, gsQGROUP_LABEL)
        nodX.Tag = msQGROUP_NODE
        rsTemp.MoveNext
    Wend

    rsTemp.Close
    Set rsTemp = Nothing

    ' Ensure the unassigned items (questions) are expanded
    nodX.EnsureVisible
    ' Scroll to top
    Set nodX = trvDataList.Nodes.Item(1)
    nodX.EnsureVisible
    
    
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = QGroupQuestionNotOnEForm(ClinicalTrialId, _
                                   VersionId) ' Get the list of group questions not on eforms
            
    'Adds the Groups questions to specific unused "Question Group" node in the tree view
    While Not rsTemp.EOF
        Set nodX = trvDataList.Nodes.Add(GetQuestionGroupNodeKey(0, rsTemp!QGroupID), _
            tvwChild, _
            GetGroupQuestionNodeKey(0, rsTemp!DataItemId, rsTemp!QGroupID, 0), _
            rsTemp.Fields(sDisplayQuestion).Value, gsDATA_ITEM_LABEL)
        nodX.Tag = msDATAITEM_NODE
        rsTemp.MoveNext
    Wend
   
    rsTemp.Close
    Set rsTemp = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataList.BuildUnusedQuestionGroupNodes"
    
End Sub

'---------------------------------------------------------------------
Private Function GetCRForDataItemNodeKey(lCRFPageId As Long, lDataItemId As Long, _
                                        nCRFElementID As Integer) As String
'---------------------------------------------------------------------
' REM 20/12/01
' Function used to get the key for the Tree View
' This key is used when adding a CRFPage node to a Question node in the "List by Question" View and
' when adding a question to a CRFPage in the "List by CRFPage" View
'---------------------------------------------------------------------

'    GetCRForDataItemNodeKey = gsCRF_PAGE_LABEL & Str(lCRFPageId) _
'                        & gsDATA_ITEM_LABEL & Str(lDataItemId) _
'                        & gsCRF_ELEMENT_LABEL & Str(nCRFElementID)

    GetCRForDataItemNodeKey = lCRFPageId & msSEPARATOR & msSEPARATOR & lDataItemId & msSEPARATOR & nCRFElementID

'Debug.Print "Key = " & GetCRForDataItemNodeKey

End Function

'---------------------------------------------------------------------
Private Function GetCRFPageNodeKey(lCRFPageId As Long) As String
'---------------------------------------------------------------------
' REM 20/12/01
' Returns the Key when adding a CRFPages to the tree view in the list by CRFPage view
'---------------------------------------------------------------------

'    GetCRFPageNodeKey = gsCRF_PAGE_LABEL & Str(lCRFPageId)

    GetCRFPageNodeKey = lCRFPageId & msSEPARATOR & msSEPARATOR & msSEPARATOR

'Debug.Print "Key = " & GetCRFPageNodeKey

End Function

'---------------------------------------------------------------------
Private Function GetDataItemNodeKey(lDataItemId As Long) As String
'---------------------------------------------------------------------
' REM 20/12/01
' Returns the key when adding an unused question to the tree view or
' when adding all the question to the tree view in the "List by Questions" tree view
'---------------------------------------------------------------------

'    GetDataItemNodeKey = gsDATA_ITEM_LABEL & Str(lDataItemId)

    GetDataItemNodeKey = msSEPARATOR & msSEPARATOR & lDataItemId & msSEPARATOR
    
'Debug.Print "Key = " & GetDataItemNodeKey

End Function

'---------------------------------------------------------------------
Private Function GetQuestionGroupNodeKey(lCRFPageId As Long, lQGroupId As Long) As String
'---------------------------------------------------------------------
' REM 20/12/01
' Returns the key when adding a question group to a CRFpage in the tree view
' If CRFpageId = 0 then adding an Unused Question Group to the Unused Question Group Node
'---------------------------------------------------------------------

    If lCRFPageId = 0 Then
        'GetQuestionGroupNodeKey = gsQGROUP_LABEL & Str(lQGroupId)
        GetQuestionGroupNodeKey = msSEPARATOR & lQGroupId & msSEPARATOR & msSEPARATOR
    Else
'        GetQuestionGroupNodeKey = gsCRF_PAGE_LABEL & Str(lCRFPageId) _
'                                  & gsQGROUP_LABEL & Str(lQGroupId)
        
        GetQuestionGroupNodeKey = lCRFPageId & msSEPARATOR & lQGroupId & msSEPARATOR & msSEPARATOR
    End If

'Debug.Print "Key = " & GetQuestionGroupNodeKey

End Function

'---------------------------------------------------------------------
Private Function GetGroupQuestionNodeKey(lCRFPageId As Long, lDataItemId As Long, lQGroupId As Long, nCRFElementID As Integer) As String
'---------------------------------------------------------------------
' REM 20/12/01
' Returns the key when adding the Group Questions to the Question Group node in the "List by CRFPage" tree view
' If nCRFElementId = 0 then adding a Question Group to a Question node in the "List by Question" tree view
' If the CRFPageId and CRFElementId = 0 then adding Group Questions to an Unused Question Group Node
' in the "List by CRFPage" tree view
'---------------------------------------------------------------------

    If (lCRFPageId = 0) And (nCRFElementID = 0) Then
    
'        GetGroupQuestionNodeKey = gsDATA_ITEM_LABEL & Str(lDataItemId) _
'                                & gsQGROUP_LABEL & Str(lQGroupId)
                                
        GetGroupQuestionNodeKey = msSEPARATOR & lQGroupId & msSEPARATOR & lDataItemId & msSEPARATOR
                                        
    ElseIf (nCRFElementID = 0) Then
    
'        GetGroupQuestionNodeKey = gsCRF_PAGE_LABEL & Str(lCRFPageId) _
'                                & gsQGROUP_LABEL & Str(lQGroupId) _
'                                & gsDATA_ITEM_LABEL & Str(lDataItemId)
        
        GetGroupQuestionNodeKey = lCRFPageId & msSEPARATOR & lQGroupId & msSEPARATOR & lDataItemId & msSEPARATOR
        
    Else
        
'        GetGroupQuestionNodeKey = gsCRF_PAGE_LABEL & Str(lCRFPageId) _
'                                & gsQGROUP_LABEL & Str(lQGroupId) _
'                                & gsDATA_ITEM_LABEL & Str(lDataItemId) _
'                                & gsCRF_ELEMENT_LABEL & Str(nCRFElementID)

        GetGroupQuestionNodeKey = lCRFPageId & msSEPARATOR & lQGroupId & msSEPARATOR & lDataItemId & msSEPARATOR & nCRFElementID
        
    End If

'Debug.Print "Key = " & GetGroupQuestionNodeKey

End Function

'---------------------------------------------------------------------
Public Sub ChangeDataItemName(ByVal lDataItemId As Long, _
                              ByVal sNewDataItemName As String, _
                              ByVal sNewDataItemCode As String)
'---------------------------------------------------------------------
'Changes DataItem name in the tree-view when it is changed in the data definition
'---------------------------------------------------------------------
Dim sDisplayText As String
Dim oNode As Node
Dim sKey As String
Dim vKey As Variant
Dim lQuestionKey As Long

On Error GoTo ErrHandler
    
    ' display the name OR the code in the datalist window?
    If frmMenu.mnuVDisplayDataByName.Checked Then
        sDisplayText = sNewDataItemName
    Else
        sDisplayText = sNewDataItemCode
    End If

    'Loop through all the nodes in the tree view
    For Each oNode In trvDataList.Nodes
        'only look at question nodes
        If oNode.Tag = msDATAITEM_NODE Then
            'get the nodes key
            sKey = oNode.Key
            'split the key into an array
            vKey = Split(sKey, msSEPARATOR)
            'Return the DataItemId
            lQuestionKey = CLng(Val(vKey(2)))
    
            'reset the text for dataitem in the tree view
            If (lQuestionKey = lDataItemId) Then
                oNode.Text = sDisplayText
            End If

        End If
    Next
        
'    '   SR 650  3/12/1998   ATN
'    '   Modified check to match correctly in all different views of the tree
'        If frmMenu.mnuVDataCRFPage.Checked = True Then
'            If InStr(oNode.Key, gsDATA_ITEM_LABEL & Str$(sDataItemId) & gsCRF_ELEMENT_LABEL) > 0 Then
'                oNode.Text = sDisplayText
'            ElseIf InStr(oNode.Key, 4 & Str$(sDataItemId)) = 1 Then
'                oNode.Text = sDisplayText
'            End If
'        Else
'            If InStr(oNode.Key, gsDATA_ITEM_LABEL & Str$(sDataItemId)) > 0 Then
'                oNode.Text = sDisplayText
'            End If
'        End If

                               
    trvDataList.Refresh

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "ChangeDataItemName")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'---------------------------------------------------------------------
Public Sub AddCRFToDataItem(lDataItemId As Long, lCRFPageId As Long, nCRFElementID As Integer, sCRFTitle As String, sDataItemName As String)
'---------------------------------------------------------------------
'Adds a DataItems or CRFPage Node to the tree view:
'When the tree view is listed by CRFPage, the routine adds a DataItem node to CRFPage Node
'When the tree view is listed by Question, the routine adds the CRFpage node that the questions what dropped on,
'onto the question node.
'---------------------------------------------------------------------
Dim nodX As Node

    On Error GoTo ErrLabel

    'If Tree View listed by CRFPage then
    If frmMenu.mnuVDataCRFPage.Checked = True Then
        
        On Error Resume Next 'Turn error handling off when setting the node, in case it doesn't exist,
                             'this occurs if dragging a question form another study
        
        Set nodX = trvDataList.Nodes.Item(GetDataItemNodeKey(lDataItemId))
        
        On Error GoTo ErrLabel
        
        'If node exists in the current tree views unused question list then remove it
        If Not (nodX Is Nothing) Then
            trvDataList.Nodes.Remove nodX.Index
        End If
        
        'create new question node under the appropriate CRFPage
        Set nodX = trvDataList.Nodes.Add(GetCRFPageNodeKey(lCRFPageId), tvwChild, _
            GetCRForDataItemNodeKey(lCRFPageId, lDataItemId, nCRFElementID), _
            Str(lDataItemId), gsDATA_ITEM_LABEL)
        nodX.Text = sDataItemName
        nodX.Tag = msDATAITEM_NODE
    
    Else 'If Tree View listed by Question then
    
        On Error Resume Next 'Turn error handling off when setting the node, in case it doesn't exist,
                             'this occurs if dragging a question form another study
        
        Set nodX = trvDataList.Nodes.Item(GetDataItemNodeKey(lDataItemId))
        
        On Error GoTo ErrLabel
        
        'If the node is nothing then create new question node, occurs if dragging a question from another study
        If (nodX Is Nothing) Then
            Set nodX = trvDataList.Nodes.Add(, tvwLast, _
                GetDataItemNodeKey(lDataItemId), Str(lDataItemId), gsDATA_ITEM_LABEL)
            nodX.Text = sDataItemName
            nodX.Tag = msDATAITEM_NODE
        End If
        
        'create a new CRFPage node under the appropriate question node
        Set nodX = trvDataList.Nodes.Add(GetDataItemNodeKey(lDataItemId), tvwChild, _
            GetCRForDataItemNodeKey(lCRFPageId, lDataItemId, nCRFElementID), _
            Str(lCRFPageId), gsCRF_PAGE_LABEL)
        nodX.Text = sCRFTitle
        nodX.Tag = msFORM_NODE
    
    End If
    
    nodX.Selected = True
    nodX.EnsureVisible
    Call CacheSelectedItemKey

    'ASH 21/06/2002 Bug 2.2.16 no.11
    Call frmMenu.EnableUnusedQuestionsMenu(mlClinicalTrialId, mnVersionId)

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataList.RemoveCRFFromDataItem"
    
End Sub

'---------------------------------------------------------------------
Public Sub RemoveCRFFromDataItem(lDataItemId As Long, lCRFPageId As Long, nCRFElementID As Integer)
'---------------------------------------------------------------------
' REM 20/12/01
' Routine that is called when a question is removed from a CRFPage and checks first to see if the question
' resides on any other CRFPages or within any question groups, if not it adds the question to the
' unused question node in the tree view
'---------------------------------------------------------------------
Dim nodX As Node
Dim bElement As Boolean
Dim tmpName As String

    On Error GoTo ErrLabel
    
    Set nodX = trvDataList.Nodes.Item(GetCRForDataItemNodeKey(lCRFPageId, lDataItemId, nCRFElementID))
    
    If Not (nodX Is Nothing) Then
        tmpName = nodX.Text
        Me.trvDataList.Nodes.Remove nodX.Index
    End If
    
    If frmMenu.mnuVDataCRFPage.Checked = True Then
        'Returns true if a question resides on another CRFpage or Question Group
        bElement = IsQuestionOnCRFPageOrQGroup(frmMenu.ClinicalTrialId, _
                                          frmMenu.VersionId, _
                                          lDataItemId)
                                          
        If bElement = False Then
            'adds the question to the Unused Question node
            Set nodX = Me.trvDataList.Nodes.Add(msNOT_ASSIGNED_LABEL, tvwChild, _
                GetDataItemNodeKey(lDataItemId), tmpName, gsDATA_ITEM_LABEL)
            nodX.Text = tmpName
            ' PN change 43
            nodX.Tag = msDATAITEM_NODE
        End If
    
    End If

    'ASH 21/06/2002 Bug 2.2.16 no.11
    Call frmMenu.EnableUnusedQuestionsMenu(mlClinicalTrialId, mnVersionId)

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataList.RemoveCRFFromDataItem"
    
End Sub

'---------------------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------------------

    Call trvDataList_Click
    
End Sub


'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
Dim tmpForm As Form
Dim nFormCount As Integer
    
    On Error GoTo ErrHandler
    
    Me.Icon = frmMenu.Icon
    
    ' Turn on key preview for form, so that F1 (Help) can be trapped by form
    Me.KeyPreview = True
    
    'Initialize TreeView control
    'TA 09/05/2000 SR3288: frmMenu Images now used
    trvDataList.ImageList = frmMenu.imglistSmallIcons

    mbRefreshing = False
    
    Call RefreshDataList
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Form_Load")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'---------------------------------------------------------------------
        
        Call frmMenu.HideQuestions(Me.UpdateMode, Me.ClinicalTrialId, False)

End Sub

'---------------------------------------------------------------------
Private Sub Form_Resize()
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If Me.Height > 400 And Me.Width > 100 Then
        Me.trvDataList.Height = Me.Height - 400
        Me.trvDataList.Width = Me.Width - 100
    End If

Exit Sub
ErrHandler:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, _
                            "Form_Resize", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub tmrTreeViewClick_Timer()
'---------------------------------------------------------------------
' PN change 43
' use a timer control to allow the click event to complete processing before resetting the
' selected item
' the click event on the treeview sets the selected item in the treeview
' to the user selected item when it fires
'---------------------------------------------------------------------

    tmrTreeViewClick.Enabled = False
    If msSelectedItemKey <> vbNullString Then
        Set trvDataList.SelectedItem = trvDataList.Nodes(msSelectedItemKey)
    End If

    
End Sub

'---------------------------------------------------------------------
Private Sub trvDataList_Click()
'---------------------------------------------------------------------
' Click in data list
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    ' Reset selected item
    Call frmMenu.ChangeSelectedItem("", "")
    
    ' NCJ 20 Jun 06 - MUSD - Refresh if necessary
    If frmMenu.RefreshIsNeeded Then Exit Sub
    
    If Not (trvDataList.SelectedItem Is Nothing) Then

        Select Case SelectedNodeTag
    
        Case eNodeTag.QuestionTag
            ' It's a data item
            ' Show selected item
            If Me.ClinicalTrialId = frmMenu.ClinicalTrialId And UpdateMode <> gsREAD Then
                frmMenu.ChangeSelectedItem gsDATA_ITEM_LABEL, "Question: " & Me.SelectedDataItemName
            End If
            
         Case eNodeTag.QGroupTag
            'REM 22/07/02 - added QGroup selected item
            If Me.ClinicalTrialId = frmMenu.ClinicalTrialId And UpdateMode <> gsREAD Then
                frmMenu.ChangeSelectedItem gsQGROUP_LABEL, "Question Group: " & Me.SelectedQGroupItemName
            End If
         
         End Select
        'End If
    End If

Exit Sub
ErrHandler:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, _
                            "trvDataList_Click", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select
    
End Sub

'---------------------------------------------------------------------
Public Function CanBeDropped() As Boolean
'---------------------------------------------------------------------
' Determine whether currently selected data item can be dropped on CRF Page
' Assume selected item is a data item
' Can't be dragged if it's already used on current CRF page,
' otherwise it's draggable if it hasn't been used before,
' or if it's been used elsewhere and "single use" flag isn't set
' NCJ 18 Oct 99
' NCJ 29 Nov 01 - Use frmCRFDesign to check existence
' NCJ 15 Jun 06 - Can't drop unless form is RW
'---------------------------------------------------------------------
Dim oCRFElement As CRFElement
Dim lSelDataID As Long
Dim nCurCRFPageId As Long
Dim sMSG As String
Dim bCanBeDropped As Boolean
Dim sSelDataItemCode As String

    On Error GoTo ErrHandler
    
    ' NCJ 15 Jun 06 - Disallow dropping for RO forms
    If frmMenu.eFormAccessMode = sdReadOnly Then
        CanBeDropped = False
        Exit Function
    End If
    
    lSelDataID = SelectedDataItemId
    
    ' Default to TRUE
    CanBeDropped = True
    
    ' NCJ 3 Jan 01 - Updated to deal with case
    ' when the data list doesn't belong to this trial...
    ' (Based on fix done by ASH in 2.2)
    If Me.ClinicalTrialId <> frmMenu.ClinicalTrialId Then
        ' Check use of this code on this page
        ' Get the data item code
        sSelDataItemCode = DataItemCodeFromId(Me.ClinicalTrialId, lSelDataID)
        
        ' Ask current form if it exists
        If frmCRFDesign.QuestionCodeExists(sSelDataItemCode) Then
            DialogInformation "This question has already been used on this eForm"
            CanBeDropped = False
            Exit Function
        End If
        
        ' Now check for single use dataitems and see if it's been used already
        If gbSingleUseDataItems Then
            If DataItemCodeUsedInStudy(frmMenu.ClinicalTrialId, sSelDataItemCode) Then
                Call DialogInformation("This question has already been used in this study")
                CanBeDropped = False
            End If
        End If
        
        ' That's all if it's a different trial
        Exit Function
    End If
    
    ' The following code is for a data item being dragged
    ' from the current trial's data list
    
    bCanBeDropped = True
    
    sMSG = "This question has already been used "
    
    ' Is it used on the current page?
    If frmCRFDesign.QuestionIDExists(lSelDataID) Then
        sMSG = sMSG & "on this eForm"
        bCanBeDropped = False
        
'    ElseIf InStr(1, trvDataList.SelectedItem.Key, gsCRF_PAGE_LABEL) = 0 _
'      And trvDataList.SelectedItem.Children = 0 Then
    'REM 22/07/02 - changed to check for new key
    ElseIf GetIdFromSelectedItemKey(CRFPageId) = 0 _
        And trvDataList.SelectedItem.Children = 0 Then
        ' It isn't used anywhere else either
        
    ElseIf gbSingleUseDataItems Then
        ' It's not on the current eForm but it is somewhere else
        ' and they're only allowed to use it once
        bCanBeDropped = False
        
    End If
    
    If Not bCanBeDropped Then
        DialogInformation sMSG
    End If
    
    CanBeDropped = bCanBeDropped

Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "CanBeDropped")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Function

'---------------------------------------------------------------------
Public Sub ShowSelectedDataItem()
'---------------------------------------------------------------------

    If AttemptLoadingDataDefinition Then
        frmMenu.ChangeSelectedItem gsDATA_ITEM_LABEL, "Question: " & Me.SelectedDataItemName
        
    End If

End Sub

'---------------------------------------------------------------------
Private Function AttemptLoadingDataDefinition() As Boolean
'---------------------------------------------------------------------
   
   On Error GoTo ErrHandler

    If frmDataDefinition.ShowDataDefinition(Me.ClinicalTrialId, _
                                            Me.VersionId, _
                                           Me.ClinicalTrialName, _
                                           SelectedDataItemId, _
                                            frmMenu.StudyAccessMode) Then
        Call CacheSelectedItemKey
        AttemptLoadingDataDefinition = True
        
    Else
        ' changes made to the data definition are not valid
        ' and the user wants to keep them
        ' so go back to the previous selected item
        tmrTreeViewClick.Enabled = True
                        
    End If
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "AttemptLoadingDataDefinition")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Function

'---------------------------------------------------------------------
Private Function IsDataDefinitionVisible() As Boolean
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

   If Not frmMenu.gFindForm(Me.ClinicalTrialId, _
                             Me.VersionId, _
                             "frmDataDefinition") Is Nothing Then
                             
        If frmMenu.gFindForm(Me.ClinicalTrialId, _
                             Me.VersionId, _
                             "frmDataDefinition").Visible Then
            IsDataDefinitionVisible = True
            
        End If
        
    End If
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "IsDataDefinitionVisible")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Function

'---------------------------------------------------------------------
Private Sub CacheSelectedItemKey()
'---------------------------------------------------------------------

    If Not trvDataList.SelectedItem Is Nothing Then
        msSelectedItemKey = trvDataList.SelectedItem.Key
    End If

End Sub

'---------------------------------------------------------------------
Private Sub trvDataList_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------
' NCJ 16/12/99 - Authorisation checks included in InsertDataItem and DeleteDataItem
' NCJ 11 May 06 - Check study access mode
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If frmMenu.StudyAccessMode >= sdReadWrite Then
        Select Case KeyCode
        Case vbKeyInsert
            InsertDataItem
            
        Case vbKeyDelete
            DeleteDataItem
            
        End Select
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "trvDataList_KeyUp")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub trvDataList_DblClick()
'---------------------------------------------------------------------
' Double-click on Question Tree
' NCJ 11 May 06 - Check study access mode
'---------------------------------------------------------------------
Dim lQGroupId As Long
Dim oQGroup As QuestionGroup

    On Error GoTo ErrHandler
    
    'Call the click event to make sure it has fired before the double click
    Call trvDataList_Click

    ' NCJ 3 Mar 04 - Do nothing if no selected item (to prevent RTE 91 later)
    If trvDataList.SelectedItem Is Nothing Then Exit Sub
    
    If UpdateMode = gsREAD Then
        '    MsgBox "You cannot edit a question in this study definition or library", _
        '        vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
        Exit Sub
    End If
    
    Select Case SelectedNodeTag
    
    Case eNodeTag.QuestionTag
    
        If goUser.CheckPermission(gsFnAmendQuestion) Then    ' NCJ 10/12/99
        'Else
    '        If Not (Me.trvDataList.SelectedItem Is Nothing) Then
    '            If InStr(Me.trvDataList.SelectedItem.Key, gsDATA_ITEM_LABEL) > gnZERO Then
            If GetIdFromSelectedItemKey(DataItemId) > 0 Then
                'is datadefintion window visible
                If IsDataDefinitionVisible Then
                    If trvDataList.SelectedItem.Key <> msSelectedItemKey Then
                        'user double clicked different dataitem so reload window
                        Call AttemptLoadingDataDefinition
                    Else
                        'bring form to the front
                        frmDataDefinition.ZOrder
                    End If
                        
                Else
                    'window not visible so show it
                    Call AttemptLoadingDataDefinition
                        
                End If
                    
            End If
        End If
        
    Case eNodeTag.CRFpageTag
    
        'nothing yet
        
    Case eNodeTag.QGroupTag ' edit the QGroup
        'Check user permissions
        ' NCJ 11 May 06 - And study access mode
        If goUser.CheckPermission(gsFnMaintainQGroups) And frmMenu.StudyAccessMode >= sdReadWrite Then
            'Check if it is a QGroup
            If GetIdFromSelectedItemKey(QGroupID) > 0 Then
                'get the QGroupId
                lQGroupId = GetIdFromSelectedItemKey(QGroupID)
                
                'Set the QGroup object to the one identified by the QGroupId
                Set oQGroup = frmMenu.QuestionGroups.GroupById(lQGroupId)
                
                ' Display the Group Definition dialog
                Call frmMenu.EditQuestionGroup(oQGroup)
                
                Set oQGroup = Nothing
            
            End If
        
        End If
    
    Case eNodeTag.UnUsedQuestandRQGTag
        'Do nothing if user dbl clicked the UnusedQuestion node or the UnusedQGroup node
    
    End Select
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "trvDataList_DblClick")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub trvDataList_MouseDown(Button As Integer, Shift As Integer, _
                                    X As Single, Y As Single)
'---------------------------------------------------------------------
Dim oNode As Node
'Dim bDisplayPopup As Boolean

    On Error GoTo ErrHandler

    '   ATN 1/5/99
    '   Set the selected node on mouse down
    With trvDataList
        Set oNode = .HitTest(X, Y)
        If Not oNode Is Nothing Then
            Set .SelectedItem = .HitTest(X, Y)
        End If
    End With
    
    ' Only continue for Right Button
    If Button <> vbRightButton Then Exit Sub
    
    ' NCJ 20 Jun 06 - MUSD - Refresh if necessary
    If frmMenu.RefreshIsNeeded Then Exit Sub
    
    ' No popup on blank part of window if RO
    If frmMenu.StudyAccessMode < sdReadWrite And oNode Is Nothing Then Exit Sub

    ' Menu only pops up when right mouse is clciked AND not in library mode
    '   ATN 16/12/99 SR 2379 and if its the window of the trial being edited
    ' (ie, not a window just open for copying)
    ' NCJ 23/5/00 - Check UpdateMode too
    ' NCJ 24 May 06 - Consider study access mode
    If ((ClinicalTrialName = gsLIBRARY_LABEL And frmMenu.IsAppInLibraryMode) Or _
            (ClinicalTrialName <> gsLIBRARY_LABEL _
             And Me.ClinicalTrialId = frmMenu.ClinicalTrialId _
             And Me.UpdateMode = gsUPDATE)) Then
        ' prepare the show popup dialog for this Node
        Call ShowPopupMenu(oNode)
    End If
         
'    If (ClinicalTrialName = gsLIBRARY_LABEL And frmMenu.IsAppInLibraryMode) Or _
'        (ClinicalTrialName <> gsLIBRARY_LABEL _
'         And Me.ClinicalTrialId = frmMenu.ClinicalTrialId _
'         And Me.UpdateMode = gsUPDATE) Then
'        bDisplayPopup = True
'
'    Else
'        bDisplayPopup = False
'
'    End If
'
'    If bDisplayPopup Then
'        ' prepare the show popup dialog for this Node
'        Call ShowPopupMenu(oNode)
'    End If

    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "trvDataList_MouseDown", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub ShowPopupMenu(oNode As Node)
'---------------------------------------------------------------------
' Show the rightmouse popup menu for this tree node
' NB oNode may be Nothing
' NCJ 7 Jan 02 - New routine based on original ShowPopup
' MLM 01/04/03: Added "Duplicate eForm" menu item
' NCJ 11 May 06 - Consider Study Access Mode
'---------------------------------------------------------------------
Dim bCanEditStudy As Boolean

    On Error GoTo ErrLabel

    ' NCJ 11 May 06 - Can they edit the study?
    bCanEditStudy = (frmMenu.StudyAccessMode >= sdReadWrite)
    
    With frmMenu
        ' Disable them all first
        .mnuPDataListDuplicateDataD.Enabled = False
        .mnuPDataListEditDataD.Enabled = False
        .mnuPDataListDeleteDataD.Enabled = False
    
        ' Group definition options
        .mnuPDataListEditEFG.Enabled = False
        .mnuPDataListEditQGroup.Enabled = False
        'REM 01/03/02 - added Dupliacte QGroups
        .mnuPDataListDuplicateQGroup.Enabled = False
        'REM 22/07/02 - added Delete QGroups
        .mnuPDataListDeleteQGroup.Enabled = False
        
        ' form definition options
        .mnuPDataListEditForm.Enabled = False
        .mnuPDataListDeleteForm.Enabled = False
        .mnuPDataListViewForm.Enabled = False
        .mnuPDataListDuplicateForm.Enabled = False
            
        ' study definition options
        .mnuPDataListInsertForm.Enabled = False
        .mnuPDataListInsertdataD.Enabled = False
        .mnuPDatalistInsertQGroup.Enabled = False
        
        ' Now enable the ones that are appropriate
        ' Study definition options (only for RW access)
        If oNode Is Nothing Then
            If bCanEditStudy Then
                If goUser.CheckPermission(gsFnCreateEForm) Then
                    .mnuPDataListInsertForm.Enabled = True
                End If
                If goUser.CheckPermission(gsFnCreateQuestion) Then
                    .mnuPDataListInsertdataD.Enabled = True
                End If
                If goUser.CheckPermission(gsFnMaintainQGroups) Then
                    .mnuPDatalistInsertQGroup.Enabled = True
                End If
            End If
        Else
            ' oNode is not Nothing
            Select Case oNode.Tag
            Case msDATAITEM_NODE
                ' Data definition options
                If goUser.CheckPermission(gsFnCreateQuestion) Then
                    .mnuPDataListDuplicateDataD.Enabled = bCanEditStudy
                End If
                If goUser.CheckPermission(gsFnDelQuestion) Then
                    .mnuPDataListDeleteDataD.Enabled = (frmMenu.StudyAccessMode = sdFullControl)
                End If
                If goUser.CheckPermission(gsFnAmendQuestion) Then
                    .mnuPDataListEditDataD.Enabled = True
                    If bCanEditStudy Then
                        .mnuPDataListEditDataD.Caption = "Edit Question Definition"
                    Else
                        ' They can view but not edit
                        .mnuPDataListEditDataD.Caption = "View Question Definition"
                    End If
                End If
        
            Case msQGROUP_NODE
            ' Group definition options
                If goUser.CheckPermission(gsFnMaintainQGroups) Then
                    .mnuPDataListEditQGroup.Enabled = True
                    If bCanEditStudy Then
                        .mnuPDataListEditQGroup.Caption = "Edit Question Group"
                        'REM 01/03/02 - added the enable for the Duplicate QGroups menu
                        .mnuPDataListDuplicateQGroup.Enabled = True
                        'REM 22/07/02 - added the enabled for Deleting QGroups menu
                        .mnuPDataListDeleteQGroup.Enabled = (frmMenu.StudyAccessMode = sdFullControl)
                    Else
                        .mnuPDataListEditQGroup.Caption = "View Question Group"
                    End If
                End If
                ' Editing an eForm group is like amending a question
                ' Check that it's on an eForm
                If GetIdFromSelectedItemKey(CRFPageId) > 0 And goUser.CheckPermission(gsFnAmendQuestion) Then
                    .mnuPDataListEditEFG.Enabled = True
                    If bCanEditStudy Then
                        .mnuPDataListEditEFG.Caption = "Edit eForm Group"
                    Else
                        .mnuPDataListEditEFG.Caption = "View eForm Group"
                    End If
                End If
                
            Case msFORM_NODE
                ' eForm definition options
                If goUser.CheckPermission(gsFnMaintEForm) Then
                    .mnuPDataListDuplicateForm.Enabled = bCanEditStudy
                    .mnuPDataListEditForm.Enabled = True
                    If bCanEditStudy Then
                        .mnuPDataListEditForm.Caption = "Edit eForm Definition"
                    Else
                        ' They can view but not edit
                        .mnuPDataListEditForm.Caption = "View eForm Definition"
                    End If
                End If
                If goUser.CheckPermission(gsFnDelEForm) Then
                    .mnuPDataListDeleteForm.Enabled = (frmMenu.StudyAccessMode = sdFullControl)
                End If
                .mnuPDataListViewForm.Enabled = True
            End Select
        End If
        
        PopupMenu .mnuPDataList, vbPopupMenuRightButton

    End With
    
Exit Sub
ErrLabel:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, "ShowPopupMenu", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Function CanBeDragged(oItem As Node) As Boolean
'---------------------------------------------------------------------
' Can we drag this item from the treeview?
' NCJ 11 May 06 - Consider study access mode
'---------------------------------------------------------------------
Dim sTag As String

    CanBeDragged = False
    
    ' NCJ 11 May 06
    If frmMenu.StudyAccessMode = sdReadOnly Then Exit Function
    
    sTag = oItem.Tag
    
    'don't allow dragging of the UnusedQuestion or UnusedQGroup node
    If sTag = msUNUSED_NODE Then Exit Function
    
    ' Can't drag forms in current study
    If Me.ClinicalTrialId = frmMenu.ClinicalTrialId And Me.UpdateMode <> gsREAD Then
        If oItem.Tag = msFORM_NODE Then Exit Function
    End If
    
    ' NCJ 15 Jan 04 - Can't drag QGroups from other studies
    If Me.ClinicalTrialId <> frmMenu.ClinicalTrialId Then
        If oItem.Tag = msQGROUP_NODE Then Exit Function
    End If
    
    ' It's OK if we get to here
    CanBeDragged = True
    
End Function

'---------------------------------------------------------------------
Private Sub trvDataList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' User is trying to drag a data item from the data list
' Check their level of authorisation before continuing
' NCJ 11 May 06 - Consider study access mode
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    ' NCJ 11 May 06
    If frmMenu.StudyAccessMode = sdReadOnly Then Exit Sub

    If Button = vbLeftButton Then ' Signal a Drag operation.
    
        ' Check user can drag questions around
        If Not goUser.CheckPermission(gsFnMaintEForm) Then
            Exit Sub
        End If
        ' Check user can copy from Library
        If Me.ClinicalTrialId = 0 _
            And Not goUser.CheckPermission(gsFnCopyQuestionFromLib) Then
                Exit Sub
        End If
        ' Check user can copy from another study
        If Me.ClinicalTrialId > 0 And Me.ClinicalTrialId <> frmMenu.ClinicalTrialId _
            And (Not goUser.CheckPermission(gsFnCopyQuestionFromStudy) Or frmMenu.StudyAccessMode = sdReadOnly) Then
                Exit Sub
        End If
        
    '   ATN 1/5/99
    '   Don't set the drag in progress if the node is the 'unused data item' node
    '   or if its a form in the study being edited, or if no item is seletected
        With trvDataList
            If Not .SelectedItem Is Nothing Then
                ' NCJ 3 Jan 02 - Use new CanBeDragged routine
                If CanBeDragged(.SelectedItem) Then
                    ' the item selected is one that can be dragged
                    mblnDragInProgress = True ' Set the flag to true.
                    ' Set the drag icon with the CreateDragImage method.
                    .DragIcon = .SelectedItem.CreateDragImage
                    .Drag vbBeginDrag       ' Begin Drag operation.
                Else
                    ' Not dragging
                    ' Debug.Print "Cancelling drag for " & .SelectedItem.Key
                    ' Clear out previous selection on frmCRFDesign - NCJ 4/10/99
                    frmCRFDesign.CRFelementID = 0
                    mblnDragInProgress = False
                End If
            End If
        End With
    
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "trvDataList_MouseMove")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub trvDataList_DragDrop(Source As Control, X As Single, Y As Single)
'---------------------------------------------------------------------
' Something's been dropped on us
' NCJ 11 May 06 - Consider Study Access Mode
'---------------------------------------------------------------------
Dim lNewId As Long
Dim oForm As Form
Dim vTab As Variant

    On Error GoTo ErrLabel
    
    ' NCJ 11 May 06
    If frmMenu.StudyAccessMode < sdReadWrite Then Exit Sub

    If TypeOf Source Is TreeView Then
        If Source.Name = "trvDataList" And Source.Parent.hWnd = Me.hWnd Then
            ' We can't drag-drop from ourselves
            Set Me.trvDataList.DropHighlight = Nothing
            mblnDragInProgress = False
    
        ElseIf Source.Name = "trvDataList" Then
            ' Drag is from another tree view
            If UpdateMode = gsREAD Then
                Call DialogInformation("You cannot add anything to this study definition or library")
            Else
                'copying a question
                If Source.SelectedItem.Tag = msDATAITEM_NODE Then
                    ' Make a copy of the dragged data item
                    ' REM 16/01/02 - added optional prameter "True" for dragging a single questions
                    lNewId = CopyDataItem(Source.Parent.ClinicalTrialId, Source.Parent.VersionId, _
                        Source.Parent.SelectedDataItemId, Me.ClinicalTrialId, Me.VersionId, , True)
                'REM 28/02/02 - added copying a dragged question group
                ElseIf Source.SelectedItem.Tag = msQGROUP_NODE Then
                    lNewId = CopyingQGroup(Source.Parent.ClinicalTrialId, Source.Parent.VersionId, _
                            Source.Parent.SelectedQGroupID, Me.ClinicalTrialId, Me.VersionId)
                'copying an eForm
                ElseIf Source.SelectedItem.Tag = msFORM_NODE Then
                    ' Make a copy of the dragged CRF Page
'                    lNewId = CopyCRFPage(Source.Parent.ClinicalTrialId, Source.Parent.VersionId, _
'                        Me.ClinicalTrialId, Me.VersionId, _
'                        Mid(Source.SelectedItem.Key, Len(gsCRF_PAGE_LABEL) + 1))
                    lNewId = CopyCRFPage(Source.Parent.ClinicalTrialId, Source.Parent.VersionId, _
                        Me.ClinicalTrialId, Me.VersionId, _
                        Source.Parent.SelectedCRFFormId)
                    If Not frmMenu.gFindForm(Me.ClinicalTrialId, Me.VersionId, _
                        "frmCRFDesign") Is Nothing _
                        And lNewId > 0 Then
                        
                    ' ATN 29/4/99 SR 830
                    ' After copying the form, force a refresh of the form window,
                    ' and select the copied form
                        Set oForm = frmMenu.gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmCRFDesign")
                        oForm.RefreshCRF
                        ' Cycle through each tab until the copied form is found
                        For Each vTab In oForm.tabCRF.Tabs
                            If InStr(vTab.Key, lNewId) > 0 Then
                                Set oForm.tabCRF.SelectedItem = vTab
                            End If
                        Next
                    End If
                End If
                ' NCJ 10 Jan 02 - Refresh ALL data lists for this study
                ' Me.RefreshDataList
                Call frmMenu.RefreshQuestionLists(Me.ClinicalTrialId)
                ' NCJ 11 Sept 06 - Mark study as changed
                Call frmMenu.MarkStudyAsChanged
            End If
        End If
    End If
    
Exit Sub
ErrLabel:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, _
                            "trvDataList_DragDrop", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    If KeyCode = vbKeyF1 Then               ' Show user guide
        'REM 07/12/01 - New Call for the MACRO Help
        Call MACROHelp(Me.hWnd, App.Title)
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, _
                            "Form_KeyDown", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select
    
End Sub

'--------------------------------------------------------------------------------------------------------
Public Function DoUnusedQuestionsExist(ByVal lClinicalTrialId As Long, nVersionId As Integer) As Boolean
'--------------------------------------------------------------------------------------------------------
'ASH 21/11/2001  DoUnusedQuestionsExist routine added to fix current buglist 2.2 no.23
'checks to see if unused questions exist
'Revsion:
' REM 05/07/02 - Changed to take RQG's into account when checking for unused questions
'--------------------------------------------------------------------------------------------------------
Dim rsUnusedQuestionList As ADODB.Recordset

    On Error GoTo ErrLabel
           
    Set rsUnusedQuestionList = UnusedQuestionList(lClinicalTrialId, nVersionId)

    If rsUnusedQuestionList.RecordCount <= 0 Then
        DoUnusedQuestionsExist = False
    Else
        DoUnusedQuestionsExist = True
    End If
    
    rsUnusedQuestionList.Close
    Set rsUnusedQuestionList = Nothing

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataList.DoUnusedQGroupsExist"
End Function


'--------------------------------------------------------------------------------------------------------
Public Function DoUnusedQGroupsExist(ByVal lClinicalTrialId As Long, nVersionId As Integer) As Boolean
'--------------------------------------------------------------------------------------------------------
'REM 19/07/02
'Checks to see if there are any unused question groups
'--------------------------------------------------------------------------------------------------------
Dim rsUnusedQGroups As ADODB.Recordset

    On Error GoTo ErrLabel

    Set rsUnusedQGroups = New ADODB.Recordset
    
    Set rsUnusedQGroups = QGroupNotOnEForm(lClinicalTrialId, nVersionId)
        
    If rsUnusedQGroups.RecordCount <= 0 Then
        DoUnusedQGroupsExist = False
    Else
        DoUnusedQGroupsExist = True
    End If
    
    rsUnusedQGroups.Close
    Set rsUnusedQGroups = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataList.DoUnusedQGroupsExist"
End Function



