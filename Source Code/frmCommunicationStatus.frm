VERSION 5.00
Begin VB.Form frmCommunicationStatus 
   BorderStyle     =   0  'None
   Caption         =   "Select Status"
   ClientHeight    =   4050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3825
   ControlBox      =   0   'False
   DrawStyle       =   6  'Inside Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicStatus 
      DrawStyle       =   6  'Inside Solid
      Height          =   4215
      Left            =   60
      ScaleHeight     =   4155
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.ListBox lstStatus 
         Height          =   3660
         Left            =   0
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   0
         Width           =   3675
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "Select All"
         Height          =   345
         Left            =   0
         TabIndex        =   3
         Top             =   3660
         Width           =   1125
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "Clear All"
         Height          =   345
         Left            =   1260
         TabIndex        =   2
         Top             =   3660
         Width           =   1125
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   345
         Left            =   2550
         TabIndex        =   1
         Top             =   3660
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmCommunicationStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   File:       frmCommunicationsStatus.frm
'   Author:     Ashitei Trebi-Ollennu, October 2002
'   Purpose:    Shows the status messages.
'------------------------------------------------------------------------------
Option Explicit
Private mbOKClicked As Boolean
Private mColStatus As Collection
Private mColCurrentStatuses As Collection
Private mbSelectionMade As Boolean

'---------------------------------------------------------------
Public Function Display(lleft As Long, lTop As Long, _
                        ByRef Col As Collection, _
                        ByRef ColOriginal As Collection) As Boolean
'---------------------------------------------------------------
'
'---------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    mbOKClicked = False
    
    Set mColStatus = Col
    Set mColCurrentStatuses = ColOriginal
    
    If Not mbSelectionMade Then
        LoadListBox
        CheckAll
    Else
        LoadPreviousSelection
    End If
    
    Me.Left = lleft
    Me.Top = lTop
    Me.Show vbModal
    Display = mbOKClicked
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsStatus.Display"
End Function

'----------------------------------------------------------------
Private Sub CheckAll()
'----------------------------------------------------------------
'checks all the stutuses in the list box
'----------------------------------------------------------------
Dim n As Integer

    On Error GoTo ErrHandler
    
    For n = 0 To lstStatus.ListCount
        Call ListCtrl_ListSelect(lstStatus, lstStatus.List(n))
    Next

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsStatus.CheckAll"
End Sub

'---------------------------------------------------------------
Private Sub cmdClearAll_Click()
'---------------------------------------------------------------
'
'---------------------------------------------------------------

    UnCheckAll

End Sub

'----------------------------------------------------------------
Private Sub cmdOK_Click()
'----------------------------------------------------------------
'
'----------------------------------------------------------------
Dim sMsg As String

    sMsg = "Please select at least one status before proceeding"
    
    If Selectionmade Then
        mbSelectionMade = True
        LoadPreviousSelection
        mbOKClicked = True
        Unload Me
    Else
        Call DialogInformation(sMsg, "No Selections")
    End If

End Sub

'-----------------------------------------------------------------
Private Sub cmdSelectAll_Click()
'-----------------------------------------------------------------
'selects all statuses
'-----------------------------------------------------------------

    CheckAll

End Sub

'------------------------------------------------------------------
Private Sub UnCheckAll()
'------------------------------------------------------------------
'deselects all selected statuses
'------------------------------------------------------------------
Dim n As Integer

    On Error GoTo ErrHandler
    
    For n = 0 To lstStatus.ListCount - 1
        If lstStatus.Selected(n) = True Then
            lstStatus.Selected(n) = False
        End If
    Next

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsStatus.UnCheckAll"
End Sub

'--------------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------------
'
'--------------------------------------------------------------------

    Me.Icon = frmMenu.Icon
    Me.Left = frmCommunicationsLog.cmdStatus.Left
    FormCentre Me

End Sub

'-------------------------------------------------------------------
Private Sub LoadListBox()
'-------------------------------------------------------------------
'
'-------------------------------------------------------------------
    
    lstStatus.AddItem "Error"
    lstStatus.AddItem "Locked"
    lstStatus.AddItem "Not Received"
    lstStatus.AddItem "Received"
    lstStatus.AddItem "Skipped"
    lstStatus.AddItem "Superseded"
    lstStatus.AddItem "Pending OverRule"

End Sub

'-----------------------------------------------------------------------------------
Private Sub LoadPreviousSelection()
'-----------------------------------------------------------------------------------
'loads the previous selection made by the user
'-----------------------------------------------------------------------------------
Dim i As Integer
Dim j As Integer
Dim sKey As String
Dim retKey As String

    On Error GoTo ErrHandler
    
    lstStatus.Clear
    
    For i = 1 To mColCurrentStatuses.Count
        If mColCurrentStatuses.Item(i) = 2 Then
            sKey = "Error"
        ElseIf mColCurrentStatuses.Item(i) = 3 Then
            sKey = "Locked"
        ElseIf mColCurrentStatuses.Item(i) = 0 Then
            sKey = "Not Received"
        ElseIf mColCurrentStatuses.Item(i) = 1 Then
            sKey = "Received"
        ElseIf mColCurrentStatuses.Item(i) = 4 Then
            sKey = "Skipped"
        ElseIf mColCurrentStatuses.Item(i) = 5 Then
            sKey = "Superseded"
        Else
            sKey = "Pending OverRule"
        End If
        
        lstStatus.AddItem sKey
    Next
       
    For j = 0 To lstStatus.ListCount - 1
        retKey = GetKey(lstStatus.List(j))
        If CollectionMember(mColStatus, retKey, False) Then
            lstStatus.Selected(j) = True
        Else
            lstStatus.Selected(j) = False
        End If
    Next
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsStatus.LoadPreviousSelection"
End Sub

'---------------------------------------------------------------------------------
Private Sub lstStatus_ItemCheck(Item As Integer)
'---------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------
Dim n As Integer

    On Error GoTo ErrHandler
    
    If lstStatus.Selected(Item) = True Then
        If Not CollectionMember(mColStatus, lstStatus.ListIndex, False) Then
            If Item = 0 Then
                n = 2
            ElseIf Item = 1 Then
                n = 3
            ElseIf Item = 2 Then
                n = 0
            ElseIf Item = 3 Then
                n = 1
            ElseIf Item = 4 Then
                n = 4
            ElseIf Item = 5 Then
                n = 5
            Else
                n = 6
            End If
            mColStatus.Add n, CStr(Item)
        End If
    Else
        mColStatus.Remove CStr(Item)
    End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsStatus.lstStatus_ItemCheck"
End Sub

'---------------------------------------------------------------------------------
Private Function GetKey(ByVal sText As String) As String
'---------------------------------------------------------------------------------
'returns the text to be loaded in the listbox
'---------------------------------------------------------------------------------
Dim n As String

    On Error GoTo ErrHandler

    GetKey = ""
    
    If sText = "Error" Then
        n = "0"
    ElseIf sText = "Locked" Then
        n = "1"
    ElseIf sText = "Not Received" Then
        n = "2"
    ElseIf sText = "Received" Then
        n = "3"
    ElseIf sText = "Skipped" Then
        n = "4"
    ElseIf sText = "Superseded" Then
        n = "5"
    Else
        n = "6"
    End If
    
    GetKey = n

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsStatus.GetKey"
End Function

'-------------------------------------------------------------------------
Private Function Selectionmade() As Boolean
'-------------------------------------------------------------------------
'
'-------------------------------------------------------------------------
Dim i As Integer

    On Error GoTo ErrHandler
    
    Selectionmade = False
    
    For i = 0 To lstStatus.ListCount - 1
        If lstStatus.Selected(i) = True Then
            Selectionmade = True
            Exit Function
        End If
    Next

Exit Function:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsStatus.Selectionmade"
End Function

