VERSION 5.00
Begin VB.Form frmCommunicationSites 
   BorderStyle     =   0  'None
   Caption         =   "Select Site"
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3915
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3915
      Left            =   120
      ScaleHeight     =   3855
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   60
      Width           =   3795
      Begin VB.ListBox lstSites 
         Height          =   3435
         Left            =   0
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   0
         Width           =   3735
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "Select All"
         Height          =   345
         Left            =   0
         TabIndex        =   3
         Top             =   3480
         Width           =   1125
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "Clear All"
         Height          =   345
         Left            =   1305
         TabIndex        =   2
         Top             =   3480
         Width           =   1125
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   345
         Left            =   2610
         TabIndex        =   1
         Top             =   3480
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmCommunicationSites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   File:       frmCommunicationsSite.frm
'   Author:     Ashitei Trebi-Ollennu, October 2002
'   Purpose:    Gets/Shows all sites in the study.
'------------------------------------------------------------------------------

Option Explicit
Private mbOKClicked As Boolean
Private mColSites As Collection
Private mColAllSites As Collection
Private mbSelectionMade As Boolean

'---------------------------------------------------------------
Public Function Display(lleft As Long, lTop As Long, _
                        ByRef Col As Collection, _
                        ByRef ColOriginal As Collection) As Boolean
'---------------------------------------------------------------
'
'---------------------------------------------------------------
    
    mbOKClicked = False
    
    Set mColSites = Col
    Set mColAllSites = ColOriginal
    
    If Not mbSelectionMade Then
        LoadSiteListBox
        CheckAll
    Else
        LoadPreviousSelection
    End If
    
    Me.Left = lleft
    Me.Top = lTop
    Me.Show vbModal
    Display = mbOKClicked
   
End Function

'---------------------------------------------------------------------------------
Private Sub CheckAll()
'---------------------------------------------------------------------------------
'checks all the sites in the list box
'---------------------------------------------------------------------------------
Dim n As Integer

    On Error GoTo ErrHandler
    
    For n = 0 To lstSites.ListCount
        Call ListCtrl_ListSelect(lstSites, lstSites.List(n))
    Next

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsSite.CheckAll"
End Sub

'----------------------------------------------------------------------------------
Private Sub UnCheckAll()
'----------------------------------------------------------------------------------
'deselects all selected sites
'----------------------------------------------------------------------------------
Dim n As Integer

    On Error GoTo ErrHandler
    
    For n = 0 To lstSites.ListCount - 1
        If lstSites.Selected(n) = True Then
            lstSites.Selected(n) = False
        End If
    Next
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsSite.UnCheckAll"
End Sub

'--------------------------------------------------------------------------------
Private Sub cmdClearAll_Click()
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------

    UnCheckAll

End Sub

'-----------------------------------------------------------------------------------
Private Sub cmdOK_Click()
'-----------------------------------------------------------------------------------
'
'-----------------------------------------------------------------------------------
Dim sMsg As String

    sMsg = "Please select at least one site before proceeding"
    
    If Selectionmade Then
        mbSelectionMade = True
        LoadPreviousSelection
        mbOKClicked = True
        Unload Me
    Else
        Call DialogInformation(sMsg, "No Selection")
    End If

End Sub

'----------------------------------------------------------------------------------
Private Sub cmdSelectAll_Click()
'----------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------

    CheckAll

End Sub

'------------------------------------------------------------------------------------
Private Sub Form_Load()
'------------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------------

    Me.Icon = frmMenu.Icon
    FormCentre Me

End Sub

'--------------------------------------------------------------------------------------
Private Sub LoadSiteListBox()
'--------------------------------------------------------------------------------------
'add sites to the listbox
'--------------------------------------------------------------------------------------
Dim rsSites As ADODB.Recordset
Dim sSQL As String
Dim nNum As Integer

    On Error GoTo ErrHandler
    
    For nNum = 1 To mColSites.Count
        lstSites.AddItem mColSites(nNum)
    Next
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsSite.LoadSiteListBox"
End Sub

'------------------------------------------------------------------------------
Private Sub LoadPreviousSelection()
'------------------------------------------------------------------------------
'loads the selection the user made previously
'------------------------------------------------------------------------------
Dim i As Integer
Dim j As Integer
    
    On Error GoTo ErrHandler
    
    lstSites.Clear
    
    For i = 1 To mColAllSites.Count
        lstSites.AddItem mColAllSites.Item(i)
    Next
    
    For j = 0 To lstSites.ListCount - 1
        If CollectionMember(mColSites, lstSites.List(j), False) Then
            lstSites.Selected(j) = True
        End If
    Next
         
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsSite.LoadPreviousSelection"
End Sub

'-------------------------------------------------------------------------------
Private Sub lstSites_ItemCheck(Item As Integer)
'-------------------------------------------------------------------------------
'adds or removes items from collection based on selections made in the listbox
'-------------------------------------------------------------------------------
    
    If lstSites.Selected(Item) = True Then
        If Not CollectionMember(mColSites, lstSites.List(Item), False) Then
            mColSites.Add lstSites.List(Item), lstSites.List(Item)
        End If
    Else
        mColSites.Remove (lstSites.List(Item))
    End If

End Sub

'-------------------------------------------------------------------------
Private Function Selectionmade() As Boolean
'-------------------------------------------------------------------------
'
'-------------------------------------------------------------------------
Dim i As Integer

    On Error GoTo ErrHandler
    
    Selectionmade = False
    'REM 09/04/03 - check to see if there are any sites first
    If lstSites.ListCount <> 0 Then
        For i = 0 To lstSites.ListCount - 1
            If lstSites.Selected(i) = True Then
                Selectionmade = True
                Exit Function
            End If
        Next
    Else
        Selectionmade = True
    End If

Exit Function:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsSite.Selectionmade"
End Function
