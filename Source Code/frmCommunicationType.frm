VERSION 5.00
Begin VB.Form frmCommunicationType 
   BorderStyle     =   0  'None
   Caption         =   "Select Type"
   ClientHeight    =   4155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicType 
      Height          =   3975
      Left            =   60
      ScaleHeight     =   3915
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.ListBox lstType 
         Height          =   3435
         Left            =   -30
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   60
         Width           =   3675
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "Select All"
         Height          =   345
         Left            =   0
         TabIndex        =   3
         Top             =   3540
         Width           =   1125
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "Clear All"
         Height          =   345
         Left            =   1260
         TabIndex        =   2
         Top             =   3540
         Width           =   1125
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   345
         Left            =   2520
         TabIndex        =   1
         Top             =   3540
         Width           =   1124
      End
   End
End
Attribute VB_Name = "frmCommunicationType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   File:       frmCommunicationsType.frm
'   Author:     Ashitei Trebi-Ollennu, October 2002
'   Purpose:    Shows all message types/enums.
'------------------------------------------------------------------------------

Option Explicit
Private mbOKClicked As Boolean
Private mColType As Collection
Private mColTypeOriginal As Collection
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
    
    Set mColTypeOriginal = ColOriginal
    Set mColType = Col
    
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
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsType.Display"
End Function


'---------------------------------------------------------------------------------
Private Sub CheckAll()
'---------------------------------------------------------------------------------
'checks all the sites in the list box
'---------------------------------------------------------------------------------
Dim n As Integer

    On Error GoTo ErrHandler
    
    For n = 0 To lstType.ListCount - 1
        Call ListCtrl_ListSelect(lstType, lstType.List(n))
    Next

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsType.CheckAll"
End Sub

'--------------------------------------------------------------------------------
Private Sub cmdClearAll_Click()
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------

    UnCheckAll

End Sub

'----------------------------------------------------------------------------------
Private Sub cmdOK_Click()
'----------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------
Dim sMsg As String

    sMsg = "Please select as least one type before proceeding"
    
    If Selectionmade Then
        mbSelectionMade = True
        LoadPreviousSelection
        mbOKClicked = True
        Unload Me
    Else
        Call DialogInformation(sMsg, "No Selections")
    End If
        
End Sub

'------------------------------------------------------------------------------------
Private Sub cmdSelectAll_Click()
'------------------------------------------------------------------------------------
'
'-------------------------------------------------------------------------------------

    CheckAll

End Sub

'--------------------------------------------------------------------------------------
Private Sub UnCheckAll()
'--------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------
Dim n As Integer

On Error GoTo ErrHandler
    
    For n = 0 To lstType.ListCount - 1
        If lstType.Selected(n) = True Then
            lstType.Selected(n) = False
        End If
    Next

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsType.UnCheckAll"
End Sub

'-------------------------------------------------------------------------------------
Private Sub Form_Load()
'-------------------------------------------------------------------------------------
'
'-------------------------------------------------------------------------------------

    Me.Icon = frmMenu.Icon
    FormCentre Me

End Sub

'------------------------------------------------------------------------------
Private Sub LoadListBox()
'------------------------------------------------------------------------------
'manual adding to list box. this allows the listindex to be known in advance
'for it to be used as key when the enums are added to the collection
'------------------------------------------------------------------------------
    
    lstType.AddItem "New Trial"
    lstType.AddItem "In Preparation"
    lstType.AddItem "Trial Open"
    lstType.AddItem "Closed Recruitment"
    lstType.AddItem "Closed FollowUp"
    lstType.AddItem "Trial Suspended"
    lstType.AddItem "New Version"
    lstType.AddItem "Patient Data"
    lstType.AddItem "Mail"
    lstType.AddItem "Locking/Freezing"
    lstType.AddItem "Unlocking"
    lstType.AddItem "Lab Definition Server To Site"
    lstType.AddItem "Lab Definition Site To Server"
    lstType.AddItem "User"
    lstType.AddItem "User Role"
    lstType.AddItem "Password Change"
    lstType.AddItem "Role"
    lstType.AddItem "Password Policy"
    lstType.AddItem "SDV"
    lstType.AddItem "Note"
    lstType.AddItem "Discrepancy"
    lstType.AddItem "System Log"
    lstType.AddItem "User Log"
    lstType.AddItem "Restore User Role"

End Sub

'----------------------------------------------------------------------------------
Private Sub lstType_ItemCheck(Item As Integer)
'----------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------
 Dim n As String
 
 On Error GoTo ErrHandler
    
    If lstType.Selected(Item) = True Then
        If Not CollectionMember(mColType, lstType.ListIndex, False) Then
            n = GetEnum(lstType.List(Item))
            mColType.Add n, CStr(Item)
        End If
    Else
        mColType.Remove CStr(Item)
    End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsType.lstType_ItemCheck"
End Sub

'--------------------------------------------------------------------------------------
Private Function GetEnum(sIndex As String) As String
'--------------------------------------------------------------------------------------
'receives the text from the listbox and returns the ENUM
'--------------------------------------------------------------------------------------
Dim sRetText As String

    On Error GoTo ErrHandler
    
    Select Case sIndex
        Case Is = "New Trial": sRetText = "0"
        Case Is = "In Preparation": sRetText = "1"
        Case Is = "Trial Open": sRetText = "2"
        Case Is = "Closed Recruitment": sRetText = "3"
        Case Is = "Closed FollowUp": sRetText = "4"
        Case Is = "Trial Suspended": sRetText = "5"
        Case Is = "New Version": sRetText = "8"
        Case Is = "Patient Data": sRetText = "10"
        Case Is = "Mail": sRetText = "11"
        Case Is = "Locking/Freezing": sRetText = "16,17,18,19"
        Case Is = "Unlocking": sRetText = "20,21,22"
        Case Is = "Lab Definition Server To Site": sRetText = "30"
        Case Is = "Lab Definition Site To Server": sRetText = "31"
        Case Is = "User": sRetText = "32"
        Case Is = "User Role": sRetText = "33"
        Case Is = "Password Change": sRetText = "34"
        Case Is = "Role": sRetText = "35"
        Case Is = "System Log": sRetText = "36"
        Case Is = "User Log": sRetText = "37"
        Case Is = "Restore User Role": sRetText = "38"
        Case Is = "Password Policy": sRetText = "40"
        Case Is = "SDV": sRetText = "18|3"
        Case Is = "Note": sRetText = "19|2"
        Case Is = "Discrepancy": sRetText = "20|0"
    End Select
    
    GetEnum = sRetText
    
Exit Function:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsType.GetEnum"
End Function

'-----------------------------------------------------------------------------------------
Private Sub LoadPreviousSelection()
'-----------------------------------------------------------------------------------------
'reloads the listbox with user selection
'-----------------------------------------------------------------------------------------
Dim i As Integer
Dim j As Integer
Dim sKey As String
Dim retKey As String
Dim n As Integer

    On Error GoTo ErrHandler
    
    lstType.Clear
    
    For i = 1 To mColTypeOriginal.Count
        sKey = GetEnumText(CStr(mColTypeOriginal.Item(i)))
        lstType.AddItem sKey
    Next
       
    For j = 0 To lstType.ListCount - 1
        retKey = GetCollectionKey(lstType.List(j))
        If CollectionMember(mColType, retKey, False) Then
            lstType.Selected(j) = True
        Else
            lstType.Selected(j) = False
        End If
    Next
    
Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsType.LoadPreviousSelection"
End Sub

'--------------------------------------------------------------------------------------
Private Function GetEnumText(nEnum As String) As String
'--------------------------------------------------------------------------------------
'returns the text for a given enum
'--------------------------------------------------------------------------------------
Dim sText As String

    GetEnumText = ""
    
    On Error GoTo ErrHandler

    Select Case nEnum
        Case Is = "0": sText = "New Trial"
        Case Is = "1": sText = "In Preparation"
        Case Is = "2": sText = "Trial Open"
        Case Is = "3": sText = "Closed Recruitment"
        Case Is = "4": sText = "Closed FollowUp"
        Case Is = "5": sText = "Trial Suspended"
        Case Is = "8": sText = "New Version"
        Case Is = "10": sText = "Patient Data"
        Case Is = "11": sText = "Mail"
        Case Is = "16,17,18,19": sText = "Locking/Freezing"
        Case Is = "20,21,22": sText = "Unlocking"
        Case Is = "30": sText = "Lab Definition Server To Site"
        Case Is = "31": sText = "Lab Definition Site To Server"
        Case Is = "32": sText = "User"
        Case Is = "33": sText = "User Role"
        Case Is = "34": sText = "Password Change"
        Case Is = "35": sText = "Role"
        Case Is = "36": sText = "System Log"
        Case Is = "37": sText = "User Log"
        Case Is = "38": sText = "Restore User Role"
        Case Is = "40": sText = "Password Policy"
        Case Is = "18|3": sText = "SDV"
        Case Is = "19|2": sText = "Note"
        Case Is = "20|0": sText = "Discrepancy"
    End Select
    
    GetEnumText = sText

Exit Function:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsType.GetEnumText"
End Function

'-------------------------------------------------------------------------------------
Private Function GetCollectionKey(sItem As String) As String
'-------------------------------------------------------------------------------------
'receives the item and returns the key
'-------------------------------------------------------------------------------------
Dim sText As String

    GetCollectionKey = ""
    
    On Error GoTo ErrHandler

    Select Case sItem
        Case Is = "New Trial": sText = "0"
        Case Is = "In Preparation": sText = "1"
        Case Is = "Trial Open": sText = "2"
        Case Is = "Closed Recruitment": sText = "3"
        Case Is = "Closed FollowUp": sText = "4"
        Case Is = "Trial Suspended": sText = "5"
        Case Is = "New Version": sText = "6"
        Case Is = "Patient Data": sText = "7"
        Case Is = "Mail": sText = "8"
        Case Is = "Locking/Freezing": sText = "9"
        Case Is = "Unlocking": sText = "10"
        Case Is = "Lab Definition Server To Site": sText = "11"
        Case Is = "Lab Definition Site To Server": sText = "12"
        Case Is = "User": sText = "13"
        Case Is = "User Role": sText = "14"
        Case Is = "Password Change": sText = "15"
        Case Is = "Role": sText = "16"
        Case Is = "Password Policy": sText = "17"
        Case Is = "SDV": sText = "18"
        Case Is = "Note": sText = "19"
        Case Is = "Discrepancy": sText = "20"
        Case Is = "System Log": sText = "21"
        Case Is = "User Log": sText = "22"
        Case Is = "Restore User Role": sText = "23"
        
    End Select

    GetCollectionKey = sText

Exit Function:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsType.GetCollectionKey"
End Function

'-------------------------------------------------------------------------
Private Function Selectionmade() As Boolean
'-------------------------------------------------------------------------
'
'-------------------------------------------------------------------------
Dim i As Integer

    On Error GoTo ErrHandler
    
    Selectionmade = False
    
    For i = 0 To lstType.ListCount - 1
        If lstType.Selected(i) = True Then
            Selectionmade = True
            Exit Function
        End If
    Next

Exit Function:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsType.Selectionmade"
End Function
