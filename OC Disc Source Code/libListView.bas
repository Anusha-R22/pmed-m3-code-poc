Attribute VB_Name = "libListView"
'----------------------------------------------------------------------------------------'
'   File:       modListView.bas
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   Author:     Toby Aldridge, April 2001
'   Purpose:    ListView functions
'----------------------------------------------------------------------------------------'
' Revisions:
'
'----------------------------------------------------------------------------------------'

Option Explicit

' windows api declerations
Public Const LVM_FIRST As Long = &H1000
Public Const LVM_SETCOLUMNWIDTH As Long = LVM_FIRST + 30

' windows api declerations
Public Enum LVSCW_Styles
   LVSCW_AUTOSIZE = -1
   LVSCW_AUTOSIZE_USEHEADER = -2
End Enum

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'----------------------------------------------------------------------------------------'
Public Sub lvw_SetColWidth(lvw As ListView, ByVal ColumnIndex As Long, ByVal Style As LVSCW_Styles)
'----------------------------------------------------------------------------------------'
' autoresize a column so no need to specify the column size in forms
'----------------------------------------------------------------------------------------'
   
   ' if you include the header in the sizing then the last column will
   ' automatically size to fill the remaining listview width.
   With lvw
      ' verify that the listview is in report view and that the column exists
      If .View = lvwReport Then
         If ColumnIndex >= 1 And ColumnIndex <= .ColumnHeaders.Count Then
            Call SendMessage(.hWnd, LVM_SETCOLUMNWIDTH, ColumnIndex - 1, ByVal Style)
         End If
      End If
   End With
 
End Sub
'----------------------------------------------------------------------------------------'
Public Sub lvw_SetAllColWidths(lvw As ListView, ByVal Style As LVSCW_Styles)
'----------------------------------------------------------------------------------------'
' autoresize a column so no need to specify sizes in forms
'----------------------------------------------------------------------------------------'
Dim ColumnIndex As Long

   '--- loop through all of the columns in the listview and size each
   With lvw
      For ColumnIndex = 1 To .ColumnHeaders.Count
         lvw_SetColWidth lvw, ColumnIndex, Style
      Next ColumnIndex
   End With
 
End Sub


'----------------------------------------------------------------------------------------'
Public Function lvw_ListItembyTag(lvw As ListView, sTag As String) As ListItem
'----------------------------------------------------------------------------------------'
'return a listitem by its tag
'----------------------------------------------------------------------------------------'

Dim olstItem As MSComctlLib.ListItem
    For Each olstItem In lvw.ListItems
        If olstItem.Tag = sTag Then
            Set lvw_ListItembyTag = olstItem
            Exit Function
        End If
    Next
End Function

'----------------------------------------------------------------------------------------'
Public Sub lvw_HighlightListItembyTag(lvw As ListView, sTag As String)
'----------------------------------------------------------------------------------------'
'selects and makes bold the listitem with sTag tag
'undoes the bold on all otheres
'----------------------------------------------------------------------------------------'

Dim olistItem As MSComctlLib.ListItem
Dim i As Integer
    For Each olistItem In lvw.ListItems
        With olistItem
            If .Tag = sTag Then
                .Bold = True
                'TA 15/12/2000: changed to count rather than count - 1
                For i = 1 To .ListSubItems.Count
                    .ListSubItems(i).Bold = True
                Next
                .Selected = True
                .EnsureVisible
            Else
                .Bold = False
                For i = 1 To .ListSubItems.Count
                    .ListSubItems(i).Bold = False
                Next
            End If
        End With
    Next
   
    lvw.Refresh

End Sub

'----------------------------------------------------------------------------------------'
Public Sub lvw_Sort(lvw As ListView, ByVal ColumnHeader As ColumnHeader)
'----------------------------------------------------------------------------------------'
' sort listview according to column click
'----------------------------------------------------------------------------------------'
    If Not lvw.Sorted Then
        lvw.Sorted = True
    End If
    
    If lvw.SortKey = ColumnHeader.Index - 1 Then
        If lvw.SortOrder = lvwDescending Then
            lvw.SortOrder = lvwAscending
        Else
            lvw.SortOrder = lvwDescending
        End If
    Else
        lvw.SortKey = ColumnHeader.Index - 1
        lvw.SortOrder = lvwAscending
    End If
    
End Sub

'--------------------------------------------------------------------------------------------------------
Public Sub lvw_FromArray(lvw As ListView, vData As Variant, vHeader As Variant, Optional vCols As Variant = "", Optional lKeyCol As Long = -1)
'--------------------------------------------------------------------------------------------------------
' Populates a listview control from a data array
'
' MLM 14/09/01: If vData is Null, only display column headings in the listview.
'--------------------------------------------------------------------------------------------------------
Dim i As Long
Dim j As Long
Dim lFields As Long
Dim sValue As String

    lFields = UBound(vHeader)
    
    lvw.ListItems.Clear
    lvw.ColumnHeaders.Clear
    
    For i = 0 To lFields
        lvw.ColumnHeaders.Add , , vHeader(i), 3000
    Next
    
    If Not IsNull(vData) Then
        
        'if vCols wasn't specified, default to displaying the first columns from the array, in order.
        If VarType(vCols) = vbString Then
            ReDim vCols(UBound(vData, 1)) As Long
            For i = 0 To UBound(vCols)
                vCols(i) = i
            Next
        End If
        For i = 0 To UBound(vData, 2)
            sValue = ConvertFromNull(vData(vCols(0), i), vbString)
            With lvw.ListItems.Add(, , sValue)
                If lKeyCol <> -1 Then
                    .Key = vData(lKeyCol, i)
                End If
                For j = 1 To lFields
                    sValue = ConvertFromNull(vData(vCols(j), i), vbString)
                    .SubItems(j) = sValue
                Next
            End With
        Next
    End If

    'adjust column widths
    Call lvw_SetAllColWidths(lvw, LVSCW_AUTOSIZE_USEHEADER)

End Sub



Public Function lvw_ListItembyText(lvw As ListView, sText As String, nColumn As Integer) As ListItem

    Dim olistItem As ListItem
    
    For Each olistItem In lvw.ListItems
        If nColumn = 0 Then
            If olistItem.Text = sText Then
                Set lvw_ListItembyText = olistItem
                Exit Function
            End If
            
        Else
            If olistItem.SubItems(nColumn) = sText Then
                Set lvw_ListItembyText = olistItem
                Exit Function
            End If
        End If
    Next
        
End Function
