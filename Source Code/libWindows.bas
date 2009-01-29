Attribute VB_Name = "libWindows"
'----------------------------------------------------------------------------------------'
'   File:       modWindows.bas
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   Author:     Toby Aldridge, April 2001
'   Purpose:    ListCtrl libray
'----------------------------------------------------------------------------------------'
' Revisions:
'
'----------------------------------------------------------------------------------------'

Option Explicit

'background colour of a control that contains invalid data
Public Const CTRL_INVALID_COLOUR = vbYellow


'----------------------------------------------------------------------------------------'
Public Sub Ctrl_Colour(conControl As Control, bValid As Boolean)
'----------------------------------------------------------------------------------------'
' if not valid colour control yellow
' if valid colour control with windowsbackground colour
'----------------------------------------------------------------------------------------'

    If bValid Then
        conControl.BackColor = vbWindowBackground
    Else
        conControl.BackColor = CTRL_INVALID_COLOUR
    End If

End Sub


Public Sub ListCtrl_ListSelect(lst As ListBox, sText As String)
Dim i As Long
    For i = 0 To lst.ListCount - 1
        If lst.List(i) = sText Then
            lst.Selected(i) = True
        End If
    Next
End Sub



Public Sub ListCtrl_FromArray(lst As ListBox, vData As Variant, Optional lCol As Long = 0)
Dim i As Long
    
    lst.Clear
    
    For i = 0 To UBound(vData, 2)
        lst.AddItem ConvertFromNull(vData(lCol, i), vbString)
    Next
    

End Sub


'----------------------------------------------------------------------------------------'
Public Sub ListCtrl_Pick(oCtrl As Control, lId As Long)
'----------------------------------------------------------------------------------------'
'pick a list tiem form its itemdata
'----------------------------------------------------------------------------------------'

Dim i As Long
    For i = 0 To oCtrl.ListCount - 1
        If oCtrl.ItemData(i) = lId Then
            oCtrl.ListIndex = i
            Exit For
        End If
    Next

End Sub

'---------------------------------------------------------------------
Public Sub FormCentre(frmForm As Form, Optional frmParent As Form = Nothing)
'---------------------------------------------------------------------
'   Centre form on screen or parent form
'---------------------------------------------------------------------

    If frmForm.WindowState = vbNormal Then
        If frmParent Is Nothing Then
            frmForm.Top = (Screen.Height - frmForm.Height) \ 2
            frmForm.Left = (Screen.Width - frmForm.Width) \ 2
        Else
            With frmParent
                frmForm.Top = .Top + ((.Height - frmForm.Height) \ 2)
                frmForm.Left = .Left + ((.Width - frmForm.Width) \ 2)
            End With
        End If
    End If
    
End Sub

'---------------------------------------------------------------------
Public Function FormIsLoaded(sFormName As String) As Boolean
'---------------------------------------------------------------------
'   Is an instance of this form with this name loaded?
'---------------------------------------------------------------------
Dim bLoaded As Boolean
Dim oForm As Form
  
    bLoaded = False
    
    For Each oForm In Forms   'iterate through the forms collection
        If oForm.Name = sFormName Then    'same name
            bLoaded = True
            Exit For                   'exit from loop
        End If
    Next

    Set oForm = Nothing

    FormIsLoaded = bLoaded
    
End Function

'---------------------------------------------------------------------
Public Function FormByName(sFormName As String) As Form
'---------------------------------------------------------------------
'   Return instance of this form if loaded - nothing if not
'---------------------------------------------------------------------
Dim oForm As Form
    
    For Each oForm In Forms   'iterate through the forms collection
        If oForm.Name = sFormName Then    'same name
            Set FormByName = oForm
            Exit For                   'exit from loop
        End If
    Next

    Set oForm = Nothing
    
End Function

