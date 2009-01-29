Attribute VB_Name = "modComboManipulation"
Option Explicit

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                           ByVal wMsg As Long, _
                                                                           ByVal wParam As Long, _
                                                                           ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDc As Long, _
                                                                  ByVal lpStr As String, _
                                                                  ByVal nCount As Long, _
                                                                  lpRect As RECT, _
                                                                  ByVal wFormat As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, _
                                                   ByVal hObject As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDc As Long, _
                                                                                         ByVal lpsz As String, _
                                                                                         ByVal cbString As Long, _
                                                                                         lpSize As SIZE) As Long

Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hDc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const CB_GETLBTEXTLEN = &H149
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_GETDROPPEDWIDTH = &H15F
Private Const CB_SETDROPPEDWIDTH = &H160

Private Const ANSI_FIXED_FONT = 11
Private Const ANSI_VAR_FONT = 12
Private Const SYSTEM_FONT = 13
Private Const DEFAULT_GUI_FONT = 17 'win95/98 only

Private Const SM_CXHSCROLL = 21
Private Const SM_CXHTHUMB = 10
Private Const SM_CXVSCROLL = 2

Private Const DT_CALCRECT = &H400
 
Private Type SIZE
    cx As Long
    cy As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Function GetFontDialogUnits(TargetForm As Form) As Long

   Dim hFont As Long
   Dim hFontOld As Long
   Dim r As Long
   Dim avgWidth As Long
   Dim hDc As Long
   Dim tmp As String
   Dim sz As SIZE
   
  'get the hdc to the main window
   hDc = GetDC(TargetForm.hwnd)
   
  'with the current font attributes, select the font
   hFont = GetStockObject(ANSI_VAR_FONT)
   hFontOld = SelectObject(hDc, hFont&)
   
  'get its length, then calculate the average character width
   tmp = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
   Call GetTextExtentPoint32(hDc, tmp, 52, sz)
   avgWidth = (sz.cx \ 52)
   
  're-select the previous font & delete the hDc
   Call SelectObject(hDc, hFontOld)
   Call DeleteObject(hFont)
   Call ReleaseDC(TargetForm.hwnd, hDc)
   
  'return the average character width
   GetFontDialogUnits = avgWidth

End Function

Public Sub SetComboDropdownWidth(TargetForm As Form, TargetCombo As ComboBox)

   Dim cwidth As Long
   Dim nCount As Long
   Dim NumOfChars As Long
   Dim LongestComboItem As Long
   Dim avgCharWidth As Long
   Dim NewDropDownWidth As Long
   
  'loop through the combo entries, using SendMessageLong
  'with CB_GETLBTEXTLEN to determine the longest item
  'in the dropdown portion of the combo

   For nCount = 0 To TargetCombo.ListCount - 1

      NumOfChars = SendMessageLong(TargetCombo.hwnd, CB_GETLBTEXTLEN, nCount, 0)
      If NumOfChars > LongestComboItem Then LongestComboItem = NumOfChars

    Next
   
  'get the average size of the characters using the
  'GetFontDialogUnits API. Because a dummy string is
  'used in GetFontDialogUnits, avgCharWidth is an
  'approximation based on that string.
   avgCharWidth = GetFontDialogUnits(TargetForm)
   
  'compute the size the dropdown needs to be to accommodate
  'the longest string. Here I subtract 2 because I find that
  'on my system, using the dummy string in GetFontDialogUnits,
  'the width is just a bit too wide.
   NewDropDownWidth = (LongestComboItem * 1.5) * avgCharWidth
   
  'resize the dropdown portion of the combo box
   Call SendMessageLong(TargetCombo.hwnd, CB_SETDROPPEDWIDTH, NewDropDownWidth, 0)
   
  'reflect the new dropdown list width in Label2 and in Text1
   'cwidth = SendMessageLong(TargetCombo.hWnd, CB_GETDROPPEDWIDTH, 0, 0)

   
  'finally, drop the list down by code to show the new size
   'Call SendMessageLong(TargetCombo.hWnd, CB_SHOWDROPDOWN, True, 0)

End Sub

