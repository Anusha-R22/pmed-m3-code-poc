Attribute VB_Name = "libAPI"
'----------------------------------------------------------------------------------------'
'   File:       libAPI.bas
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   Author:     Zulfi Ahmed, Nov 2001
'   Purpose:    General API Library Functions
'----------------------------------------------------------------------------------------'
' Revisions:
'----------------------------------------------------------------------------------------'
Option Explicit

' Purpose:    Windows APIs to display explorer style view for selecting a folder
'----------------------------------------------------------------------------------------'
Public Type BrowseInfo
     lhwndOwner As Long
     pIDLRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Const BIF_USENEWUI = &H40

Public Const MAX_PATH = 260
Public Const LMEM_FIXED = &H0
Public Const LMEM_ZEROINIT = &H40
Public Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)
Public Const BFFM_INITIALIZED = 1
Public Const WM_USER = &H400
Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)

Public Declare Sub CoTaskMemFree Lib "OLE32.DLL" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Public Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Public Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long



'---------------------------------------------------------------------------------
Public Function BrowseForFolder(lhwndOwner As Long, sPrompt As String) As String
'---------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------
    'declare variables to be used
     Dim iNull As Integer
     Dim lpIDList As Long
     Dim lResult As Long
     Dim sPath As String
     Dim udtBI As BrowseInfo

    'initialise variables
     With udtBI
        .lhwndOwner = lhwndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = BIF_USENEWUI
     End With

    'Call the browse for folder API
     lpIDList = SHBrowseForFolder(udtBI)
     
    'get the resulting string path
     If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        'Try to free the block of task memory allocated during the call of SHBrowseForFolder routine
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then sPath = Left$(sPath, iNull - 1)
     End If

    'If cancel was pressed, sPath = ""
     BrowseForFolder = sPath

End Function
'--------------------------------------------------------------------------------------------------
Public Function BrowseForFolderByPath(sSelPath As String) As String
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim BI As BrowseInfo
Dim pidl As Long
Dim lpSelPath As Long
Dim sPath As String * MAX_PATH
  
    With BI
        .lhwndOwner = .lhwndOwner
        '.hOwner = Me.hWnd
         .pIDLRoot = 0
        ' .lpszTitle = "Pre-selecting folder using the folder's string."
        .lpszTitle = lstrcat(sSelPath, " ")
     
        'use these only if u have a value for optional parameter
     
        .lpfnCallback = FARPROC(AddressOf BrowseCallbackProcStr)
        lpSelPath = LocalAlloc(LPTR, Len(sSelPath) + 1)
        CopyMemory ByVal lpSelPath, ByVal sSelPath, Len(sSelPath) + 1
        .lParam = lpSelPath
        .ulFlags = BIF_USENEWUI
    End With
    
    pidl = SHBrowseForFolder(BI)
   
    If pidl Then
     
        If SHGetPathFromIDList(pidl, sPath) Then
           BrowseForFolderByPath = Left$(sPath, InStr(sPath, vbNullChar) - 1)
        End If
        
        Call CoTaskMemFree(pidl)
   
    End If
   
    Call LocalFree(lpSelPath)

End Function

'------------------------------------------------------------------------------------
Public Function FARPROC(pfn As Long) As Long
'------------------------------------------------------------------------------------
'
'-----------------------------------------------------------------------------------
  
  'A dummy procedure that receives and returns
  'the value of the AddressOf operator.
 
  'This workaround is needed as you can't assign
  'AddressOf directly to a member of a user-
  'defined type, but you can assign it to another
  'long and use that (as returned here)
  FARPROC = pfn

End Function
'--------------------------------------------------------------------------------------------------
Public Function BrowseCallbackProcStr(ByVal hWnd As Long, _
                                      ByVal uMsg As Long, _
                                      ByVal lParam As Long, _
                                      ByVal lpData As Long) As Long
'-------------------------------------------------------------------------------------------------
                                       
'Callback for the Browse STRING method.
'On initialization, set the dialog's
'pre-selected folder from the pointer
'to the path allocated as bi.lParam,
'passed back to the callback as lpData param.
'--------------------------------------------------------------------------------------------------
    Select Case uMsg
      Case BFFM_INITIALIZED
      
         Call SendMessage(hWnd, BFFM_SETSELECTIONA, _
                          True, ByVal lpData)
                          
         Case Else:
         
   End Select
          
End Function

'--------------------------------------------------------------------------------------------------
Public Function CreateGUID() As String
'--------------------------------------------------------------------------------------------------
' MLM 18/12/01  This function adapted from Knowledge Base article Q176790, which was
' (c) 2000 Gus Molina, apparently.
' To make it look like a real one, it should be formatted as {12345678-1234-1234-1234-123456789ABC}
'--------------------------------------------------------------------------------------------------

Dim udtGUID As GUID

    If (CoCreateGuid(udtGUID) = 0) Then
        CreateGUID = _
        String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
        String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
        String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
        IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
        IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
        IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
        IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
        IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
        IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
        IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
        IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
    End If

End Function
