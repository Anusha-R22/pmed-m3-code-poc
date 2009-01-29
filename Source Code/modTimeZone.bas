Attribute VB_Name = "modTimeZone"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   File:       modTimeZone.bas
'   Author:     ZA, 26/09/2002
'   Purpose:    Various date/time related functions
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'Revisions

Option Explicit

Public Const LOCALE_SLANGUAGE       As Long = &H2  'localized name of language
Public Const LOCALE_SABBREVLANGNAME As Long = &H3  'abbreviated language name
Public Const LCID_INSTALLED         As Long = &H1  'installed locale ids
Public Const LCID_SUPPORTED         As Long = &H2  'supported locale ids
Public Const LCID_ALTERNATE_SORTS   As Long = &H4  'alternate sort locale ids

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
  (Destination As Any, Source As Any, ByVal Length As Long)


Public Declare Function EnumSystemLocales Lib "kernel32" Alias "EnumSystemLocalesA" _
  (ByVal lpLocaleEnumProc As Long, ByVal dwFlags As Long) As Long
   
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" _
(ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

'public dictionary to store the system locales
Public gdicSystemLocales As New Scripting.Dictionary

'-------------------------------------------------------------------------
Public Function EnumSystemLocalesProc(lpLocaleString As Long) As Long
'-------------------------------------------------------------------------
'application-defined callback function for EnumSystemLocales
'-------------------------------------------------------------------------
Dim nPos As Integer
Dim ldwLocaleDec As Long
Dim sdwLocaleHex As String
Dim sLocaleName As String
Dim sLocaleAbbrev As String

    'pad a string to hold the format
    sdwLocaleHex = Space$(32)
     
    'copy the string pointed to by the return value
    CopyMemory ByVal sdwLocaleHex, lpLocaleString, ByVal Len(sdwLocaleHex)
     
    'locate the terminating null
    nPos = InStr(sdwLocaleHex, Chr$(0))
     
    If nPos Then
       'strip the null
        sdwLocaleHex = Left$(sdwLocaleHex, nPos - 1)
        
        'we need the last 4 chrs - this
        'is the locale ID in hex
        sdwLocaleHex = (Right$(sdwLocaleHex, 4))
        
        'convert the string to a long
        ldwLocaleDec = CLng("&H" & sdwLocaleHex)
        
        'get the language and abbreviation for that locale
        sLocaleName = GetUserLocaleInfo(ldwLocaleDec, LOCALE_SLANGUAGE)
        
        'we are not using abbrreviation yet
        'sLocaleAbbrev = GetUserLocaleInfo(ldwLocaleDec, LOCALE_SABBREVLANGNAME)
       
    End If
    
    On Error Resume Next
    'add locale name into the dictionary and use the locale id as key
    gdicSystemLocales.Add ldwLocaleDec, sLocaleName
    
    'and return 1 to continue enumeration
    EnumSystemLocalesProc = 1
    
End Function

'-------------------------------------------------------------------------
Public Function GetUserLocaleInfo(ByVal dwLocaleID As Long, _
                                  ByVal dwLCType As Long) As String
'-------------------------------------------------------------------------
'
'-------------------------------------------------------------------------
Dim sReturn As String
Dim lSize As Long

    'call the function passing the Locale type
    'variable to retrieve the required size of
    'the string buffer needed
    lSize = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
    
    'if successful..
    If lSize Then
    
        'pad a buffer with spaces
        sReturn = Space$(lSize)
       
        'and call again passing the buffer
        lSize = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
     
        'if successful (lSize > 0)
        If lSize Then
      
            'lSize holds the size of the string
            'including the terminating null
            GetUserLocaleInfo = Left$(sReturn, lSize - 1)
      
      End If
   
    End If
    
End Function


