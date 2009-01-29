VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeZone"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------'
'   File:       Timezone.cls
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Ronald Schravendeel, September 2002
'   Purpose:    Timezone Management
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'
' REVISIONS:
'   RS 21/09/2002   Initial Version
'   RS 08/10/2002   Corrected ConvertToLocalTime routine
'   RS 29/10/2002   Corrected bug in calculating offset with regards to DST
'   NCJ 7 Mar 03 - Corrected bug in ConvertDateTimeToLocal
'   RS 19/05/2003   Handle unknown Timezones (such as +11 (Solomon Islands)
'----------------------------------------------------------------------------------------'

Private Declare Function GetTimeZoneInformation _
   Lib "KERNEL32" (lpTimeZoneInformation As _
   TIME_ZONE_INFORMATION) As Long

Private Declare Function GetLocaleInfo Lib "KERNEL32" _
    Alias "GetLocaleInfoA" (ByVal lLocale As Long, _
        ByVal lLocaleType As Long, ByVal sLCData As String, _
        ByVal lBufferLength As Long) As Long

Private Const LOCALE_SSHORTDATE As Long = &H1F
Private Const LOCALE_STIMEFORMAT As Long = &H1003

Private Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
   Bias As Long
   StandardName(0 To 63) As Byte
   StandardDate As SYSTEMTIME
   StandardBias As Long
   DaylightName(0 To 63) As Byte
   DaylightDate As SYSTEMTIME
   DaylightBias As Long
End Type

Private Const TIME_ZONE_ID_INVALID = &HFFFFFFFF
Private Const TIME_ZONE_ID_UNKNOWN = 0
Private Const TIME_ZONE_ID_STANDARD = 1
Private Const TIME_ZONE_ID_DAYLIGHT = 2

Private mlTimezoneID As Long
Private msTimezoneName As String
Private mnOffsetToGMT As Integer           ' Minutes to add to get GMT equivalent

Private mlDifferenceInMins As Long
Private mlDaylightSavingInMins As Long
Private mTZ As TIME_ZONE_INFORMATION

Private sLocalDateFormat As String          ' RS 20/01/03  Used for dataentry in local format
Private sLocalTimeFormat As String          ' RS 20/01/03  Used for dataentry in local format

'-----------------------------------------------------------------------------------
Public Property Get LocalDateFormat() As String
'-----------------------------------------------------------------------------------
' Get the local date format
'-----------------------------------------------------------------------------------
    LocalDateFormat = sLocalDateFormat
End Property

'-----------------------------------------------------------------------------------
Public Property Get LocalTimeFormat() As String
'-----------------------------------------------------------------------------------
' Get the local time format
'-----------------------------------------------------------------------------------
    LocalTimeFormat = sLocalTimeFormat
End Property

'-----------------------------------------------------------------------------------
Public Property Get TimezoneOffset() As Integer
'-----------------------------------------------------------------------------------
' Get the TimeZoneOffset on the current PC to store with the Timestamp
'-----------------------------------------------------------------------------------
    
    TimezoneOffset = mnOffsetToGMT
    
End Property

'-----------------------------------------------------------------------------------
Public Sub SetTimezoneOffset(TimezoneOffset As Integer)
'-----------------------------------------------------------------------------------
' Set the Timezone 'manually', (used from WEB client)
'-----------------------------------------------------------------------------------
    mnOffsetToGMT = TimezoneOffset
End Sub

'-----------------------------------------------------------------------------------
Private Sub Class_Initialize()
'-----------------------------------------------------------------------------------
    
    ' Get Timezone Info from system settings
    ReadTimeZone
    
    sLocalDateFormat = DateFormat()
    sLocalTimeFormat = TimeFormat()
    ' RS 29/10/2002: Moved to ReadTimezone
    ' mnOffsetToGMT = mTZ.Bias + mTZ.DaylightBias
    
End Sub

'-----------------------------------------------------------------------------------
Private Sub ReadTimeZone()
'-----------------------------------------------------------------------------------
' Get the timezone the computer has been set in
' Find out if Daylight saving time applies
' Set module-level variables
' NCJ 9/11/00 - Changed so that it correctly takes account of Daylight Saving
' RS 20/05/2003 - Handle Unknown Timezones
'-----------------------------------------------------------------------------------
Dim lRet As Long

    ' Initialise variables
    mlDaylightSavingInMins = 0
    mlDifferenceInMins = 0
    
   ' lRet tells you whether it's Standard time or Daylight Savings time
    lRet = GetTimeZoneInformation(mTZ)
    
    Select Case lRet
    Case TIME_ZONE_ID_INVALID
        ' The call failed for some reason
        ' Still set Bias???
    Case TIME_ZONE_ID_UNKNOWN
        ' Although Timezone ID is unknown, we can still set the offset
        mlDifferenceInMins = mTZ.Bias
        mlDaylightSavingInMins = mTZ.DaylightBias
    Case TIME_ZONE_ID_STANDARD
        ' Standard Time - pick up general time zone difference
        mlDifferenceInMins = mTZ.Bias
    Case TIME_ZONE_ID_DAYLIGHT
        ' Daylight Savings Time - pick up general time zone difference AND daylight saving
        mlDifferenceInMins = mTZ.Bias
        mlDaylightSavingInMins = mTZ.DaylightBias
    End Select
    
    ' RS 29/10/2002: Make sure that DST is taken into account
    mnOffsetToGMT = mlDifferenceInMins + mlDaylightSavingInMins

End Sub

'-----------------------------------------------------------------------------------
Public Function ConvertDateTimeToLocal(dblTimestamp, lOffset) As Double
'-----------------------------------------------------------------------------------
' Converts the given Timestamp/TimeZone into the local equivalent
'-----------------------------------------------------------------------------------
Dim dtGMT As Date

    ReadTimeZone

    ' Convert to the equivalent GMT Time
    ' RS 08/10/2002 Correction: used to be -lOffset
    dtGMT = DateAdd("n", lOffset, CDate(dblTimestamp))
    
    ' Convert to the local time
    ' NCJ 7 Mar 03 - Must use variables already set up in ReadTimeZone!
'    ConvertDateTimeToLocal = CDbl(DateAdd("n", -mTZ.Bias - mTZ.DaylightBias, dtGMT))
    ConvertDateTimeToLocal = CDbl(DateAdd("n", -mnOffsetToGMT, dtGMT))
    
End Function


'-----------------------------------------------------------------------------------
Public Property Get DateFormat() As String
'-----------------------------------------------------------------------------------
' Get Local Date & Time Format
'
' This function will return the Locale date format for the system. Note that the
' returned Year is always formatted to 'YYYY' regardless, to ensure compliance with
' Y2k stuff.
'
    Dim sDateFormat As String
    
    On Error GoTo vbErrorHandler
    
    sDateFormat = GetLocaleString(LOCALE_SSHORTDATE)
'
' Make sure we always have YYYY format for y2k
'
    If InStr(1, sDateFormat, "YYYY", vbTextCompare) = 0 Then
        Replace sDateFormat, "YY", "YYYY"
    End If
        
    DateFormat = sDateFormat
    
    Exit Property

vbErrorHandler:
    Err.Raise Err.Number, "CGLocaleInfo GetDateFormat", Err.Description
End Property

'-----------------------------------------------------------------------------------
Public Property Get TimeFormat() As String
'-----------------------------------------------------------------------------------
' This function returns the locale's defined Time Format.
'-----------------------------------------------------------------------------------

    TimeFormat = GetLocaleString(LOCALE_STIMEFORMAT)
    
    Exit Property
    
vbErrorHandler:
    Err.Raise Err.Number, "CGLocaleInfo GetTimeFormat", Err.Description
End Property

'-----------------------------------------------------------------------------------
Private Function GetLocaleString(ByVal lLocaleNum As Long) As String
'-----------------------------------------------------------------------------------
' Generic routine to get the locale string from the Operating system.
'-----------------------------------------------------------------------------------
Dim lBuffSize As String
Dim sBuffer As String
Dim lRet As Long
'
' Create a string buffer large enough to hold the returned value, 256 should
' be more than enough
'
    lBuffSize = 256
    sBuffer = String$(lBuffSize, vbNullChar)
'
' Get the information from the registry
'
    lRet = GetLocaleInfo(LOCALE_USER_DEFAULT, lLocaleNum, sBuffer, lBuffSize)
'
' If lRet > 0 then success - lret is the size of the string returned
'
    If lRet > 0 Then
        GetLocaleString = Left$(sBuffer, lRet - 1)
    End If
    
End Function







