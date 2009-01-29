Attribute VB_Name = "modMACRODM_UI"
'----------------------------------------------------
' File: modMACRODM_UI.bas
' Copyright: InferMed, August 2001, All Rights Reserved
' Author: Nicky Johns, InferMed, August 2001
' Purpose: Declarations for MACRO 2.2 Data Management User Interface
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 8 Aug 01 - Initial Development
' TA 10/10/01: simple routine to return icon names by status
' ZA 14/08/01: Added gblnValidString function from Macro 2.2
' NCJ 21 Sep 01 - Multimedia File constants
' NCJ 15 Nov 01 - Control element types now global consts here
' MLM 02/09/02: Added bunch of constants.
' NCJ 8 Nov 02 - Added gn_HOTLINK
'----------------------------------------------------

Option Explicit

' NCJ 21 Sep 01 - The labels for a MultiMedia question button
Public Const gsATTACH_FILE = "Attach file ..."
Public Const gsVIEW_FILE = "View attachment"

' This is used to show that this is a new subject with no subject id yet
Public Const g_NEW_SUBJECT_ID = -1

Public Const g_DATAENTRY_FORM_NAME = "frmEFormDataEntry"
Public Const g_SCHEDULE_FORM_NAME = "frmSchedule"

Public Const glYELLOW = &H80FFFF
Public Const glLIGHT_GREY = &HEEEEEE
Public Const glDARK_GREY = &HA9A9A9

'MLM 02/09/02: Constants for visit and eForm dates
Public Const glDATE_FONT_COLOUR = &HAAAAAA
Public Const gsDATE_FONT = "Verdana"
Public Const gnDATE_FONT_SIZE = 8

' The eFormElement Control types
'1 = Text Box
'2 = Option Buttons
'4 = PopUp List
'8 = Calendar
'16 = Rich Text Box
'32 = Attachment
'512 = Mask Edit Box
'258 = Push Buttons
'16385 = Line
'16386 = Comment
'16388 = Picture

Public Const gn_TEXT_BOX = 1
Public Const gn_OPTION_BUTTONS = 2
Public Const gn_POPUP_LIST = 4
Public Const gn_CALENDAR = 8
Public Const gn_RICH_TEXT_BOX = 16
Public Const gn_ATTACHMENT = 32
Public Const gn_PUSH_BUTTONS = 258
Public Const gn_MASK_ED_BOX = 512

Public Const gn_LINE = 16385
Public Const gn_COMMENT = 16386
Public Const gn_PICTURE = 16388
Public Const gn_HOTLINK = 16390

'---------------------------------------------------------------------
Public Function GetIconFromStatus(nStatus As Integer) As String
'----------------------------------------------------
' TA this routine is temporarily here
' until either we establish a new way of displaying icons
' or we restructure the current way
'----------------------------------------------------'

    Select Case nStatus
    Case eStatus.InvalidData
        GetIconFromStatus = gsVALIDATION_MANDATORY_LABEL
    Case eStatus.Warning
        GetIconFromStatus = DM30_ICON_WARNING
    Case eStatus.OKWarning
        GetIconFromStatus = DM30_ICON_OK_WARNING
    Case eStatus.Missing
          GetIconFromStatus = DM30_ICON_MISSING
    Case eStatus.Success, eStatus.Inform
           GetIconFromStatus = DM30_ICON_OK
    Case eStatus.NotApplicable
           GetIconFromStatus = DM30_ICON_NA
    Case eStatus.Unobtainable
           GetIconFromStatus = DM30_ICON_UNOBTAINABLE
    Case Else
        GetIconFromStatus = ""
    End Select

End Function

'
