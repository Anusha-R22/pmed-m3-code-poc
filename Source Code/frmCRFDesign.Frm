VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3BA5E907-B7F4-4504-90CD-12033A3F93F8}#1.1#0"; "MACROOptionGroup30.ocx"
Begin VB.Form frmCRFDesign 
   Caption         =   "Case Report Form"
   ClientHeight    =   8580
   ClientLeft      =   6150
   ClientTop       =   6900
   ClientWidth     =   12060
   ClipControls    =   0   'False
   Icon            =   "frmCRFDesign.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8580
   ScaleWidth      =   12060
   Begin VB.PictureBox picUsedForPrinting 
      Height          =   540
      Left            =   300
      ScaleHeight     =   480
      ScaleWidth      =   1215
      TabIndex        =   4
      Top             =   3900
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.HScrollBar hsbCRFPage 
      Height          =   375
      Left            =   4935
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   6495
   End
   Begin VB.Frame frmCRFPage 
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   9615
      Begin VB.PictureBox picCRFPage 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         DataSource      =   "datPicture"
         Height          =   6015
         Left            =   60
         ScaleHeight     =   5955
         ScaleWidth      =   9195
         TabIndex        =   5
         Top             =   60
         Width           =   9255
         Begin VB.PictureBox picDictionary 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   8520
            Picture         =   "frmCRFDesign.frx":12A2
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   31
            Top             =   540
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton cmdDictionary 
            Height          =   375
            Index           =   0
            Left            =   8040
            Picture         =   "frmCRFDesign.frx":1628
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   480
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdExpand 
            Caption         =   "..."
            Height          =   375
            Index           =   0
            Left            =   7560
            TabIndex        =   29
            Top             =   480
            Visible         =   0   'False
            Width           =   375
         End
         Begin MACROOptionGroup30.OptionGroup optCRFGroup 
            Height          =   375
            Index           =   0
            Left            =   5760
            TabIndex        =   25
            Top             =   3960
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
         End
         Begin VB.PictureBox picQuGroup 
            AutoRedraw      =   -1  'True
            Height          =   1935
            Index           =   0
            Left            =   180
            MouseIcon       =   "frmCRFDesign.frx":19B8
            MousePointer    =   99  'Custom
            ScaleHeight     =   1875
            ScaleWidth      =   2715
            TabIndex        =   20
            Top             =   4020
            Visible         =   0   'False
            Width           =   2775
            Begin MSComCtl2.FlatScrollBar hsbQuGroup 
               Height          =   240
               Index           =   0
               Left            =   120
               TabIndex        =   27
               Top             =   1560
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   423
               _Version        =   393216
               MousePointer    =   1
               Arrows          =   65536
               Orientation     =   1245185
            End
            Begin MSComCtl2.FlatScrollBar vsbQuGroup 
               Height          =   1275
               Index           =   0
               Left            =   2280
               TabIndex        =   23
               Top             =   300
               Visible         =   0   'False
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   2249
               _Version        =   393216
               MousePointer    =   1
               Orientation     =   1245184
            End
            Begin VB.PictureBox picFixed 
               AutoRedraw      =   -1  'True
               Height          =   1515
               Index           =   0
               Left            =   60
               ScaleHeight     =   1455
               ScaleWidth      =   2055
               TabIndex        =   21
               Top             =   60
               Visible         =   0   'False
               Width           =   2115
               Begin VB.PictureBox picHeaders 
                  Height          =   195
                  Index           =   0
                  Left            =   120
                  ScaleHeight     =   135
                  ScaleWidth      =   1755
                  TabIndex        =   28
                  Top             =   60
                  Width           =   1815
               End
               Begin VB.PictureBox picScroller 
                  AutoRedraw      =   -1  'True
                  Height          =   1035
                  Index           =   0
                  Left            =   60
                  ScaleHeight     =   975
                  ScaleWidth      =   1815
                  TabIndex        =   22
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   1875
               End
            End
         End
         Begin VB.TextBox txtEditCaption 
            Height          =   375
            Left            =   3360
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   19
            Top             =   4020
            Width           =   1995
         End
         Begin VB.OptionButton optCRFElement 
            Caption         =   "optCRFElement"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2340
            TabIndex        =   14
            Top             =   1260
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.PictureBox picCRFElement 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            FillColor       =   &H00FFFFFF&
            Height          =   615
            Index           =   0
            Left            =   5880
            ScaleHeight     =   615
            ScaleWidth      =   3255
            TabIndex        =   13
            Top             =   4440
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.ComboBox cboCRFElement 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1800
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.CommandButton cmdPopupList 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2940
            MousePointer    =   1  'Arrow
            Picture         =   "frmCRFDesign.frx":1CC2
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   360
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdAttachment 
            Caption         =   "Attach file ..."
            Height          =   375
            Index           =   0
            Left            =   1920
            TabIndex        =   10
            Top             =   3240
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdCalendar 
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   9
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   4200
            MousePointer    =   1  'Arrow
            Picture         =   "frmCRFDesign.frx":1DC4
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   840
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.OptionButton optGraphical 
            Caption         =   "optGraphical"
            Height          =   375
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
         End
         Begin RichTextLib.RichTextBox rtfCRFElement 
            Height          =   375
            Index           =   0
            Left            =   3840
            TabIndex        =   6
            Top             =   360
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393217
            ScrollBars      =   3
            TextRTF         =   $"frmCRFDesign.frx":1EC6
         End
         Begin MSACAL.Calendar calCRFElement 
            Height          =   2295
            Left            =   3480
            TabIndex        =   8
            Top             =   1440
            Visible         =   0   'False
            Width           =   4095
            _Version        =   524288
            _ExtentX        =   7223
            _ExtentY        =   4048
            _StockProps     =   1
            BackColor       =   12632256
            Year            =   1997
            Month           =   8
            Day             =   14
            DayLength       =   1
            MonthLength     =   2
            DayFontColor    =   0
            FirstDay        =   1
            GridCellEffect  =   1
            GridFontColor   =   10485760
            GridLinesColor  =   -2147483632
            ShowDateSelectors=   -1  'True
            ShowDays        =   -1  'True
            ShowHorizontalGrid=   -1  'True
            ShowTitle       =   0   'False
            ShowVerticalGrid=   -1  'True
            TitleFontColor  =   10485760
            ValueIsNull     =   0   'False
            BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   1200
            Top             =   3240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSMask.MaskEdBox mskCRFElement 
            Height          =   375
            Index           =   0
            Left            =   1440
            TabIndex        =   15
            Top             =   2280
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            HideSelection   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "-"
         End
         Begin MSComctlLib.ListView lvwPopupList 
            Height          =   975
            Left            =   4140
            TabIndex        =   16
            Top             =   4680
            Visible         =   0   'False
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   1720
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Value"
               Object.Width           =   5292
            EndProperty
         End
         Begin MSComctlLib.ImageList imglistSmallIcons 
            Left            =   120
            Top             =   3240
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
         End
         Begin VB.Shape shpLasso 
            BorderStyle     =   3  'Dot
            Height          =   555
            Left            =   300
            Top             =   2040
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Shape shpSelectControl 
            BorderStyle     =   3  'Dot
            Height          =   195
            Index           =   0
            Left            =   180
            Top             =   2640
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label shpDragHandle 
            BackColor       =   &H00FF0000&
            Caption         =   " "
            Height          =   240
            Left            =   180
            TabIndex        =   26
            Top             =   1740
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label lblGroupHeader 
            AutoSize        =   -1  'True
            Caption         =   "lblGroupHeader"
            Height          =   195
            Index           =   0
            Left            =   5880
            TabIndex        =   24
            Top             =   240
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Line linPageBreak 
            BorderStyle     =   3  'Dot
            Index           =   0
            Visible         =   0   'False
            X1              =   0
            X2              =   9240
            Y1              =   860
            Y2              =   860
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00000000&
            Index           =   0
            Visible         =   0   'False
            X1              =   0
            X2              =   9240
            Y1              =   3120
            Y2              =   3120
         End
         Begin VB.Label lblCRFElement 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblCRFElement"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1260
            TabIndex        =   18
            Top             =   360
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblCaption"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Visible         =   0   'False
            Width           =   690
         End
      End
      Begin VB.VScrollBar vsbCRFPage 
         Height          =   6015
         Left            =   9345
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   0
         Width           =   375
      End
   End
   Begin MSComctlLib.TabStrip tabCRF 
      Height          =   405
      Left            =   720
      TabIndex        =   2
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCRFDesign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998 - 2006. All Rights Reserved
'   File:       frmCRFDesign.frm
'   Author:     Andrew Newbigging August 1997
'   Purpose:    Allows user to add new CRF elements (data items dragged from frmDataList
'   or visual elements like lines, comments or pictures) to a CRF page.  Elements may
'   be moved around using drag and drop, and format details (font, colour) may also be
'   modified.  Note that any changes to the VB controls on this form MUST also be changed
'   on frmCRFDesign.  These forms were split apart because the size was getting too large
'   for VB to handle.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
' WillC Deleted previous comments 12/7/00
'       Force setfocus so that txtEditCaption has definitely lost focus
'   40  Andrew Newbigging       9/11/98     SPR 579
'       Modified AddDataItemToCRF so that it checks that the DataList form exists
'       before trying to update it
'   40  Andrew Newbigging       9/11/98     SPR 557
'       Modified AddDataItemToCRF so that it checks if user has cancelled the
'       operation
'   41  Andrew Newbigging       24/11/98    SR 616
'   Added validation to check that code starts with alphabetic character (required by Prolog)
'   on InsertCRFpage
'       Mo Morris               10/12/98    SPR 451
'       UpdateCRFPageBackgroundColour changed so that it calls BuildCRFPage after a
'       background colour change, for the purpose of displaying the new colour.
'       Mo Morris               6/1/99      SR 668
'       InsertCRFPage now makes an additional call to gblnNotAReservedWord
'       Mo Morris               29/4/99     SR 799
'       UpdateFieldFont changed to trap the error condition of their being no default printer
'       UpdateFieldFont changed to display the font settings of the current control when
'       the fontchanging commondialog is opened.
'   Andrew Newbigging           30/4/99     SR 832
'       Added check in SnapToGrid to ensure that a CRFElement has been selected before
'       trying to do the snap
'       Mo Morris               30/4/99     SR 833
'       MoveCRFElement changed (in line with AddDataItemToCRF) so that it now checks for elements
'       dragged into the bottom 5000 twips of the form and extends it.
'       picCRFPage now has a fixed single border so that it can be distinguished from the
'       surrounding area when both are grey.
'       Mo Morris               4/5/99      SR 877
'       DragOver changed. See Subroutine for reason
'       Andrew Newbigging       4/5/99      SR 906
'       Don't need to set mouse icon if a data item or form is being dragged across
'       the page from the data list window.  Modified DragOver subroutine.
'       Mo Morris               4/5/99      SR 889
'       picCRFPage_MouseDown changed, call to MouseDownOnCRFElement added to enable the correct
'       menu options for the new comment.
'       Mo Morris               6/8/99
'       InsertCRFPage changed to incorporate database changes
'   41  Paul Norris             20/08/99    Updated Delete() to use generic AttemptDeleteCRFPage()
'   42  Paul Norris             20/08/99    SR1540 Added DisplayDRFPage() to tab to a crf page
'   43  Paul Norris             02/09/99    Conserve memory by setting the icon of lblCRFElement
'   PN  20/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   PN  26/09/99    Amended DeleteCRFELement(), Form_Activate(),
'                   tabCRF_Click() for mtm1.6 changes
'   NCJ 1 Oct 99    Use gsDOCUMENTS_PATH
'   MM/NCJ 1 Oct 99 Changed MouseMoveOverCRFElement to correctly handle Visual elements
'   NCJ 15 Oct 99 Changed SnapToGrid and MouseMoveOverCRFElement
'       to fix captions/controls parting company on CRF page when dragging
'   NCJ 18 Oct 99 - Check a data item CanBeDropped onto CRF page
'   NCJ 8 Nov 99 - Tweaked WC's error handlers to do Exit Sub before ErrHandler label
'   Mo Morris   8/11/99, 11/11/99 DAO to ADO conversion
'   NCJ 9/11/99 - Ignore errors during unload
'   Mo Morris   15/11/99    DAO to ADO conversion
'   NCJ 13 Dec 99   Ids to Long
'   Mo Morris   14/2/00     sections of InsertCRFPage re-written
'   SDM 22/02/00 SR3074 Grid display and alignment
'   TA 29/04/2000   subclassing removed
'   'WillC 11/7/00 SR3686 showed Label in error handler was wrong in SnapToGrid and EditCaption
'  'WillC 12/7/00 SR3688 & SR3689 Cutting and pasting erroneously allowed quotes in the caption.
'   TA 19/07/2000: HideCaption procedure updated as caption property is now a string
'   NCJ 16/10/00 - Ensure ClinicalTrialId and CRFPageId are reset when form unloads
'               (because they can still be there when the form next loads)
'   NCJ 23/10/00 SR3951 - Added linPageBreak line array to display Page Breaks
'   Mo  15/11/2000  ReplaceCharacters call replaced by call to VB Function Replace
'   TA 17/11/2000: Renumbering code moved here from frmCRFPageDefinition
'   NCJ 19/2/01 - Display current date in calendar control
'   ZA 03/08/01   SR2895 - eform selection and popup menu display
'   ZA 07/08/01   SR4232 - forcing of LostFocus event of txtEditCaption control
'   NCJ 27 Nov 01 - Added IMedCRFgroup Control and QuestionGroup picture boxes
' NCJ 7 Dec 01 - Tidied up handling of Question Groups & EForm Groups
' NCJ 11-13 Dec 01 - Replaced AlignLeft and AlignTop with AlignToGrid
'               Fixed bugs in DragOver
'               Made sure Renumber works with Question Groups
' NCJ 18 Dec 01 - Dropping groups onto the eForm
' NCJ 15 Jan 02 - Make sure all question lists are updated as required
' NCJ 27 Feb 02 - Fixed bug in clicking/double-clicking in captions within group
'               Don't allow adding of QGroups if they contain no questions
' NCJ 11 Mar 02 - Sorted out eForm update problem in EditEFG
' Ash 14/06/2002     Fix bug 2.2.15 no7
' ZA 25/07/2002 - Added functionality for font/colour of a question/caption.
' NCJ 15 Aug 02 - Sorted out null captions in NewCRFElement
' RS    20 Aug 2002 - Added MultiSelect (Lasso) Functionality
' RS     3 Sep 2002 - Changed Multiselect object to Private
' NCJ 4 Sep 02 - Sorted out eForm Group refreshments
' ZA 06/09/2002 - Added auto numbering for questions based on user selection
' MLM 16/09/02: Assign FieldOrder of 0 to form and visit dates.
' NCJ 2 Oct 02 - Added RQG horizontal scroll bar
' NCJ 7 Oct 02 - Removed references to unused txtCRFElement;
'           fixed bug in DragOver events (need to offset coords for items within question groups)
' RS 22 Oct 2002 -  Fixed bug in multiselect mode when CTRL click on LINE. Mouseup event of
'                   picCRFpage is now ignored when control key is pressed
' NCJ 4 Nov 02 - picScroller and picHeaders events mapped to picQuGroup events;
'               change to HScrollGroup;
'               "hand" mousepointer for RQGs (and correct mousepointers for other things).
' NCJ 6-7 Nov 02 - Added Hotlink handling
' RS 14 Jan 2003 - Added corrections for multiselect support
' NCJ 31 Jan 03 - Disallow Hotlink definition if Schedule is empty
' RS 13/02/2002: SR2750 InsertCRFPage: Check number of eForms so far: Do not allow > 255 eForms (VB limitation)
' RS 24/02/2003: Undid previous fix in optGraphical_DragOver and optCRFElement_DragOver which caused 'earthquake effects'
' NCJ 2 Apr 03 - Bug 1437 - Extra bIsGroupCaption argument to ChangeSelectedItem in ChangeSelectedElement
' TA 21/05/2003: Ensure font size positive and less than 24000 in UpdateFieldFont
' NCJ 29 May 03 - Added PasteBufferAvailable property (BUG 1815)
' NCJ 1 Jul 03 - Use new CRFElement Expandable property (Bug 1905)
' NCJ 11 Aug 03 - Added DeleteQuestion routine (Roche bug 1944)
' ic 04/07/2005 added clinical coding
'   TA  07/07/2005: seeting of Font.Charset = 1 to allow eastern european characters. CBD2591.
' NCJ 11 May - 22 Jun 06 - Multi-user stuff
' NCJ 6-13 Sept 06 - Fixing multi-user stuff
' NCJ 17 Apr 07 - Fixing MUSD things - Added mbSelectingEForm
'------------------------------------------------------------------------------------'

Option Explicit
Option Compare Binary
Option Base 0

Private mlClinicalTrialId As Long
Private mnVersionId As Integer
Private msClinicalTrialName As String

' NCJ 29 Nov 01 - The eForm object that will be looking after us
Private moEForm As EFormSD

Private mlCRFPageId As Long

'changed by Mo Morris 16/9/99 SR 1840
Private mnX As Single
Private mnY As Single ' save coordinates to assist in placement of drag-drop operation
Private msGridValue As String
Private mnGridSize As Integer   ' NCJ 15 Oct 99
Private mbDisplayGrid As Boolean   '   SDM 22/02/00 SR3074 For use with DisplayGrid
Private mnCRFElementId As Integer 'save selected field to change attributes
Private msMode As String
Private msCaption() As String

Private mbCaptionControlMove As Boolean
Private mblnMoveCaptionAndElement As Boolean
Private mblnDuringLoad As Boolean

Private msDefaultFontName As String
Private msDefaultFontSize As Single
Private msDefaultFontBold As Integer
Private msDefaultFontItalic As Integer
Private msDefaultFontColour As Long
Private msDefaultCRFColour As Long
Private mnPrevX As Single
Private mnPrevY As Single


'RS 20-08-2002: LASSO: Lasso/MultiSelect Worker Class
Private moMultiSelect As MultiSelect

Private Declare Function SetPixelV Lib "gdi32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
'Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDc As Long, ByVal nIndex As Long) As Long
'Private Declare Function LPtoDP Lib "gdi32" (ByVal hDc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function DPtoLP Lib "gdi32" (ByVal hDc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function GetMapMode Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hDc As Long, ByVal nMapMode As Long) As Long
'Private Const MM_TEXT = 1
Private Const MM_TWIPS = 6
Private mbEditBox_LostFocus As Boolean
Private mnPrevIndex As Integer

'za 24/07/2002 - stores a value of True if a user has clicked on a question caption
Private mbIsCaption As Boolean

' NCJ 11 May 06
Private meFormAccessMode As eSDAccessMode

' NCJ 17 Apr 07 - Remember when they're clicking on a new tab
Private mbSelectingEForm As Boolean

'---------------------------------------------------------------------
Public Property Let AccessMode(enMode As eSDAccessMode)
'---------------------------------------------------------------------
' NCJ 11 May 06 - Set the eForm access mode
'---------------------------------------------------------------------

    meFormAccessMode = enMode
    
End Property

'---------------------------------------------------------------------
Public Property Get AccessMode() As eSDAccessMode
'---------------------------------------------------------------------
' NCJ 11 May 06 - Get the eForm access mode
'---------------------------------------------------------------------

    AccessMode = meFormAccessMode

End Property

'---------------------------------------------------------------------
Public Sub RefreshMe(Optional lCRFPageId As Long = 0)
'---------------------------------------------------------------------
' Draw (or redraw) the eForm identified by lCRFPageId
' If lCRFPageId not given, use the current value
'---------------------------------------------------------------------
Dim lHscroll As Long
Dim lVscroll As Long

    ' Clear stuff before starting again
    Call TidyUpAndClearObjects

    If lCRFPageId <> mlCRFPageId And lCRFPageId <> 0 Then
       ' RS 6/9/2002: Removed prior paste method
        moMultiSelect.EndSelectMode
    End If

    ' NCJ 6 Jun 05 - Get eForm lock if necessary
    Call frmMenu.LockEForm(lCRFPageId)

    ' See if we're changing the CRFPageID
    If lCRFPageId > 0 Then
        mlCRFPageId = lCRFPageId
    End If
    
    If mlCRFPageId > 0 Then
        Set moEForm = New EFormSD
        
        ' RS 20-08-2002: Update MultiSelectClass
        moMultiSelect.Init Me, moEForm
        
        ' RS 14/01/2003: Save Scroll Position
        lHscroll = hsbCRFPage.Value
        lVscroll = vsbCRFPage.Value
        Call moEForm.InitAndBuild(Me, mlClinicalTrialId, mnVersionId, mlCRFPageId)
        ' RS 14/01/2003: Restore Scroll Position
        If lCRFPageId = 0 Then
            hsbCRFPage.Value = lHscroll
            vsbCRFPage.Value = lVscroll
        End If
        
    End If
    
End Sub

'--------------------------------------------------------------------------'
Public Function CRFElementById(nElementID As Integer) As CRFElement
'--------------------------------------------------------------------------'
' Return the CRFElement corresponding to this ID
' (simply pass the call on to the eForm object)
'--------------------------------------------------------------------------'

    Set CRFElementById = moEForm.CRFElementById(nElementID)
    
End Function

'--------------------------------------------------------------------------'
Public Function CRFElementByDataItemId(ByVal lDataItemId As Long) As CRFElement
'--------------------------------------------------------------------------'
' Return CRFElement corresponding to this DataItemId (if any)
' NB May be a group member
' Returns Nothing if question not on eForm
'--------------------------------------------------------------------------'

    Set CRFElementByDataItemId = moEForm.CRFElementByDataItemId(lDataItemId)

End Function

'--------------------------------------------------------------------------'
Public Property Get CurrentCRFElement() As CRFElement
'--------------------------------------------------------------------------'
' Return the currently selected CRFElement
' The ElementID is stored in mnCRFElementId variable
' Causes error if no currently selected element
'--------------------------------------------------------------------------'

    Set CurrentCRFElement = moEForm.CRFElementById(mnCRFElementId)
    
End Property

'---------------------------------------------------------------------
Public Function EFormGroupById(lQGroupId As Long) As EFormGroupSD
'---------------------------------------------------------------------
' Get the eFormGroup corresponding to this GroupId (if any)
' Returns Nothing if none exists
'---------------------------------------------------------------------

    Set EFormGroupById = moEForm.EFormGroups.EFormGroupById(lQGroupId)

End Function

'---------------------------------------------------------------------
Public Function GetNewEFormGroup(lQGroupId As Long) As EFormGroupSD
'---------------------------------------------------------------------
' Get a new eFormGroup for this GroupId
'---------------------------------------------------------------------

    Set GetNewEFormGroup = moEForm.EFormGroups.NewEFormGroup(lQGroupId, mlCRFPageId)

End Function

'---------------------------------------------------------------------
Private Sub AddGroupToCRF(oEFQG As EFormGroupSD, oQGroup As QuestionGroup)
'---------------------------------------------------------------------
' Add an eFormGroup to the eForm and draw it
' Assume mouse was clicked at gsgMouseDownOnFormX, gsgMouseDownOnFormY
'---------------------------------------------------------------------
Dim oElement As CRFElement
Dim sglCaptionX As Single
Dim sglCaptionY As Single
Dim vCaption As Variant

    On Error GoTo ErrLabel
    
    ' Create the caption
    If frmMenu.mnuOAutoCaption.Checked = True Then
        vCaption = oQGroup.QGroupName
        sglCaptionY = AlignToGrid(gsgMouseDownOnFormY)
        ' Shift the caption over to the left of the field
        sglCaptionX = AlignToGrid(gsgMouseDownOnFormX - CaptionWidth(CStr(vCaption)))
    Else
        ' RS 29/08/2002: Changed Null to "" as this causes crash when AutoCaption is off
        ' As instructed by NCJ
        ' vCaption = Null
        vCaption = ""
        sglCaptionX = 0
        sglCaptionY = 0
    End If
    
    ' Create the CRFElement structure
    Set oElement = moEForm.NewCRFGroup(oEFQG, _
                        AlignToGrid(gsgMouseDownOnFormX), AlignToGrid(gsgMouseDownOnFormY), _
                        sglCaptionX, sglCaptionY, vCaption)
    ' Now build it
    Call moEForm.BuildNewCRFElement(oElement)
    
    ' Ensure eForm updating happens
    Call RefreshMe
    
    Set oElement = Nothing

Exit Sub
ErrLabel:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, "AddGroupToCRF", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select

End Sub

'---------------------------------------------------------------------
Public Sub DropEFormGroupOnCRF(ByVal lQGroupId As Long)
'---------------------------------------------------------------------
' Put a new eFormGroup on the eForm
' This does the checks to see if it's allowed before creating it,
' and also displays the edit eForm dialog first
' Assume mouse was clicked at coordinates gsgMouseDownOnFormX, gsgMouseDownOnFormY
'---------------------------------------------------------------------
Dim oQGroup As QuestionGroup
Dim oEFormGroup As EFormGroupSD
Dim bAddGroup As Boolean

    Set oQGroup = frmMenu.QuestionGroups.GroupById(lQGroupId)
    'Check to see if EFormGroup already exists
    Set oEFormGroup = EFormGroupById(lQGroupId)
    
    If oEFormGroup Is Nothing Then  ' Group not already on eForm
        bAddGroup = True
        ' NCJ 27 Feb 02 - Warn about a group with no questions
        If oQGroup.Questions.Count = 0 Then
            bAddGroup = (DialogQuestion("This group '" & oQGroup.QGroupCode & "' contains no questions." & vbCrLf _
                & "Are you sure you wish to add it to this eForm?") = vbYes)
        End If
        
        If bAddGroup Then
            ' Check they're allowed to add this group
            If Not CanAddQGroup(oQGroup) Then
                DialogInformation "This group cannot be added to this eForm " _
                    & "because it contains a question that already exists on the eForm"
            Else
                ' Create a new eFormgroup object
                Set oEFormGroup = GetNewEFormGroup(lQGroupId)
                ' Let the user edit the details
                If frmEFormQGroupDefinition.Display(oEFormGroup, oQGroup, True) Then
                    ' If they clicked OK, display it on the eForm
                    Call AddGroupToCRF(oEFormGroup, oQGroup)
                    ' NCJ 19 Jun 06 - Mark study as changed
                    Call frmMenu.MarkStudyAsChanged

                    ' NCJ 3 Jan 02 - Refresh Data List
                    ' NCj 15 Jan 02 - Refresh ALL data lists
                    Call frmMenu.RefreshQuestionLists(Me.ClinicalTrialId)
                Else
                    ' They clicked Cancel so delete the eForm group (it's unused)
                    Call moEForm.EFormGroups.Delete(oEFormGroup.QGroupID)
                End If
            End If
        End If
    Else
        ' They can't insert the same group twice
        DialogInformation "This group has already been used on this eForm"
    End If

End Sub

'---------------------------------------------------------------------
Public Property Get ClinicalTrialId() As Long
'---------------------------------------------------------------------

    ClinicalTrialId = mlClinicalTrialId

End Property

'---------------------------------------------------------------------
Public Property Let ClinicalTrialId(ByVal vClinicalTrialId As Long)
'---------------------------------------------------------------------

    mlClinicalTrialId = vClinicalTrialId

End Property

'---------------------------------------------------------------------
Public Property Get VersionId() As Integer
'---------------------------------------------------------------------

    VersionId = mnVersionId

End Property

'---------------------------------------------------------------------
Public Property Let VersionId(ByVal vVersionId As Integer)
'---------------------------------------------------------------------

    mnVersionId = vVersionId

End Property

'---------------------------------------------------------------------
Public Property Get ClinicalTrialName() As String
'---------------------------------------------------------------------

    ClinicalTrialName = msClinicalTrialName

End Property

'---------------------------------------------------------------------
Public Property Let ClinicalTrialName(ByVal vClinicalTrialName As String)
'---------------------------------------------------------------------

    msClinicalTrialName = vClinicalTrialName

End Property

'---------------------------------------------------------------------
Public Property Get CRFPageId() As Long
'---------------------------------------------------------------------

    CRFPageId = mlCRFPageId

End Property

'---------------------------------------------------------------------
Public Property Let CRFPageId(tmpX As Long)
'---------------------------------------------------------------------

    mlCRFPageId = tmpX

End Property

'---------------------------------------------------------------------
Public Property Get CRFelementID() As Integer
'---------------------------------------------------------------------

    CRFelementID = mnCRFElementId

End Property

'---------------------------------------------------------------------
Public Property Let CRFelementID(tmpX As Integer)
'---------------------------------------------------------------------

    mnCRFElementId = tmpX

End Property

'---------------------------------------------------------------------
Public Property Get Grid() As String
'---------------------------------------------------------------------
' Returned as "Large", "Medium", "Small" or "None"
'---------------------------------------------------------------------

    Grid = msGridValue
    
End Property

'---------------------------------------------------------------------
Public Property Let Grid(sGrid As String)
'---------------------------------------------------------------------
' Set grid property and set actual grid size
'   SDM 22/02/00 SR3074 Grid sizes changed to be divisable by Screen Twips
'---------------------------------------------------------------------

    msGridValue = sGrid
    ' Change the grid size accordingly NCJ 15 Oct 99
    Select Case sGrid
    Case "None"
        mnGridSize = 1
    Case "Small"
        mnGridSize = Screen.TwipsPerPixelX * 3  'SDM 22/02/00 SR3074
'        mnGridSize = 50
    Case "Medium"
        mnGridSize = Screen.TwipsPerPixelX * 6  'SDM 22/02/00 SR3074
        'mnGridSize = 100
    Case "Large"
        mnGridSize = Screen.TwipsPerPixelX * 12 'SDM 22/02/00 SR3074
'        mnGridSize = 200
    End Select
    
   ' SaveSetting "ImedMACRO", "Options", "GridSize", sGrid


End Property

'---------------------------------------------------------------------
Public Property Get CombinedCaptionControlMove() As Boolean
'---------------------------------------------------------------------
'To Toggle the CombinedCaptionControlMove on/off depending on registry setting
'---------------------------------------------------------------------

    CombinedCaptionControlMove = mbCaptionControlMove

End Property

'---------------------------------------------------------------------
Public Property Get GridSize() As Integer
'---------------------------------------------------------------------
' Return actual gridsize as integer
' NCJ 15 Oct 99
'---------------------------------------------------------------------

    GridSize = mnGridSize
    
End Property

'---------------------------------------------------------------------
Public Property Get UpdateMode() As String
'---------------------------------------------------------------------

    UpdateMode = msMode

End Property

'---------------------------------------------------------------------
Public Property Let UpdateMode(tmpMode As String)
'---------------------------------------------------------------------

    msMode = tmpMode

End Property

'---------------------------------------------------------------------
Public Property Get DefaultFontName() As String
'---------------------------------------------------------------------

    DefaultFontName = msDefaultFontName

End Property

'---------------------------------------------------------------------
Public Property Let DefaultFontName(tmpDefaultFontName As String)
'---------------------------------------------------------------------

    msDefaultFontName = tmpDefaultFontName

End Property

'---------------------------------------------------------------------
Public Property Get DefaultFontSize() As Single
'---------------------------------------------------------------------

    DefaultFontSize = msDefaultFontSize

End Property

'---------------------------------------------------------------------
Public Property Let DefaultFontSize(tmpDefaultFontSize As Single)
'---------------------------------------------------------------------

    msDefaultFontSize = tmpDefaultFontSize

End Property

'---------------------------------------------------------------------
Public Property Get DefaultFontBold() As Integer
'---------------------------------------------------------------------

    DefaultFontBold = msDefaultFontBold

End Property

'---------------------------------------------------------------------
Public Property Let DefaultFontBold(tmpDefaultFontBold As Integer)
'---------------------------------------------------------------------

    msDefaultFontBold = tmpDefaultFontBold

End Property

'---------------------------------------------------------------------
Public Property Get DefaultFontItalic() As Integer
'---------------------------------------------------------------------

    DefaultFontItalic = msDefaultFontItalic

End Property

'---------------------------------------------------------------------
Public Property Let DefaultFontItalic(tmpDefaultFontItalic As Integer)
'---------------------------------------------------------------------

    msDefaultFontItalic = tmpDefaultFontItalic

End Property

'---------------------------------------------------------------------
Public Property Get DefaultFontColour() As Long
'---------------------------------------------------------------------

    DefaultFontColour = msDefaultFontColour

End Property

'---------------------------------------------------------------------
Public Property Let DefaultFontColour(tmpDefaultFontColour As Long)
'---------------------------------------------------------------------

    msDefaultFontColour = tmpDefaultFontColour

End Property

'---------------------------------------------------------------------
Public Property Get DefaultCRFColour() As Long
'---------------------------------------------------------------------

    DefaultCRFColour = msDefaultCRFColour

End Property

'---------------------------------------------------------------------
Public Property Let DefaultCRFColour(tmpDefaultCRFColour As Long)
'---------------------------------------------------------------------

    msDefaultCRFColour = tmpDefaultCRFColour

End Property


'---------------------------------------------------------------------
Public Sub DisplayCRFPage(lCRFPageId As Long)
'---------------------------------------------------------------------
' Display the CRFPage with the given ID
' sCRFPageID may be 0 - if so, ignore errors
'---------------------------------------------------------------------

    On Error GoTo InvalidPageKey
    
    If lCRFPageId = 0 Then
        ' Need to see if we're just re-displaying the current eForm
        If mlCRFPageId > 0 Then
            Call frmMenu.LockEForm(mlCRFPageId)
        End If
    Else
        tabCRF.Tabs(gsCRF_PAGE_LABEL & lCRFPageId).Selected = True
    End If
    
'    tabCRF.Tabs(gsCRF_PAGE_LABEL & lCRFPageId).Selected = True
    
InvalidPageKey:
    ' Do nothing

End Sub

'---------------------------------------------------------------------
Public Sub RefreshCRF()
'---------------------------------------------------------------------
' Rebuild the tabstrips and redraw the eForm area
' NCJ 17 Apr 07 - Changed tabCRF_Click to TabCRFClick
'---------------------------------------------------------------------
Dim rsCRFPageList As ADODB.Recordset
Dim oTabStrip As TabStrip
Dim bFormAlphabeticOrder As Boolean

    On Error GoTo ErrHandler
    
    ' Initialise tabstrip
    Set oTabStrip = tabCRF
    Set oTabStrip.ImageList = imglistSmallIcons
    oTabStrip.Tabs.Clear
    oTabStrip.TabWidthStyle = tabNonJustified
    
    'specific SQL statement replaced by a call to gdsCRFPageList, which returns
    'a recordset that is either sorted by CRFPageOrder or CRFTitle
    bFormAlphabeticOrder = IsFormDisplayOrderAlphabetic
    Set rsCRFPageList = New ADODB.Recordset
    Set rsCRFPageList = gdsCRFPageList(ClinicalTrialId, _
                              VersionId, bFormAlphabeticOrder)

    ' Add a tab for each CRF page
    '   ATN 29/4/99 SR 836
    '   Replace & with && because single & is used for keyboard shortcuts
    Do While Not rsCRFPageList.EOF
        oTabStrip.Tabs.Add , gsCRF_PAGE_LABEL & CStr(rsCRFPageList!CRFPageId), _
                            Replace(rsCRFPageList!CRFTitle, "&", "&&")
        oTabStrip.Tabs(gsCRF_PAGE_LABEL & CStr(rsCRFPageList!CRFPageId)).Image = gsCRF_PAGE_LABEL
        rsCRFPageList.MoveNext
    Loop
    
    rsCRFPageList.Close
    Set rsCRFPageList = Nothing
    
    If oTabStrip.Tabs.Count > 0 Then
        Me.picCRFPage.Visible = True
        'Ash 14/06/2002 'To fix bug 2.2.15 no7
        'Ash 18/06/2002 moved from previous place (below)
        ' because could not add new eform to new studies.
        ' NCJ 12 Sept 06 - Don't do click if we're not visible
        ' NCJ 17 Apr 07 - Changed to TabCRFClick
'        If Me.Visible Then Call tabCRF_Click
        If Me.Visible Then Call TabCRFClick
        Me.picCRFPage.Visible = True
    Else
        Me.picCRFPage.Visible = False
    End If
    
   
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "RefreshCRF")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    

End Sub

'---------------------------------------------------------------------
Public Sub RebuildQuestionIfVisible(ByVal lDataItemId As Long)
'---------------------------------------------------------------------
' Rebuild data item on form if it is visible on current CRF page
' after its definition has been edited
' NCJ 14 Feb 03 - Make sure it works for Group questions
'---------------------------------------------------------------------
Dim oCRFElement As CRFElement
Dim oElement As CRFElement

    If Me.picCRFPage.Visible = False Then
        Exit Sub
    End If
    
    ' Check we have a valid CRFElements collection
    If Not moEForm.CRFElements Is Nothing Then
        For Each oCRFElement In moEForm.CRFElements
            If oCRFElement.DataItemId = lDataItemId Then
                ' NCJ 14 Feb 03
                Call moEForm.ReBuildCRFElement(oCRFElement, True, True)
'                ' Convert to "owner" element if necessary
'                Set oElement = moEForm.OwnerCRFElement(oCRFElement)
'                ' Remove element with old definition first
'                Call moEForm.RemoveCRFElement(oElement)
'                ' Then rebuild using new definition
'                Call moEForm.ReBuildCRFElement(oElement, True, False)
                Exit For
            End If
        Next
    End If

    Set oCRFElement = Nothing
    Set oElement = Nothing

End Sub

'---------------------------------------------------------------------
Private Sub SnapToGrid(Source As Control, X As Single, Y As Single)
'---------------------------------------------------------------------
' This is called while a control is being dragged around on the eForm
' and ensures that all the bits and pieces of a data entry field stick together.
' We also make sure that captions and elements stick together
' if combined caption/control movement is required,
' and that we always "snap" to the grid if it's on
' X and Y are the current mouse coordinates on the picCRFPage
' Source is the control that's being dragged
' NB mnX and mnY were stored as the local (Source) coords of where the drag started
' ic 14/06/2005 added clinical coding
'---------------------------------------------------------------------
Dim sglMoveLeft As Single
Dim sglMoveUp As Single
Dim oElement As CRFElement
Dim nIndex As Integer

    On Error GoTo ErrLabel
    
    ' We don't want to move the tree view around!
    If Source.Name = "trvDataList" Then Exit Sub
    
    ' If no CRFElement has been selected, get out quick
    If mnCRFElementId = 0 Then Exit Sub
    
    Set oElement = moEForm.CRFElementById(mnCRFElementId)
    ' NCJ 28 Nov 01 - Don't move group items (for AutoFit groups)
    If oElement.OwnerQGroupID > 0 Then Exit Sub
    
    If oElement.ControlType = gn_LINE Then
        ' Ignore the X move for a Line
        Source.Move 0, AlignToGrid(Y - mnY)
        ' That's all we do for a line
        Exit Sub
    End If
        
    ' Calculate how much to move up and left
    ' mnX and mnY are the starting mouse coords IN THE SOURCE CONTROL
    ' x and y are the mouse coords IN THE CRF PAGE
    ' i.e. Source.Left + mnX is the start coord in the CRFPage
    sglMoveLeft = Source.Left + AlignToGrid(mnX - X)
    sglMoveUp = Source.Top + AlignToGrid(mnY - Y)
    
    ' First move the actual Source control to its new (offset) coordinates
    Source.Move AlignToGrid(X - mnX), AlignToGrid(Y - mnY)

     ' Pick up the control array index for this element (assume row 1)
    nIndex = moEForm.IndexForElementRow(oElement, 1)
    
   Select Case Source.Name
    Case "lblCaption"
        ' Source is lblCaption
        If mblnMoveCaptionAndElement Then
            ' We have to move the associated control by the same amount,
            ' aligning the new coordinates to the Grid
            Select Case oElement.ControlType
            Case gn_OPTION_BUTTONS
                optCRFGroup(nIndex).Move optCRFGroup(nIndex).Left - sglMoveLeft, _
                                         optCRFGroup(nIndex).Top - sglMoveUp
            Case gn_PUSH_BUTTONS
                picCRFElement(nIndex).Move picCRFElement(nIndex).Left - sglMoveLeft, _
                                           picCRFElement(nIndex).Top - sglMoveUp
            Case gn_TEXT_BOX, gn_CALENDAR, gn_POPUP_LIST
                lblCRFElement(nIndex).Move lblCRFElement(nIndex).Left - sglMoveLeft, _
                                           lblCRFElement(nIndex).Top - sglMoveUp
                ' For Calendar and PopupList, move their command buttons too
                'ic 14/06/2005 clinical coding: move browse button
                If oElement.DataItemType = eDataType.Thesaurus Then
                    cmdDictionary(nIndex).Left = lblCRFElement(nIndex).Left + lblCRFElement(nIndex).Width
                    cmdDictionary(nIndex).Top = lblCRFElement(nIndex).Top
                End If
                If oElement.Expandable Then
                    cmdExpand(nIndex).Left = lblCRFElement(nIndex).Left + lblCRFElement(nIndex).Width
                    cmdExpand(nIndex).Top = lblCRFElement(nIndex).Top
                End If
                If oElement.ControlType = gn_CALENDAR Then
                    cmdCalendar(nIndex).Left = lblCRFElement(nIndex).Left + lblCRFElement(nIndex).Width
                    cmdCalendar(nIndex).Top = lblCRFElement(nIndex).Top
                End If
                If oElement.ControlType = gn_POPUP_LIST Then
                    cmdPopupList(nIndex).Left = lblCRFElement(nIndex).Left + lblCRFElement(nIndex).Width
                    cmdPopupList(nIndex).Top = lblCRFElement(nIndex).Top
                End If
            Case gn_ATTACHMENT
                cmdAttachment(nIndex).Move cmdAttachment(nIndex).Left - sglMoveLeft, _
                                           cmdAttachment(nIndex).Top - sglMoveUp
            End Select
            ' Move the group box too
            If oElement.QGroupID > 0 Then
                picQuGroup(nIndex).Move picQuGroup(nIndex).Left - sglMoveLeft, _
                                        picQuGroup(nIndex).Top - sglMoveUp
            End If
        End If
    
    Case Else
        ' Not lblCaption (so it's a field of some sort)
        ' Just bring associated paraphernalia with the field
        ' and move the associated label if appropriate
        Select Case oElement.ControlType
        Case gn_TEXT_BOX
            'ic 14/06/2005 clinical coding: move browse button
            If oElement.DataItemType = eDataType.Thesaurus Then
                cmdDictionary(nIndex).Left = lblCRFElement(nIndex).Left + lblCRFElement(nIndex).Width
                cmdDictionary(nIndex).Top = lblCRFElement(nIndex).Top
            End If
            If oElement.Expandable Then
                cmdExpand(nIndex).Left = lblCRFElement(nIndex).Left + lblCRFElement(nIndex).Width
                cmdExpand(nIndex).Top = lblCRFElement(nIndex).Top
            End If
        Case gn_CALENDAR
            ' Keep the calendar button to the right of the label field
            cmdCalendar(nIndex).Left = Source.Left + Source.Width
            cmdCalendar(nIndex).Top = Source.Top
        Case gn_POPUP_LIST
            ' Keep the popuplist button to the right of the label field
            cmdPopupList(nIndex).Left = Source.Left + Source.Width
            cmdPopupList(nIndex).Top = Source.Top
        End Select
        
        ' Are we moving captions & controls together?
        If mblnMoveCaptionAndElement Then
            ' Move the caption by the same amount
            lblCaption(nIndex).Left = lblCaption(nIndex).Left - sglMoveLeft
            lblCaption(nIndex).Top = lblCaption(nIndex).Top - sglMoveUp
        End If
        
    End Select
    
    Set oElement = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description _
                        & "|frmCRFDesign.SnapToGrid(" & Source.Name & ")"
    
End Sub

'---------------------------------------------------------------------
Private Sub SelectCRF(lCRFPageId As Long)
'---------------------------------------------------------------------
' NCJ 20 Jun 06 Try to select this eForm on the tabs
' If the eForm doesn't exist, we select the first one (if any)
' NCJ 27 Jun 06 -  Bug fixing!
'---------------------------------------------------------------------
Dim i As Integer
Dim bFoundIt As Boolean

    ' Do nothing if no tabs!
    If tabCRF.Tabs.Count < 1 Then Exit Sub
    
    bFoundIt = False
    For i = 1 To tabCRF.Tabs.Count
        If CLng(Mid(tabCRF.Tabs(i).Key, Len(gsCRF_PAGE_LABEL) + 1)) = lCRFPageId Then
            ' We have the one we want
            Set tabCRF.SelectedItem = tabCRF.Tabs(i)
            bFoundIt = True
            Exit For
        End If
    Next
    If Not bFoundIt Then
        ' Default to first eForm
        Set tabCRF.SelectedItem = tabCRF.Tabs(1)
    End If
    
End Sub

'---------------------------------------------------------------------
Public Sub InsertCRFPage()
'---------------------------------------------------------------------
'Create a new CRF Page
'Adjusted for CLM calls - NCJ 10/8/99
'Mo Morris  14/2/00     sections re-written
'---------------------------------------------------------------------
Dim sCRFPageCode As String
Dim lCRFPageId As Long
Dim sSQL As String
Dim sMSG As String
Dim enAutoNumbering As eAutoNumbering

    On Error GoTo ErrHandler
    
    ' RS 13/02/2002: Check number of eForms so far: Do not allow > 255 eForms (VB limitation)
    If Me.tabCRF.Tabs.Count >= 255 Then
        MsgBox "You can not create more than 255 eForms in a study", vbOKOnly + vbInformation, "MACRO"
        Exit Sub
    End If
        
    'TA 28/03/2000 - call new function to get code
    sCRFPageCode = GetItemCode(gsITEM_TYPE_EFORM, "New " & gsITEM_TYPE_EFORM & " code:")
    If sCRFPageCode = "" Then    ' if cancel, then return control to user
        Exit Sub
    End If
    
    'ZA 06/09/2002 - check if the auto numbering is on or off
    If frmMenu.mnuAutomaticNumbering.Checked Then
        enAutoNumbering = eAutoNumbering.NumberingOn
    Else
        enAutoNumbering = eAutoNumbering.NumberingOff
    End If
    
    'Insert CRF page details
    'Begin transaction
    TransBegin

    ' Create the CRFPage Arezzo task - NCJ 10/8/99
    lCRFPageId = gnNewCLMPlan(gsCLMCRFName(sCRFPageCode))
    
    'Changed by Mo Morris 6/8/99
    'CRFPageLabel,LocalCRFPageLabel,SequentialEntry,CRFPageDateLabel,DisplayNumbers added
    'ZA 10/08/01 added eFormDatePrompt field into the statement
    'Mo Morris 30/8/01 Db Audit (HideIfInactive added for default value of 0)
    sSQL = "INSERT INTO CRFPage " _
        & " ( ClinicalTrialId, VersionId, CRFPageId, CRFPageCode, " _
        & " CRFTitle, BackgroundColour, CRFPageOrder, " _
        & "CopiedFromClinicalTrialId, CopiedFromVersionId, CopiedFromCRFPageId,CRFPageLabel, " _
        & "LocalCRFPageLabel, SequentialEntry, CRFPageDateLabel, DisplayNumbers, eFormDatePrompt, HideIfInactive)" _
        & " VALUES (" & ClinicalTrialId & "," _
        & VersionId & "," & lCRFPageId & ",'" & sCRFPageCode & "','" _
        & sCRFPageCode & "',0," & mnNextCRFPageOrder(ClinicalTrialId, VersionId) & ",NULL,NULL,NULL,'',0,0,''," & enAutoNumbering & ",0,0)"
            
    MacroADODBConnection.Execute sSQL, dbSQLPassThrough + dbFailOnError

    ' This now completes the Arezzo side of things - NCJ 10/8/99
    InsertProformaCRFPage ClinicalTrialId, lCRFPageId, sCRFPageCode

    'End transaction
    TransCommit

    ' NCJ 19 Jun 06 - Mark as changed
    Call frmMenu.MarkStudyAsChanged

    ' NCJ 10 Jan 02 - Ensure ALL question lists are updated
    If frmMenu.mnuVDataCRFPage.Checked = True Then
        Call frmMenu.RefreshQuestionLists(Me.ClinicalTrialId)
    End If

    If Not frmMenu.gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmStudyVisits") Is Nothing Then
        frmStudyVisits.RefreshStudyVisits
    End If
    
    Me.RefreshCRF
    'Me.Refresh
    Set Me.tabCRF.SelectedItem = Me.tabCRF.Tabs(Me.tabCRF.Tabs.Count)
    
    
Exit Sub
ErrHandler:
    'RollBack transaction
    TransRollBack
    
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "InsertCRFPage")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Public Sub DeleteCRFElement()
'---------------------------------------------------------------------
' Remove currently selecte CRF element from the eForm
' NCJ 29 Nov 01 - This now uses new eForm object to do the whole thing
' NB If the currently selected element is a group member,
' we delete the owning group and all its components.
' If it's a group, we delete the group and all its elements
' REM 29/01/02 - added RefreshQuestionLists if question group has been removed from an EForm
' REM 06/03/02 - Added check for OwnerQGroupId for Removing a Data Item
'---------------------------------------------------------------------
Dim sMSG As String
Dim oElement As CRFElement
Dim oForm As Form

    On Error GoTo ErrHandler
    
    Set oElement = moEForm.CRFElementById(mnCRFElementId)
    
    If frmMenu.TrialStatus > eTrialStatus.InPreparation _
    And (oElement.DataItemId > 0 Or oElement.QGroupID > 0) Then
        sMSG = "You may not remove questions from an eForm once a study has been opened. " & vbCrLf
        sMSG = sMSG & "You may hide questions by editing the form definition."
        Call DialogInformation(sMSG)
        Exit Sub
    End If
    
    ' NCJ 30 Nov 01 - Let the eForm object handle it
    Call moEForm.DeleteCRFElement(oElement)
    
    'REM 06/03/02 - Added OwnerQGroup check so that when removing a QGroup from an Eform, if the user clicks on a question
    ' in the group it will not try and delete it
    If (oElement.DataItemId > 0) And (oElement.OwnerQGroupID = 0) Then    'only for data elements
    
        ' NCJ 15 Jan 02 - Must update ALL Question Lists for this study
        For Each oForm In Forms
            If oForm.Name = "frmDataList" Then
                If oForm.ClinicalTrialId = Me.ClinicalTrialId Then
                    Call oForm.RemoveCRFFromDataItem(oElement.DataItemId, mlCRFPageId, mnCRFElementId)
                End If
            End If
        Next

    ElseIf (oElement.QGroupID > 0) Or (oElement.OwnerQGroupID > 0) Then
        'REM 29/01/02
        'refresh question list if question group has been removed from an EForm
        frmMenu.RefreshQuestionLists (Me.ClinicalTrialId)
    
    End If
    
    ' PN 26/09/99
    ' change the display text to reflect the deletion
    Call frmMenu.ChangeSelectedItem(vbNullString, vbNullString)
    
    ' NCJ 19 Jun 06 - Mark as changed
    Call frmMenu.MarkStudyAsChanged

    Set oElement = Nothing

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "DeleteCRFElement")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
        
End Sub

'---------------------------------------------------------------------
Public Sub Delete()
'---------------------------------------------------------------------
' Delete the current CRFPage
'---------------------------------------------------------------------
 
    Call AttemptDeleteCRFPage(UpdateMode, tabCRF.SelectedItem, _
                        ClinicalTrialId, VersionId, CRFPageId)

End Sub

'---------------------------------------------------------------------
Public Sub RefreshCRFPageList(Optional lCRFPageId As Long = 0)
'---------------------------------------------------------------------
' Refresh the eForm tabs and select the first one (if any)
' NCJ 7 Sept 06 - If lCRFPageID > 0, try to select that tab
' NCJ 17 Apr 07 - Don't interfere with an ongoing eForm Tab selection process
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    Call Me.RefreshCRF
    
    If Me.tabCRF.Tabs.Count > 0 Then
        ' NCJ 17 Apr 07 - Don't interfere with an ongoing selection process
        If Not mbSelectingEForm Then
            If lCRFPageId > 0 Then
                Call SelectCRF(lCRFPageId)
            Else
                ' Select the first one
                Me.tabCRF.SelectedItem = Me.tabCRF.Tabs(1)
            End If
        End If
    Else
        Call moEForm.ClearEForm
        SetPageSize Me
        frmMenu.ChangeSelectedItem "", ""
        'za 20/05/2002, set the pageid to 0, upon deletion of eForm
        Me.CRFPageId = 0
    End If
        
    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "RefreshCRFPageList")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Public Sub RebuildCurrentElement()
'---------------------------------------------------------------------
' Rebuild the current element (assume it's been changed)
' Automatically rebuilds the owning group
' if current element is a group member
'---------------------------------------------------------------------

    Call moEForm.ReBuildCRFElement(CRFElementById(mnCRFElementId), True, True)
    
End Sub

'---------------------------------------------------------------------
Public Sub UpdateFieldInputTool(nControlType As Integer)
'---------------------------------------------------------------------
' Change the Control type of a question
' NCJ 4 Sept 02 - If a group question, refresh whole eForm
' NCJ 7 Sept 06 - Mark study as changed
'---------------------------------------------------------------------
Dim oElement As CRFElement
Dim oOwnerElement As CRFElement

    On Error GoTo ErrHandler
    
    If mnCRFElementId = 0 Then Exit Sub
    
    Set oElement = moEForm.CRFElementById(mnCRFElementId)
    ' Get the Owner element if a group member)
    Set oOwnerElement = moEForm.OwnerCRFElement(oElement)
    
    ' Must remove old one before changing its control type
    ' (otherwise it won't know what controls to remove)
    Call moEForm.RemoveCRFElement(oOwnerElement)
    
    ' Set the control type on this one
    Call oElement.SetControlType(nControlType)
    
    If oElement.OwnerQGroupID = 0 Then
        ' Now rebuild but don't remove previous because it's already gone
        Call moEForm.ReBuildCRFElement(oOwnerElement, False, False)
    Else
        ' NCJ 4 Sept 02
        Call RefreshEFormGroup(oElement.OwnerQGroupID)
    End If
    
    frmMenu.ChangeSelectedItem gsCRF_ELEMENT_LABEL & Str(mnCRFElementId), _
        " " & frmMenu.lblSelectedItem.Caption
    
    Set oElement = Nothing
    Set oOwnerElement = Nothing

    ' NCJ 7 Sept 06
    Call frmMenu.MarkStudyAsChanged
    
Exit Sub
ErrHandler:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, "UpdateFieldInputTool", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select
    
End Sub

'---------------------------------------------------------------------
Public Sub EditEFG()
'---------------------------------------------------------------------
' NCJ 3 Dec 01
' Edit an eFormGroup (assume mnCRFElementID set up)
' The element may be a group or a group member
'---------------------------------------------------------------------
Dim oElement As CRFElement
Dim oEFG As EFormGroupSD

    On Error GoTo ErrLabel
    
    ' Pick up the selected CRFElement
    Set oElement = CRFElementById(mnCRFElementId)
    If oElement.OwnerQGroupID > 0 Then
        ' Pick up the owner group definition
        Set oEFG = moEForm.EFormGroups.EFormGroupById(oElement.OwnerQGroupID)
        ' Reset oElement to be the owning group
        Set oElement = CRFElementById(oEFG.ElementID)
    Else
        ' Assume the element is the group itself
        Set oEFG = moEForm.EFormGroups.EFormGroupById(oElement.QGroupID)
    End If
    ' Let the user edit the details
    Call frmEFormQGroupDefinition.Display(oEFG, _
            frmMenu.QuestionGroups.GroupById(oElement.QGroupID), (meFormAccessMode > sdReadOnly))
    
    ' NCJ 4 Sept 02 - Call RefreshEFormGroup anyway because sometimes VB splats it
    ' even if we didn't do anything...
    Call RefreshEFormGroup(oElement.QGroupID)
    
    Set oElement = Nothing
    Set oEFG = Nothing

Exit Sub
ErrLabel:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, "EditEFG", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select

End Sub

'---------------------------------------------------------------------
Public Sub EditCaption()
'---------------------------------------------------------------------
'Program changed by Mo Morris 25/9/98 (SPR 430)
'EditCaption has been re-written and now forms an editing area of at least 2 lines
'or 1 line more than the comment/caption being edited
'Note that this routine is used to edit captions and comments as well as creating
'new comments via its call from picCRFPage_MouseDown (Create a Comment)
'---------------------------------------------------------------------
Dim oElement As CRFElement
Dim sglWidth As Single
Dim sglHeight As Single
Dim sglTop As Single
Dim sglLeft As Single
Dim sglLineHeight As Single
Dim oCaptionLabel As Label
Dim nIndex As Integer

    On Error GoTo ErrHandler
    
    Set oElement = moEForm.CRFElementById(mnCRFElementId)
    
    If oElement.ControlType = gn_HOTLINK Then
        Call EditHotlink(oElement)
        Exit Sub
    End If
    
    If oElement.OwnerQGroupID > 0 Then
        ' For group items, we use row 0 for the header labels
        nIndex = moEForm.IndexForElementRow(oElement, 0)
        Set oCaptionLabel = lblGroupHeader(nIndex)
        ' Offset coordinates by question group container
        ' (Labels sit inside picHeaders inside picFixed inside picQuGroup)
        Call moEForm.GetControlTopLeft(oCaptionLabel, sglTop, sglLeft)
    Else
        nIndex = moEForm.IndexForElementRow(oElement, 1)
        Set oCaptionLabel = lblCaption(nIndex)
        sglLeft = oCaptionLabel.Left
        sglTop = oCaptionLabel.Top
    End If
    
    sglTop = sglTop + Me.Top + picCRFPage.Top + frmCRFPage.Top + frmMenu.Height - frmMenu.ScaleHeight
    sglLeft = sglLeft + Me.Left + picCRFPage.Left
    sglTop = sglTop + frmMenu.Top
    sglLeft = sglLeft + frmMenu.Left
    
    'WillC 12/7/00 SR3688 & SR3689 Cutting and pasting erroneously allowed quotes in the caption.
    'put original text in tag for later use
'    txtEditCaption.Tag = txtEditCaption.Text

    If frmEditCaption.Display(oCaptionLabel, oElement.Caption, sglLeft, sglTop) Then
        'the user clicked OK
        ' If it's a comment being edited and it's an empty string then delete it
        If oCaptionLabel.Caption = "" And oElement.DataItemId = 0 And oElement.QGroupID = 0 Then
            Call moEForm.DeleteCRFElement(oElement)
        Else
            Call oElement.SetCaption(oCaptionLabel.Caption)
            'Call moEForm.ReBuildCRFElement(oElement, False, True)
        End If
        'MLM 05/06/03: 3.0 buglist 1657: Moved Me.RefreshMe outside the above If.
        Me.RefreshMe
    ElseIf oElement.Caption = "" And oElement.DataItemId = 0 And oElement.QGroupID = 0 Then
        'the user clicked Cancel; delete if it's a comment that was empty to start with
        Call moEForm.DeleteCRFElement(oElement)
    End If
    
    Set oElement = Nothing

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "EditCaption")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Public Sub HideCaption()
'---------------------------------------------------------------------
' Hide the caption of a CRFElement
'---------------------------------------------------------------------
Dim oElement As CRFElement

    On Error GoTo ErrHandler
    
    Set oElement = moEForm.CRFElementById(mnCRFElementId)
    
    Call oElement.SetCaption("")
        
    Call moEForm.ReBuildCRFElement(oElement, False, True)
    
    Set oElement = Nothing
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "HideCaption")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Public Sub ShowCRFPageDefinition()
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
 
    Call frmCRFPageDefinition.Display(ClinicalTrialId, VersionId, _
                            ClinicalTrialName, mlCRFPageId, meFormAccessMode)
    
    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "ShowCRFPageDefinition")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
    
End Sub

'---------------------------------------------------------------------
Public Function EditHotlink(oElement As CRFElement) As Boolean
'---------------------------------------------------------------------
' Edit the properties of a hotlink
' Returns TRUE if they did anything
' NCJ 31 Jan 03 - Don't let them do it if the Schedule is empty
'---------------------------------------------------------------------
Dim sCaption As String
Dim sHotlink As String
Dim bDidSomething As Boolean

    bDidSomething = False
    
    ' NCJ 31 Jan 03 - Don't allow a Hotlink if nothing in the Schedule
    If Not frmMenu.ScheduleOK(False) Then
        DialogInformation "You cannot define a hotlink while the study's Schedule is empty"
    Else
    
        sCaption = oElement.Caption
        sHotlink = oElement.Hotlink
        
        If frmHotlink.Display(ClinicalTrialId, sCaption, sHotlink, (Me.AccessMode > sdReadOnly)) Then
            ' Save it if it's changed
            If (sCaption <> oElement.Caption) Or (sHotlink <> oElement.Hotlink) Then
                Call oElement.SetHotlink(sCaption, sHotlink)
                Call moEForm.ReBuildCRFElement(oElement, False, True)
                bDidSomething = True
            End If
        End If
    
    End If
    
    EditHotlink = bDidSomething

End Function

'---------------------------------------------------------------------
Public Sub UpdateFieldFontColour()
'---------------------------------------------------------------------
' Update the font colour of a CRF Element
'---------------------------------------------------------------------
Dim oElement As CRFElement

    On Error GoTo ErrHandler
    
    ' Get out if no selected element
    If mnCRFElementId = 0 Then Exit Sub
    
    Set oElement = moEForm.CRFElementById(mnCRFElementId)
    
    On Error Resume Next    ' Ignore "Cancel error"
    
    ' Show the Color dialog
    CommonDialog1.CancelError = True
    CommonDialog1.ShowColor
    
    ' Ignore if they pressed Cancel
    If Err.Number <> cdlCancel Then
    
        On Error GoTo ErrHandler
        
        ' Update the element
        'ZA 24/07/2002 - check whether to update the caption or the question
        
        If mbIsCaption Then
            Call oElement.SetCaptionFontColour(CommonDialog1.Color)
        Else
            Call oElement.SetFontColour(CommonDialog1.Color)
        End If
        
        
        ' Redraw it to show the change
        Call moEForm.ReBuildCRFElement(oElement, False, True)
    
    End If
    
    Set oElement = Nothing

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "UpdateFieldFontColour")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    

End Sub

'---------------------------------------------------------------------
Public Sub UpdateCRFPageBackgroundColour()
'---------------------------------------------------------------------
'Program Changed by Mo Morris 10/12/98 (SPR 451)
'Call to ChangeCRFPageBackgroundColour replaced by a call to BuildCRFPage
'for the purpose of displaying the new colour in a forms controls.
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    On Error Resume Next
    CommonDialog1.CancelError = True
    CommonDialog1.ShowColor
    
    If Err.Number <> cdlCancel Then
    
        On Error GoTo ErrHandler
        ' NCJ 29 Nov 01 - Use new eForm object to set the colour
        moEForm.BackgroundColour = CLng(CommonDialog1.Color)
        
        ' Redraw the page
        Call RefreshMe
    
    End If
        
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "UpdateCRFPageBackgroundColour")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Public Sub UpdateFieldFont()
'---------------------------------------------------------------------
'Changed by Mo Morris   7/9/98     (change made in Released and Developed versions)
'CommonDialog.ShowFont flags now set to cdlCFBoth Or cdlCFScalableOnly Or cdlCFWYSIWYG.
'This restricts dialog to scalable fonts that are available on both Screen and Printer.
'Changed by Mo Morris   29/4/99
'---------------------------------------------------------------------
Dim oElement As CRFElement

    On Error GoTo ErrHandler
    
    Set oElement = moEForm.CRFElementById(mnCRFElementId)
    
    ' NCJ 29 Nov 01 - Use oElement properties (already set up)
    'ZA 24/07/2002 - check if it is a question or a caption
    
    If mbIsCaption Then
        CommonDialog1.FontName = oElement.CaptionFontName
        CommonDialog1.FontBold = oElement.CaptionFontBold
        CommonDialog1.FontItalic = oElement.CaptionFontItalic
        CommonDialog1.FontSize = oElement.CaptionFontSize
    Else
        CommonDialog1.FontName = oElement.FontName
        CommonDialog1.FontBold = oElement.FontBold
        CommonDialog1.FontItalic = oElement.FontItalic
        CommonDialog1.FontSize = oElement.FontSize
    End If
        
    On Error Resume Next
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlCFBoth Or cdlCFScalableOnly Or cdlCFWYSIWYG
    CommonDialog1.ShowFont
    
    'Changed by Mo Morris   29/4/99
    'Because the fonts dialog is restricted to fonts that are available on the printer as well
    'as the computer, if no printers are installed then the user will have been informed by the
    'unhelpful message 'THERE ARE NO FONTS INSTALLED OPEN THE FONTS FOLDER etc'
    'The following code traps the generated error code and displays a more helpful message
    If Err.Number = 24574 Then
    
        DialogInformation "You will not be able to edit fonts until you have installed a default printer."

    ElseIf Err.Number <> cdlCancel Then
    
        On Error GoTo ErrHandler
                
        ' Change the element's font details
        'Za 24/07/2002 - check if it is a question or a caption
        
        'TA 21/05/2003: Font size must be positive and less than 24000 or we crash
        If CommonDialog1.FontSize < 1 Or CommonDialog1.FontSize > 24000 Then
            DialogError "Font size must be between 1 and 24000"
'EXIT SUB HERE
Exit Sub
        End If
        
        If mbIsCaption Then
            Call oElement.SetCaptionFontDetails(CommonDialog1.FontName, CommonDialog1.FontSize, _
                                                CommonDialog1.FontBold, CommonDialog1.FontItalic)
        Else
        
            Call oElement.SetFontDetails(CommonDialog1.FontName, CommonDialog1.FontSize, _
                                         CommonDialog1.FontBold, CommonDialog1.FontItalic)
        End If
        
        ' Rebuild the element to show the change
        Call moEForm.ReBuildCRFElement(oElement, False, True)
        
    End If
    
    Set oElement = Nothing
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "UpdateFieldFont")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub MouseDownOnCRFElement(nIndex As Integer, Button As Integer, _
                                Shift As Integer, X As Single, Y As Single, _
                                sToolTipText As String)
'---------------------------------------------------------------------
' NCJ 16/12/99 - Don't do anything unless they have authorisation
' NB X and Y are LOCAL COORDINATES to the control which was clicked
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    ' RS 20-08-2002: End MultiSelectMode if Element clicked without CTRL
    If moMultiSelect.MultiSelectModeActive Then
        If Not moMultiSelect.ControlKeyPressed Then
            ' End SelectMode
            moMultiSelect.EndSelectMode
        Else
            ' Toggle the Select Status of this CRFElement
            moMultiSelect.ToggleSelectStatus nIndex, mbIsCaption
            Exit Sub
        End If
    End If
    
    ' NCJ 6 Sept 06 - If things have changed, refresh ourselves and skip out
    If frmMenu.RefreshIsNeeded Then Exit Sub

    ' NCJ 27 Feb 02 - Code moved to new routine
    ' NCJ 11 May 06 - Brought out of loop below
    Call ChangeSelectedElement(nIndex, sToolTipText)
    
    ' NCJ 11 May 06
'    If meFormAccessMode = sdReadOnly Then Exit Sub
    
    If UpdateMode <> gsREAD Then
        'ZA 07/08/2001 - SR4232, check if LostFocus event of txtEditCaption
        'has not been fired yet and Index of the control has been changed.
        'If not call the LostFocus event of txtEditCaption
        If (mbEditBox_LostFocus = False) And (mnPrevIndex <> nIndex) Then
            txtEditCaption_LostFocus
        End If
        
        ' Store the click coordinates
        mnX = X
        mnY = Y
        
        If lvwPopupList.Visible = True Then
            lvwPopupList.Visible = False
        End If
        
        ' NCJ 27 Feb 02 - Check mnCRFElementID is valid
        If Button = vbRightButton And mnCRFElementId > 0 Then
            Me.SetFocus
            DoEvents    'SDM 19/01/00 SR2359
            PopupMenu frmMenu.mnuFLD
        End If
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, _
                                 "MouseDownOnCRFElement(" & nIndex & ")", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub ChangeSelectedElement(ByVal nIndex As Integer, ByVal sToolTipText As String)
'---------------------------------------------------------------------
' Change the selected element to the one with the given control index
' This sets up mnCRFElementID and the info box in top right corner of screen
' NCJ 2 Apr 03 - Decide whether it's a group caption (extra argument to ChangeSelectedItem)
'---------------------------------------------------------------------
Dim oElement As CRFElement
Dim bIsGroupCaption As Boolean

    ' NCJ 27 Feb 02 - A txtEditCaption_LostFocus call might have redrawn a question group
    ' and invalidated this Index, so we get out if there's an error here
    On Error GoTo IgnoreError

    Set oElement = moEForm.EFormItemByIndex(nIndex).Element
    
    mnCRFElementId = oElement.CRFelementID
    
    ' NCJ 27 Feb 02 - Back to normal error handling
    On Error GoTo ErrHandler
       
    ' NCJ 2 Apr 03 - For a group, include whether it's the caption
    bIsGroupCaption = True
    If oElement.QGroupID > 0 Then
        bIsGroupCaption = mbIsCaption
    End If
    
    frmMenu.ChangeSelectedItem gsCRF_ELEMENT_LABEL & Str(mnCRFElementId), _
            "eForm : " & tabCRF.SelectedItem.Caption _
            & " " & sToolTipText, bIsGroupCaption

    Set oElement = Nothing

Exit Sub
ErrHandler:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, _
                                 "ChangeSelectedElement(" & nIndex & ")", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select

Exit Sub
IgnoreError:
    ' NCJ 27 Feb 02 - Reset mnCRFElementID
    mnCRFElementId = 0

End Sub

'---------------------------------------------------------------------
Private Sub MouseMoveOverCRFElement(Source As Control, Button As Integer, _
                                Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' NCJ 16/12/99 - Check user is authorised to drag
' NB X and Y are LOCAL coords within the Source control
' NCJ 4 Oct 02 - If a group text element is dragged, drag the Group instead
' NCJ 14 Jun 06 - Do nothing if eform is read-only
'---------------------------------------------------------------------
Dim oElement As CRFElement
Dim oDragControl As Control

    On Error GoTo ErrHandler
    
    If (Not goUser.CheckPermission(gsFnMaintEForm)) Or Me.AccessMode = sdReadOnly Then
        Exit Sub
    End If
    
    If Button = vbLeftButton And mnCRFElementId > 0 Then
    
        Set oElement = moEForm.CRFElementById(mnCRFElementId)
        
        ' We don't currently allow dragging of group items
        If oElement.OwnerQGroupID > 0 Then Exit Sub
        
'        ' NCJ 4 Oct 02 - For group items, switch to dragging the group itself
'        If oElement.OwnerQGroupID > 0 Then
'            ' We can only cope with text box items!
'            ' (Other control types cause funny behaviour - to be sorted out some time)
'            If Source.Name <> "lblCRFElement" Then Exit Sub
'            ' Pick up the Group CRFElement
'            Set oElement = moEForm.OwnerCRFElement(oElement)
'            ' Pick up the picQuGroup control
'            Set oDragControl = picQuGroup(moEForm.IndexForElementRow(oElement, 1))
'            ' Switch to the correct CRFElementId
'            mnCRFElementId = oElement.CRFelementID
'            ' Now offset the coordinates of the original Source control to picQuGroup coordinates
'            ' NB Control sits inside picScroller inside picFixed inside picQuGroup
'            ' For X coord, ignore container X offsets because they should all be 0
'            ' (and we don't want to get confused by a horizontal scroll value)
'            mnX = mnX + Source.Left
'            mnY = mnY + Source.Top + Source.Container.Top + Source.Container.Container.Top
'        Else
            Set oDragControl = Source
'        End If
        
        If oDragControl.Name = "picCRFElement" Then
            oDragControl.BorderStyle = vbFixedSingle
        End If
            
    '    If Shift And oDragCOntrol.Name <> "lblCaption" Then
        If oElement.DataItemId > 0 Or oElement.QGroupID > 0 Then
            If Shift Then   ' Reverse menu option as set
                'Changed by Mo Morris 20/8/99, 1/10/99
                'Menu, Options, Combined caption/control movement setting checked when deciding
                'if a caption and/or control should be moved together
                If frmMenu.mnuOCombinedMovement.Checked Then
                    mblnMoveCaptionAndElement = False
                Else
                    mblnMoveCaptionAndElement = True
                End If
            Else        ' Not Shift
                If frmMenu.mnuOCombinedMovement.Checked Then
                    mblnMoveCaptionAndElement = True
                Else
                    mblnMoveCaptionAndElement = False
                End If
            End If
        Else        ' DataItemID = 0
            mblnMoveCaptionAndElement = False
        End If
        
'        Debug.Print Timer & " Begin drag " & oDragControl.Name
        ' NCJ 13 Dec 01 - Bring the thing being dragged to the front
        oDragControl.ZOrder
        oDragControl.Drag vbBeginDrag
    
    End If
    
    Set oElement = Nothing
    Set oDragControl = Nothing

Exit Sub
ErrHandler:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, _
                                 "MouseMoveOverCRFElement(" & Source.Name & ")", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub calCRFElement_LostFocus()
'---------------------------------------------------------------------

    calCRFElement.Visible = False

End Sub

'---------------------------------------------------------------------
Private Sub cmdAttachment_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If TypeOf Source Is TreeView Then
        DragDrop Source, cmdAttachment(Index).Left + X, cmdAttachment(Index).Top + Y
    Else
        DragDrop Source, X, Y
    End If
    
    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdAttachment_DragDrop")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdAttachment_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    mbIsCaption = False
    MouseDownOnCRFElement Index, Button, Shift, X, Y, cmdAttachment(Index).ToolTipText
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdAttachment_MouseDown")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'---------------------------------------------------------------------
Private Sub cmdAttachment_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
'---------------------------------------------------------------------
' NCJ 7 Oct 02 - Use GetControlTopLeft in case button is inside a group
'---------------------------------------------------------------------
Dim sglTop As Single
Dim sglLeft As Single

    Call moEForm.GetControlTopLeft(cmdAttachment(Index), sglTop, sglLeft)
    DragOver Source, sglLeft + X, sglTop + Y, State

'    DragOver Source, cmdAttachment(Index).Left + X, cmdAttachment(Index).Top + Y, State

End Sub

'---------------------------------------------------------------------
Private Sub cmdAttachment_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
    
    MouseMoveOverCRFElement cmdAttachment(Index), Button, Shift, X, Y
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdCalendar_Click(Index As Integer)
'---------------------------------------------------------------------
' Display the calendar control
' NCJ 19/2/01 - Set current date
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    ' NCJ 19/2/01 - Set current date
    calCRFElement.Year = DatePart("yyyy", Now)
    calCRFElement.Month = DatePart("m", Now)
    calCRFElement.Day = DatePart("d", Now)

    Call moEForm.DisplayCalendar(Index)
    calCRFElement.SetFocus
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdCalendar_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdCalendar_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    mbIsCaption = False
    MouseDownOnCRFElement Index, Button, Shift, X, Y, cmdCalendar(Index).ToolTipText
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdCalendar_MouseDown")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdPopupList_Click(Index As Integer)
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler

    Call moEForm.DisplayPopupList(Index)
    lvwPopupList.SetFocus
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdPopupList_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdPopupList_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    mbIsCaption = False
    MouseDownOnCRFElement Index, Button, Shift, X, Y, cmdPopupList(Index).ToolTipText
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdPopupList_MouseDown")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
  
    ' PN 26/09/99
    ' maintain selected item
    If InStr(frmMenu.lblSelectedItem, "CRF page") <> 1 Then
        frmMenu.ChangeSelectedItem "", ""
    End If
    
    'Mo 16/5/2003  Bug 1577
    frmMenu.mnuFPrintAllCRFPages.Enabled = True
    frmMenu.mnuFPrintCRFPage.Enabled = True
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "Form_Activate")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If TypeOf Source Is TreeView Then
        If State = 0 Then
            ' Source is entering form
            Source.DragIcon = LoadResPicture("nodrop", vbResIcon)
        ElseIf State = 1 Then
            ' Source is leaving form
            Source.DragIcon = Source.SelectedItem.CreateDragImage
        End If
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "Form_DragOver")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    If KeyCode = vbKeyF1 Then               ' Show user guide
        'ShowDocument Me.hwnd, gsMACROUserGuidePath
        
        'REM 07/12/01 - New Call for the MACRO Help
        Call MACROHelp(Me.hWnd, App.Title)
        
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "Form_KeyDown")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
    
Dim rsStudyDefinition As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    Me.Icon = frmMenu.Icon

    ' RS 3/9/2002
    Set moMultiSelect = New MultiSelect
        
    mblnDuringLoad = True
    
    ' Turn on key preview for form, so that F1 (Help) can be trapped by form
    Me.KeyPreview = True
    
    ' Specify position relative to mdiForm
    Me.Top = 0
    Me.Left = (frmMenu.Width / 4)
    Me.Width = ((frmMenu.Width / 4) * 3) - 200
    Me.Height = frmMenu.Height - (frmMenu.tlbMenu.Height) * 2.25 - frmMenu.sbrMenu.Height
    ' NCJ 16/2/00 - Changed "Case Report Forms" to "eForms"
    Me.Caption = Me.ClinicalTrialName & " eForms"
    
    txtEditCaption.Width = 10
    txtEditCaption.Left = -500
    
    imglistSmallIcons.ListImages.Add , gsCRF_PAGE_LABEL, _
            LoadResPicture(gsCRF_PAGE_LABEL, vbResIcon)
    
    ' PN change 43 - 02/09/99
    ' to conserve memory this control is going have its icon property
    ' set with the gsIconArrowData icon
    ' all dynamically built controls will not load the resource icon but instead use
    ' this controls property
    lblCRFElement(0).MouseIcon = LoadResPicture(gsIconArrowData, vbResCursor)
    
    'tabCRF.MultiRow = False
    
    Set rsStudyDefinition = New ADODB.Recordset
    Set rsStudyDefinition = gdsStudyDefinition(ClinicalTrialId, VersionId)
    
    Me.DefaultCRFColour = rsStudyDefinition!DefaultCRFPageColour
    Me.DefaultFontColour = rsStudyDefinition!DefaultFontColour
    Me.DefaultFontName = rsStudyDefinition!DefaultFontName
    Me.DefaultFontSize = rsStudyDefinition!DefaultFontSize
    Me.DefaultFontBold = rsStudyDefinition!DefaultFontBold
    Me.DefaultFontItalic = rsStudyDefinition!DefaultFontItalic
    
    CommonDialog1.Color = DefaultFontColour
    CommonDialog1.Flags = cdlCFPrinterFonts
    CommonDialog1.FontName = DefaultFontName
    CommonDialog1.FontBold = DefaultFontBold
    CommonDialog1.FontItalic = DefaultFontItalic
    CommonDialog1.FontSize = DefaultFontSize
    
    Me.Font.SIZE = DefaultFontSize
    Me.Font.Bold = DefaultFontBold
    Me.Font.Italic = DefaultFontItalic
    
    'datPicture.DatabaseName = gsMACRO_DATABASE
    
    RefreshCRF
    If Me.tabCRF.Tabs.Count > 0 Then
'        Set Me.tabCRF.SelectedItem = Me.tabCRF.Tabs(Me.tabCRF.Tabs.Count)
        Set Me.tabCRF.SelectedItem = Me.tabCRF.Tabs(1)
    End If
    
    'SDM 22/02/00 Added .Checked
    '   ATN 29/4/99 SR 841
    '   Check frmMenu setting to determine the grid size
    If frmMenu.mnuOCRFGridLarge.Checked = True Then
        Grid = "Large"
    ElseIf frmMenu.mnuOCRFGridMedium.Checked = True Then
        Grid = "Medium"
    ElseIf frmMenu.mnuOCRFGridSmall.Checked = True Then
        Grid = "Small"
    Else
        frmMenu.mnuOCRFGridNone = True
        Grid = "None"
    End If

    mblnDuringLoad = False
    'ZA 07/08/2001 - SR4232, initialise these variables when the form
    'is loaded first time
    mbEditBox_LostFocus = False
    mnPrevIndex = 0

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "Form_Load")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'---------------------------------------------------------------------
'If UnloadMode = vbFormControlMenu Then
'    Cancel = -1
'End If
    frmMenu.HideCRF
    
    'Mo 16/5/2003  Bug 1577
    frmMenu.mnuFPrintAllCRFPages.Enabled = False
    frmMenu.mnuFPrintCRFPage.Enabled = False
        
End Sub

'---------------------------------------------------------------------
Private Sub Form_Resize()
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    'changed Mo Morris 4/1/01, SR3802 call to lvwPopupList_LostFocus to prevent grey-out
    'box when the form is immediately re-sized
    If Me.WindowState = vbMinimized Then
        If Me.ActiveControl.Name = "lvwPopupList" Then
            lvwPopupList_LostFocus
        End If
    End If

    If Not mblnDuringLoad Then
        SetPageSize Me
    End If
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Form_Resize", Err.Source) = OnErrorAction.Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'---------------------------------------------------------------------

    ' Forget any errors during the unload - NCJ 9 Nov 99
    SaveSetting "ImedMACRO", "Options", "GridSize", Me.Grid
    
    Call TidyUpAndClearObjects
    
    ' NCJ 16/10/00 - Make sure the CRFPageId and ClinicalTrialId get reset to 0
    ' to prevent values being erroneously used next time the form is loaded
    mlCRFPageId = 0
    mlClinicalTrialId = -1
    
    ' RS 3/9/2002
    Set moMultiSelect = Nothing

End Sub

'---------------------------------------------------------------------
Private Sub TidyUpAndClearObjects()
'---------------------------------------------------------------------
' Tidy up before leaving or redrawing
'---------------------------------------------------------------------

    If Not moEForm Is Nothing Then
        Call moEForm.Terminate
        Set moEForm = Nothing
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub frmCRFPage_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'---------------------------------------------------------------------
    
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If TypeOf Source Is TreeView Then
        If State = 0 Then
            Source.DragIcon = LoadResPicture("nodrop", vbResIcon)
        ElseIf State = 1 Then
            Source.DragIcon = Source.SelectedItem.CreateDragImage
        End If
    End If
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "frmCRFPage_DragOver", Err.Source) = OnErrorAction.Retry Then
        Resume
    End If
End Sub

'-----------------------------------------------------
Private Sub hsbCRFPage_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'-----------------------------------------------------
    
'-----------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If TypeOf Source Is TreeView Then
        If State = 0 Then
            Source.DragIcon = LoadResPicture("nodrop", vbResIcon)
        ElseIf State = 1 Then
            Source.DragIcon = Source.SelectedItem.CreateDragImage
        End If
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "hsbCRFPage_DragOver")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'-----------------------------------------------------
Private Sub hsbQuGroup_Change(Index As Integer)
'-----------------------------------------------------
' Horizontal scroll of an RQG
'-----------------------------------------------------

    Call HScrollGroup(Index)
    
End Sub

'-----------------------------------------------------
Private Sub hsbQuGroup_Scroll(Index As Integer)
'-----------------------------------------------------
' Horizontal scroll of an RQG
'-----------------------------------------------------

    Call HScrollGroup(Index)
    
End Sub

'-----------------------------------------------------
Private Sub HScrollGroup(nIndex As Integer)
'-----------------------------------------------------
' Horizontal scroll of an RQG
'-----------------------------------------------------

    ' Must hide any calendar or popuplist that's showing
    lvwPopupList.Visible = False
    calCRFElement.Visible = False
    
    ' We've set it up with a dividing factor, stored in scroll bar Tag (NCJ 4 Nov 02)
    ' (to avoid overflow in integer scroll bar Value field)
    picScroller(nIndex).Left = -(hsbQuGroup(nIndex).Value * hsbQuGroup(nIndex).Tag)
    ' Scroll the headers by the same amount
    picHeaders(nIndex).Left = picScroller(nIndex).Left

End Sub

'-----------------------------------------------------
Private Sub lblCaption_dblClick(Index As Integer)
'-----------------------------------------------------
' NCJ 16/12/99 - Don't let them edit the caption unless they're authorised to do so
' NCJ 27 Feb 02 - Code moved to LabelDoubleClick
'-----------------------------------------------------

    If moMultiSelect.MultiSelectModeActive Then Exit Sub        ' RS: Disable doubleclick if multiselect active
    
    Call LabelDoubleClick(Index, lblCaption(Index).ToolTipText)
    
End Sub

'---------------------------------------------------------------------
Private Sub LabelDoubleClick(ByVal nIndex As Integer, ByVal sToolTipText As String)
'---------------------------------------------------------------------
' Double click on a label (caption or group header label)
' NCJ 14 Jun 06 - Ignore if ReadOnly
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler

    If goUser.CheckPermission(gsFnMaintEForm) And Me.AccessMode > sdReadOnly Then
        ' NCJ 27 Feb 02 - If we're in a group it may have just been redrawn
        ' so check that we know what the element is
        If mnCRFElementId = 0 Then
            Call ChangeSelectedElement(nIndex, sToolTipText)
        End If
        'ZA 07/08/2001 - SR4232, keep index of the current control and
        'set the lostFocus event of the txtEditCaption to false
        mnPrevIndex = nIndex
        mbEditBox_LostFocus = False
        Call EditCaption
    End If
    
    Exit Sub

ErrHandler:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, _
                                 "LabelDoubleClick(" & nIndex & ")", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub lblCaption_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'---------------------------------------------------------------------
' Either a question has landed on lblCaption from the tree view,
' or the lblCaption itself has been drag-dropped
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If TypeOf Source Is TreeView Then
        ' The drag is from the tree view
        ' Offset by the lblCpation's own coordinates
        DragDrop Source, lblCaption(Index).Left + X, lblCaption(Index).Top + Y
    Else
        DragDrop Source, X, Y
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, _
                                 "lblCaption_DragDrop(" & Index & ")", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub lblCaption_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
'---------------------------------------------------------------------
    
    DragOver Source, lblCaption(Index).Left + X, lblCaption(Index).Top + Y, State
    
End Sub

'---------------------------------------------------------------------
Private Sub lblCaption_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    mbIsCaption = True
    
    MouseDownOnCRFElement Index, Button, Shift, X, Y, lblCaption(Index).ToolTipText

Exit Sub
ErrHandler:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, _
                                 "lblCaption_MouseDown(" & Index & ")", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub lblCaption_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
    
    If moMultiSelect.MultiSelectModeActive Then Exit Sub        ' RS: Disable doubleclick if multiselect active
    
    MouseMoveOverCRFElement lblCaption(Index), Button, Shift, X, Y
    
End Sub

'---------------------------------------------------------------------
Private Sub lblCRFElement_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'---------------------------------------------------------------------
' Either a question has landed on lblCRFElement from the tree view,
' or the lblCRFElement itself has been drag-dropped
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If TypeOf Source Is TreeView Then
        DragDrop Source, lblCRFElement(Index).Left + X, lblCRFElement(Index).Top + Y
    Else
        DragDrop Source, AlignToGrid(X), AlignToGrid(Y)
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "lblCRFElement_DragDrop")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    

End Sub

'---------------------------------------------------------------------
Private Sub lblCRFElement_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
'---------------------------------------------------------------------
Dim sglTop As Single
Dim sglLeft As Single

    Call moEForm.GetControlTopLeft(lblCRFElement(Index), sglTop, sglLeft)
    DragOver Source, sglLeft + X, sglTop + Y, State
    
End Sub

'---------------------------------------------------------------------
Private Sub lblCRFElement_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    mbIsCaption = False
    
    'Changed by Mo Morris 4/1/00, force close down  of the calendar display (SR 2569)
    If calCRFElement.Visible Then
        calCRFElement.Visible = False
    End If
    
    MouseDownOnCRFElement Index, Button, Shift, X, Y, lblCRFElement(Index).ToolTipText
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "lblCRFElement_MouseDown")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub lblCRFElement_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
    
    MouseMoveOverCRFElement lblCRFElement(Index), Button, Shift, X, Y
    
End Sub

'---------------------------------------------------------------------
Private Sub lblGroupHeader_DblClick(Index As Integer)
'---------------------------------------------------------------------
' Double click on a row header in a group - same as editing a caption
'---------------------------------------------------------------------

    ' NCJ 27 Feb 02 - Code moved to LabelDoubleClick
    Call LabelDoubleClick(Index, lblGroupHeader(Index).ToolTipText)
    
End Sub

'---------------------------------------------------------------------
Private Sub lblGroupHeader_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
'---------------------------------------------------------------------
' Dragging over a row header in a group
' Need to convert local X and Y to CRFpage coordinates
'---------------------------------------------------------------------
Dim sglTop As Single
Dim sglLeft As Single

    Call moEForm.GetControlTopLeft(lblGroupHeader(Index), sglTop, sglLeft)
    DragOver Source, sglLeft + X, sglTop + Y, State
    
End Sub

'---------------------------------------------------------------------
Private Sub lblGroupHeader_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' Mouse down on a row header in a group
'---------------------------------------------------------------------

    mbIsCaption = True
    MouseDownOnCRFElement Index, Button, Shift, X, Y, lblGroupHeader(Index).ToolTipText

End Sub

'---------------------------------------------------------------------
Private Sub lvwPopupList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    ' When a ColumnHeader object is clicked, the ListView control is
    ' sorted by the subitems of that column.
    ' Set the SortKey to the Index of the ColumnHeader - 1
    lvwPopupList.SortKey = ColumnHeader.Index - 1
    
    ' Reverse the sort order
    If lvwPopupList.SortOrder = lvwAscending Then
        lvwPopupList.SortOrder = lvwDescending
    Else
        lvwPopupList.SortOrder = lvwAscending
    End If
    
    ' Set Sorted to True to sort the list.
    lvwPopupList.Sorted = True
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "lvwPopupList_ColumnClick")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub lvwPopupList_LostFocus()
'---------------------------------------------------------------------
' Get rid of the popup list
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'Mo Morris 4/1/01, clear added
    lvwPopupList.ListItems.Clear
    lvwPopupList.Height = 0
    lvwPopupList.Top = -10
    lvwPopupList.Refresh
 
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "lvwPopupList_LostFocus")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub optCRFElement_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    
    On Error GoTo ErrHandler
    
    If TypeOf Source Is TreeView Then
        DragDrop Source, optCRFElement(Index).Container.Left + X, optCRFElement(Index).Container.Top + Y
    Else
        DragDrop Source, X, Y
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "optCRFElement_DragDrop")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call MACROEnd
            Call MACROEnd
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub optCRFElement_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
'---------------------------------------------------------------------
' NCJ 7 Oct 02 - Use GetControlTopLeft in case optCRFElement is inside a group
' RS  24 Feb 2003 - Undid previous fix, which caused 'earthquake effects'
'---------------------------------------------------------------------
Dim sglTop As Single
Dim sglLeft As Single
    
    
    
    DragOver Source, optCRFElement(Index).Container.Left + X, _
                    optCRFElement(Index).Container.Top + optCRFElement(Index).Top + Y, State

    
    'Call moEForm.GetControlTopLeft(optCRFElement(Index).Container, sglTop, sglLeft)
    'DragOver Source, sglLeft + X, sglTop + Y, State
    
    
    
End Sub

'---------------------------------------------------------------------
Private Sub optCRFElement_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
        
    
    mbIsCaption = False
    MouseDownOnCRFElement optCRFElement(Index).Container.Index, Button, Shift, optCRFElement(Index).Left + X, optCRFElement(Index).Top + Y, optCRFElement(Index).ToolTipText

End Sub

'---------------------------------------------------------------------
Private Sub optCRFElement_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
    
    
    MouseMoveOverCRFElement optCRFElement(Index).Container, Button, Shift, optCRFElement(Index).Container.Left + X, optCRFElement(Index).Container.Top + Y
    
End Sub

'---------------------------------------------------------------------
Private Sub optCRFGroup_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'---------------------------------------------------------------------
' Either a question has landed on optCRFGroup from the tree view,
' or the optCRFGroup itself has been drag-dropped
'---------------------------------------------------------------------
    
    If moMultiSelect.MultiSelectModeActive Then Exit Sub
    If TypeOf Source Is TreeView Then
        DragDrop Source, optCRFGroup(Index).Left + X, optCRFGroup(Index).Top + Y
    Else
        DragDrop Source, X, Y
    End If

End Sub

'---------------------------------------------------------------------
Private Sub optCRFGroup_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
'---------------------------------------------------------------------
' Option group drag
' NCJ 7 Oct 02 - Use GetControlTopLeft in case optCRFGroup is inside a group
'---------------------------------------------------------------------
Dim sglTop As Single
Dim sglLeft As Single
    
    
    If moMultiSelect.MultiSelectModeActive Then Exit Sub
    DragOver Source, optCRFGroup(Index).Left + X, optCRFGroup(Index).Top + Y, State
    'Call moEForm.GetControlTopLeft(optCRFGroup(Index), sglTop, sglLeft)
    'DragOver Source, sglLeft + X, sglTop + Y, State

End Sub

'---------------------------------------------------------------------
Private Sub optCRFGroup_MouseDown(Index As Integer, _
                                Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' Mouse down in an option group
' X and Y are local coords within the user control
'---------------------------------------------------------------------
    
    mbIsCaption = False
    MouseDownOnCRFElement Index, Button, Shift, X, Y, optCRFGroup(Index).ToolTipText

End Sub

'---------------------------------------------------------------------
Private Sub optCRFGroup_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' Option group drag
' X and Y are local coords within the user control
'---------------------------------------------------------------------
        
    
    If moMultiSelect.MultiSelectModeActive Then Exit Sub        ' RS 14/01/2003: Disable doubleclick if multiselect active
    MouseMoveOverCRFElement optCRFGroup(Index), Button, Shift, X, Y
End Sub

'---------------------------------------------------------------------
Private Sub optGraphical_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'---------------------------------------------------------------------
' Either an optGraphical has been drag dropped
' or something has landed here
'---------------------------------------------------------------------

On Error GoTo ErrHandler

    If TypeOf Source Is TreeView Then
        DragDrop Source, optGraphical(Index).Container.Left + X, optGraphical(Index).Container.Top + Y
    Else
        DragDrop Source, X, Y
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "optGraphical_DragDrop")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub optGraphical_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
'---------------------------------------------------------------------
' NCJ 7 Oct 02 - Use GetControlTopLeft in case optCRFGroup is inside a group
' RS  24 Feb 2003 - Undid previous fix, which caused 'earthquake effects'
'---------------------------------------------------------------------
Dim sglTop As Single
Dim sglLeft As Single
    
    DragOver Source, optGraphical(Index).Container.Left + X, _
                    optGraphical(Index).Container.Top + optGraphical(Index).Top + Y, State
                    
'    Call moEForm.GetControlTopLeft(optGraphical(Index).Container, sglTop, sglLeft)
'    DragOver Source, sglLeft + X, sglTop + Y, State
       
End Sub

'---------------------------------------------------------------------
Private Sub optGraphical_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    mbIsCaption = False
    MouseDownOnCRFElement optGraphical(Index).Container.Index, Button, Shift, optGraphical(Index).Left + X, optGraphical(Index).Top + Y, optGraphical(Index).ToolTipText
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "optGraphical_MouseDown")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub optGraphical_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------

    MouseMoveOverCRFElement optGraphical(Index).Container, Button, Shift, optGraphical(Index).Left + X, optGraphical(Index).Top + Y
    
End Sub

'---------------------------------------------------------------------
Private Sub picCRFElement_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
'---------------------------------------------------------------------
' NCJ 7 Oct 02 - Use GetControlTopLeft in case optCRFGroup is inside a group
'---------------------------------------------------------------------
Dim sglTop As Single
Dim sglLeft As Single

    ' RS 05/06/2003: Disable event when in multiselect mode, to prevent picture element from moving
    If moMultiSelect.MultiSelectModeActive Then
        Exit Sub
    End If
    
'    DragOver Source, picCRFElement(Index).Left + X, picCRFElement(Index).Top + Y, State
    Call moEForm.GetControlTopLeft(picCRFElement(Index), sglTop, sglLeft)
    DragOver Source, sglLeft + X, sglTop + Y, State
    
End Sub

'---------------------------------------------------------------------
Private Sub picCRFPage_Click()
'---------------------------------------------------------------------
    
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler

    If moMultiSelect.Working Then
        Exit Sub
    End If

    Call ChangeSelectedEForm(tabCRF.SelectedItem.Caption)
'    frmMenu.ChangeSelectedItem gsCRF_PAGE_LABEL, "eForm : " & tabCRF.SelectedItem.Caption & " (" & GetAccessModeString(meFormAccessMode) & ")"
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "picCRFPage_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub picCRFPage_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    ' NCJ 11 May 06
    If meFormAccessMode = sdReadOnly Then Exit Sub
    
    If moMultiSelect.Working Then
        Exit Sub
    End If

    DragOver Source, AlignToGrid(X), AlignToGrid(Y), State
    
End Sub

'---------------------------------------------------------------------
Private Sub picCRFElement_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    ' RS 05/06/2003: Disable event when in multiselect mode, to prevent picture element from moving
    If moMultiSelect.MultiSelectModeActive Then
        Exit Sub
    End If
        
    ' NCJ 11 May 06
    If meFormAccessMode = sdReadOnly Then Exit Sub
    
    If TypeOf Source Is TreeView Then
        DragDrop Source, picCRFElement(Index).Left + X, picCRFElement(Index).Top + Y
    Else
        DragDrop Source, X, Y
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "picCRFElement_DragDrop")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub picCRFElement_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------

    ' RS 05/06/2003: Disable event when in multiselect mode, to prevent picture element from moving
    If moMultiSelect.MultiSelectModeActive Then
        Exit Sub
    End If
    
    ' NCJ 11 May 06
    If meFormAccessMode = sdReadOnly Then Exit Sub
    
    MouseMoveOverCRFElement picCRFElement(Index), Button, Shift, X, Y

End Sub

'---------------------------------------------------------------------
Private Sub picCRFPage_DragDrop(Source As Control, X As Single, Y As Single)
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler

    ' NCJ 11 May 06
    If meFormAccessMode = sdReadOnly Then Exit Sub
    
    ' RS 23/08/2002: End select mode if user adds Element to Eform
    moMultiSelect.EndSelectMode

    DragDrop Source, AlignToGrid(X), AlignToGrid(Y)
    
    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "picCRFPage_DragDrop")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'---------------------------------------------------------------------
Private Sub picCRFElement_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
      
    mbIsCaption = False
    MouseDownOnCRFElement Index, Button, Shift, X, Y, picCRFElement(Index).ToolTipText
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "picCRFElement_MouseDown")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub picCRFPage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
'Program changed by Mo Morris 25/9/98 (SPR 430)
'picCRFPage_MouseDown now calls EditCaption instead of displaying an InputBox.
'Note that an empty string comment is created prior to the call to EditCaption
' Tidying up - NCJ 1 Oct 99
' NCJ 16 Dec 99 - Ignore click unless user has authorisation to change form
' ic 30/11/00   - added '.jpg' suffix to common dialog filter
' ic 04/12/00   - removed .cur, ico, .bmp from filter as requested by AN
' NCJ 15 May 06 - Consider RO forms
'---------------------------------------------------------------------
Dim oElement As CRFElement
Dim nCRFElementID As Integer
Dim sDestFile As String
Dim sPicFileName As String

'Debug.Print Timer & " PAGE Mousedown " & X & "," & Y

    On Error GoTo ErrHandler
        
    If moMultiSelect.Working Then
        Exit Sub
    End If

    ' NCJ 6 Sept 06 - Check for changes from other users
    If frmMenu.RefreshIsNeeded Then Exit Sub
    
    ' NCJ 14 Jun 06 - Don't leap out here because we can now handle ReadOnly
'    ' NCJ 16/12/99
'    If Not goUser.CheckPermission(gsFnMaintEForm) Then Exit Sub
'
'    If UpdateMode = gsREAD Then Exit Sub
    
    Select Case Button
    Case vbRightButton
    
            ' RS 26/08/2002: End Select Mode
            moMultiSelect.EndSelectMode
    
        If tabCRF.Tabs.Count > 0 Then
            Me.SetFocus
            'line added by Mo Morris 17/12/99
            'y co-ordinate stored for use mnuFLDInsertSpace and mnuFLDRemoveSpace
            ' NCJ 30 Nov 01 - Also need gsgMouseDownOnFormX for adding question groups
            gsgMouseDownOnFormY = Y
            gsgMouseDownOnFormX = X
            
            'ZA 03/08/01 - SR2895, need to call the ChangeSelectedItem procedure to enable
            'form selection and popup menu display on right click of mouse
'            frmMenu.ChangeSelectedItem gsCRF_PAGE_LABEL, "eForm : " & tabCRF.SelectedItem.Caption
            Call ChangeSelectedEForm(tabCRF.SelectedItem.Caption)

            'line added by Mo Morris 20/8/99 - launch the forms right-mouse menu
            PopupMenu frmMenu.mnuFLD
        End If
        
    Case vbLeftButton
        ' They left-clicked - see which toolbar button is pressed
        ' NCJ 15 May 06 - Check access mode
        If goUser.CheckPermission(gsFnMaintEForm) And meFormAccessMode > sdReadOnly Then
        
             If frmMenu.tlbMenu.Buttons(gsLINE_LABEL).Value = tbrPressed Then
                 
                 ' RS 23/08/2002: End select mode if user adds Element to Eform
                 moMultiSelect.EndSelectMode
                 
                 Set oElement = moEForm.NewCRFElement(0, AlignToGrid(X), AlignToGrid(Y), _
                                                     0, 0, "", gn_LINE)
                 Call moEForm.BuildNewCRFElement(oElement)
                 mnCRFElementId = oElement.CRFelementID
                     
                 frmMenu.tlbMenu.Buttons(gsLINE_LABEL).Value = tbrUnpressed
                 
             ElseIf frmMenu.tlbMenu.Buttons(gsCOMMENT_LABEL).Value = tbrPressed Then
                 
                 ' RS 23/08/2002: End select mode if user adds Element to Eform
                 moMultiSelect.EndSelectMode
                 
                 ' Create an empty comment
                 ' NCJ 15 Aug 02 - Changed caption from Null to ""
                 Set oElement = moEForm.NewCRFElement(0, AlignToGrid(X), AlignToGrid(Y), _
                                                     0, 0, "", gn_COMMENT)
                 Call moEForm.BuildNewCRFElement(oElement)
                 mnCRFElementId = oElement.CRFelementID
                 
                 'Changed by Mo Morris 4/5/99 SR 889
                 'call to MouseDownOnCRFElement added to enable the correct menu options for the new comment
                 MouseDownOnCRFElement moEForm.IndexForElementRow(oElement, 1), _
                                         Button, Shift, AlignToGrid(X), AlignToGrid(Y), "Comment"
                 'allow user to enter/edit comment via call to EditCaption
                 frmMenu.tlbMenu.Buttons(gsCOMMENT_LABEL).Value = tbrUnpressed
                 EditCaption
                 
             ElseIf frmMenu.tlbMenu.Buttons(gsPICTURE_LABEL).Value = tbrPressed Then
                 
                 ' RS 23/08/2002: End select mode if user adds Element to Eform
                 moMultiSelect.EndSelectMode
                 
                 On Error Resume Next
                 CommonDialog1.CancelError = True
                 CommonDialog1.Flags = cdlOFNFileMustExist
                 'ic 30/11/00
                 'added extra suffix to filter to allow '.jpg' files
                 'ic 04/12/00
                 'removed .cur, ico, .bmp from filter as requested by AN
                 CommonDialog1.Filter = "Pictures (*.gif;*.wmf;*.jpg)|*.gif;*.wmf;*.jpg"
                 
                 CommonDialog1.FilterIndex = 1
                 CommonDialog1.ShowOpen
                 'If a valid picture file has been selected then store and display it
                 If Err.Number <> cdlCancel Then
                     sPicFileName = CommonDialog1.FileTitle
                     sDestFile = gsDOCUMENTS_PATH & sPicFileName
                     FileCopy CommonDialog1.FileName, sDestFile
                     Set oElement = moEForm.NewCRFElement(0, AlignToGrid(X), AlignToGrid(Y), _
                                                         0, 0, sPicFileName, gn_PICTURE)
                     Call moEForm.BuildNewCRFElement(oElement)
                     mnCRFElementId = oElement.CRFelementID
                     
                 End If
                 frmMenu.tlbMenu.Buttons(gsPICTURE_LABEL).Value = tbrUnpressed
                 
             ElseIf frmMenu.tlbMenu.Buttons(gsLINK_LABEL).Value = tbrPressed Then
                 
                 ' RS 23/08/2002: End select mode if user adds Element to Eform
                 moMultiSelect.EndSelectMode
                 
                 ' Create an empty hotlink
                 Set oElement = moEForm.NewCRFElement(0, AlignToGrid(X), AlignToGrid(Y), _
                                                     0, 0, "", gn_HOTLINK)
                 Call moEForm.BuildNewCRFElement(oElement)
                 mnCRFElementId = oElement.CRFelementID
                 
                 'call to MouseDownOnCRFElement added to enable the correct menu options for the new element
                 MouseDownOnCRFElement moEForm.IndexForElementRow(oElement, 1), _
                                         Button, Shift, AlignToGrid(X), AlignToGrid(Y), "Hotlink"
                 frmMenu.tlbMenu.Buttons(gsLINK_LABEL).Value = tbrUnpressed
                 'allow user to enter hotlink via call to EditHotlink
                 If Not EditHotlink(oElement) Then
                     ' They cancelled out so forget this element
                     Call moEForm.DeleteCRFElement(oElement)
                 End If
            Else
                 ' RS 20-08-2002: MultiSelect: Initiate Lasso Drag, as no toolbar buttons are pressed
                 moMultiSelect.Init Me, moEForm
                 moMultiSelect.StartSelectMode
                 moMultiSelect.StartLassoDrag X, Y
                 
             End If ' Which type of button
        End If ' If not ReadOnly
    End Select
    
    Set oElement = Nothing
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "picCRFPage_MouseDown")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub picHeaders_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'---------------------------------------------------------------------
' Map picHeader events to corresponding picQuGroup events after adjusting coordinates
'---------------------------------------------------------------------
Dim sglQGroupX As Single
Dim sglQGroupY As Single

    sglQGroupX = X + picHeaders(Index).Left + picFixed(Index).Left
    sglQGroupY = Y + picHeaders(Index).Top + picFixed(Index).Top
    Call picQuGroup_DragDrop(Index, Source, sglQGroupX, sglQGroupY)
    
End Sub

'---------------------------------------------------------------------
Private Sub picHeaders_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
'---------------------------------------------------------------------
' Map picHeader events to corresponding picQuGroup events after adjusting coordinates
'---------------------------------------------------------------------
Dim sglQGroupX As Single
Dim sglQGroupY As Single

    sglQGroupX = X + picHeaders(Index).Left + picFixed(Index).Left
    sglQGroupY = Y + picHeaders(Index).Top + picFixed(Index).Top
    Call picQuGroup_DragOver(Index, Source, sglQGroupX, sglQGroupY, State)

End Sub

'---------------------------------------------------------------------
Private Sub picHeaders_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' Map picHeader events to corresponding picQuGroup events after adjusting coordinates
'---------------------------------------------------------------------
Dim sglQGroupX As Single
Dim sglQGroupY As Single


    sglQGroupX = X + picHeaders(Index).Left + picFixed(Index).Left
    sglQGroupY = Y + picHeaders(Index).Top + picFixed(Index).Top
    Call picQuGroup_MouseDown(Index, Button, Shift, sglQGroupX, sglQGroupY)

End Sub

'---------------------------------------------------------------------
Private Sub picHeaders_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' Map picHeader events to corresponding picQuGroup events after adjusting coordinates
'---------------------------------------------------------------------
Dim sglQGroupX As Single
Dim sglQGroupY As Single

    sglQGroupX = X + picHeaders(Index).Left + picFixed(Index).Left
    sglQGroupY = Y + picHeaders(Index).Top + picFixed(Index).Top
    Call picQuGroup_MouseMove(Index, Button, Shift, sglQGroupX, sglQGroupY)
    
End Sub

'---------------------------------------------------------------------
Private Sub picQuGroup_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'---------------------------------------------------------------------
'---------------------------------------------------------------------
    
    If TypeOf Source Is TreeView Then
        DragDrop Source, picQuGroup(Index).Left + X, picQuGroup(Index).Top + Y
    Else
        DragDrop Source, X, Y
    End If

End Sub

'---------------------------------------------------------------------
Private Sub picQuGroup_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
'---------------------------------------------------------------------
    
    DragOver Source, picQuGroup(Index).Left + X, picQuGroup(Index).Top + Y, State

End Sub

'---------------------------------------------------------------------
Private Sub picQuGroup_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' Mouse down on a question group container
' NB The X,Y are local to picQuGroup
'---------------------------------------------------------------------
    
    mbIsCaption = False
    MouseDownOnCRFElement Index, Button, Shift, X, Y, picQuGroup(Index).ToolTipText

End Sub

'---------------------------------------------------------------------
Private Sub picQuGroup_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' Mouse move in a question group container
'---------------------------------------------------------------------
    
    MouseMoveOverCRFElement picQuGroup(Index), Button, Shift, X, Y

End Sub

'---------------------------------------------------------------------
Private Sub picScroller_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'---------------------------------------------------------------------
' Map picScroller events to corresponding picQuGroup events after adjusting coordinates
'---------------------------------------------------------------------
Dim sglQGroupX As Single
Dim sglQGroupY As Single

    sglQGroupX = X + picScroller(Index).Left + picFixed(Index).Left
    sglQGroupY = Y + picScroller(Index).Top + picFixed(Index).Top
    Call picQuGroup_DragDrop(Index, Source, sglQGroupX, sglQGroupY)
    
End Sub

'---------------------------------------------------------------------
Private Sub picScroller_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
'---------------------------------------------------------------------
' Map picScroller events to corresponding picQuGroup events after adjusting coordinates
'---------------------------------------------------------------------
Dim sglQGroupX As Single
Dim sglQGroupY As Single

    sglQGroupX = X + picScroller(Index).Left + picFixed(Index).Left
    sglQGroupY = Y + picScroller(Index).Top + picFixed(Index).Top
    Call picQuGroup_DragOver(Index, Source, sglQGroupX, sglQGroupY, State)

End Sub

'---------------------------------------------------------------------
Private Sub picScroller_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' Map picScroller events to corresponding picQuGroup events after adjusting coordinates
'---------------------------------------------------------------------
Dim sglQGroupX As Single
Dim sglQGroupY As Single

    sglQGroupX = X + picScroller(Index).Left + picFixed(Index).Left
    sglQGroupY = Y + picScroller(Index).Top + picFixed(Index).Top
    Call picQuGroup_MouseDown(Index, Button, Shift, sglQGroupX, sglQGroupY)

End Sub

'---------------------------------------------------------------------
Private Sub picScroller_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' Map picScroller events to corresponding picQuGroup events after adjusting coordinates
'---------------------------------------------------------------------
Dim sglQGroupX As Single
Dim sglQGroupY As Single

    sglQGroupX = X + picScroller(Index).Left + picFixed(Index).Left
    sglQGroupY = Y + picScroller(Index).Top + picFixed(Index).Top
    Call picQuGroup_MouseMove(Index, Button, Shift, sglQGroupX, sglQGroupY)
    
End Sub

'---------------------------------------------------------------------
Private Sub tabCRF_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If TypeOf Source Is TreeView Then
        If State = 0 Then
            Source.DragIcon = LoadResPicture("nodrop", vbResIcon)
        ElseIf State = 1 Then
            Source.DragIcon = Source.SelectedItem.CreateDragImage
        End If
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "tabCRF_DragOver")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub txtEditCaption_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
' They pressed a key in the caption edit box
'---------------------------------------------------------------------
Dim oElement As CRFElement

    On Error GoTo ErrHandler
    
    mbEditBox_LostFocus = False
    If KeyAscii = vbKeyEscape Then
        ' Shift the edit box out of the way
        txtEditCaption.Width = 10
        txtEditCaption.Left = -500
        'the tabCRF.SetFocus is to force a lostfocus on txtEditCaption
        tabCRF.SetFocus
        
        'check that prior to the edit of the comment/caption it was not
        'a newly created empty string, which needs to be deleted
        Set oElement = moEForm.CRFElementById(mnCRFElementId)
        If oElement.DataItemId = 0 And oElement.Caption = "" Then
            Call moEForm.DeleteCRFElement(oElement)
        End If

        ''WillC SR3688 & SR3689 Pipes and tildas were allowed to be typed in the text box
    ElseIf Not gblnValidString(Chr(KeyAscii), valOnlySingleQuotes) Then
        DialogInformation "A caption" & gsCANNOT_CONTAIN_INVALID_CHARS
        '  Cancel the keystroke
        KeyAscii = 0
        txtEditCaption.SetFocus
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtEditCaption_KeyPress")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub txtEditCaption_LostFocus()
'---------------------------------------------------------------------
' They finished editing a new caption
'---------------------------------------------------------------------
Dim oElement As CRFElement
Dim sNewCaption As String

    On Error GoTo ErrHandler
    
    'set the following variable to true since we are going through
    'LostFocus event
    mbEditBox_LostFocus = True
    
    '  Check if box has already been moved off the left of the form - this
    '  can only happen as a result of an Escape (ie cancel)
    ' Exit from sub if escape has previously been pressed
    If txtEditCaption.Left = -500 Then
        Exit Sub
    End If
    
    sNewCaption = Trim(txtEditCaption.Text)
    Set oElement = moEForm.CRFElementById(mnCRFElementId)
    
    
    'WillC SR3688 & SR3689 Cutting and pasting erroneously allowed quotes in the caption.
    If Not gblnValidString(sNewCaption, valOnlySingleQuotes) Then
        DialogInformation "A caption" & gsCANNOT_CONTAIN_INVALID_CHARS
        ' Reset text to its previous value
        txtEditCaption.Text = txtEditCaption.Tag
        txtEditCaption.SetFocus
        Exit Sub
    End If
    
    ' If it's a comment being edited and it's an empty string then delete it
    If sNewCaption = "" And oElement.DataItemId = 0 And oElement.QGroupID = 0 Then
        Call moEForm.DeleteCRFElement(oElement)
    Else
        Call oElement.SetCaption(sNewCaption)
        Call moEForm.ReBuildCRFElement(oElement, False, True)
    End If

    '  Hide the caption edit box
    txtEditCaption.Width = 10
    txtEditCaption.Left = -500
    
    Set oElement = Nothing

    ' We've now finished the LostFocus event
    mbEditBox_LostFocus = False

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtEditCaption_LostFocus")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select


End Sub

'---------------------------------------------------------------------
Private Sub tabCRF_Click()
'---------------------------------------------------------------------
' User clicked on an eForm tab
' NCJ 20 Jun 06 - MUSD - Check for updates
' NCJ 17 Apr 07 - Added mbSelectingEForm; moved some code to new routine
'---------------------------------------------------------------------
'Dim lNewCRFPageId As Long

    On Error GoTo ErrHandler
    
    mbSelectingEForm = True
    
    If Not Me.ActiveControl Is Nothing Then
        If Me.ActiveControl.Name = "txtEditCaption" Then
            txtEditCaption_LostFocus
        '   ATN 15/10/98
        '   Force setfocus so that txtEditCaption has definitely lost focus
            tabCRF.SetFocus
        End If
    End If
    
    ' NCJ 17 Apr 07 - Code moved to new routine
    Call TabCRFClick
    
'    ' PN 26/09/99
'    ' set current id to invalid value
'    mnCRFElementId = 0
'
'    lNewCRFPageId = CLng(Mid(tabCRF.SelectedItem.Key, Len(gsCRF_PAGE_LABEL) + 1))
'    ' NCJ 20 Jun 06 - Check for study changes
'    If frmMenu.RefreshIsNeeded Then
'        ' Try to select the eForm they originally wanted
'        Call SelectCRF(lNewCRFPageId)
'    ElseIf lNewCRFPageId <> mlCRFPageId Then
'        Call RefreshMe(lNewCRFPageId)
'    End If
'
'    Call ChangeSelectedEForm(tabCRF.SelectedItem.Caption)

    mbSelectingEForm = False
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "tabCRF_Click", Err.Source) = OnErrorAction.Retry Then
        Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub TabCRFClick()
'---------------------------------------------------------------------
' NCJ 17 Apr 07 - Code extracted from second half of tabCRF_Click
'---------------------------------------------------------------------
Dim lNewCRFPageId As Long

    ' Set current element id to invalid value
    mnCRFElementId = 0
    
    lNewCRFPageId = CLng(Mid(tabCRF.SelectedItem.Key, Len(gsCRF_PAGE_LABEL) + 1))
    ' NCJ 20 Jun 06 - Check for study changes
    If frmMenu.RefreshIsNeeded Then
        ' Try to select the eForm they originally wanted
        Call SelectCRF(lNewCRFPageId)
    ElseIf lNewCRFPageId <> mlCRFPageId Then
        Call RefreshMe(lNewCRFPageId)
    End If
    
    Call ChangeSelectedEForm(tabCRF.SelectedItem.Caption)

End Sub

'---------------------------------------------------------------------
Private Sub ChangeSelectedEForm(sCaption As String)
'---------------------------------------------------------------------
' NCJ 15 May 06 - Set the selected item to given eForm caption with its access mode
' NCJ 17 Apr 07 - Corrected call to GetAccessModeString!
'---------------------------------------------------------------------
    
    Call frmMenu.ChangeSelectedItem(gsCRF_PAGE_LABEL, _
                "eForm : " & sCaption & " (" & GetAccessModeString(meFormAccessMode, frmMenu.AllowMU) & ")")

End Sub

'---------------------------------------------------------------------
Private Sub vsbCRFPage_Change()
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    picCRFPage.Top = CSng(-vsbCRFPage.Value) * 10
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "vsbCRFPage_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub vsbCRFPage_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If TypeOf Source Is TreeView Then
        If State = 0 Then
            Source.DragIcon = LoadResPicture("nodrop", vbResIcon)
        ElseIf State = 1 Then
            Source.DragIcon = Source.SelectedItem.CreateDragImage
        End If
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "vsbCRFPage_DragOver")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub vsbCRFPage_Scroll()
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    picCRFPage.Top = CSng(-vsbCRFPage.Value) * 10
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "vsbCRFPage_Scroll")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub hsbCRFPage_Change()
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'scroll the picture box
    picCRFPage.Left = 0 - (hsbCRFPage.Value / 2)
        
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "hsbCRFPage_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub hsbCRFPage_Scroll()
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'scroll the picture box
    picCRFPage.Left = 0 - (hsbCRFPage.Value / 2)
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "hsbCRFPage_Scroll")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub DragDrop(ByVal vSource As Control, _
                     ByVal vX As Single, _
                     ByVal vY As Single)
'---------------------------------------------------------------------
' An object is being dropped
' NCJ 18 Dec 01 - Recognise Groups too
'REM 12/07/02 - Added Select Case Statment to choose between Questions and Groups
'---------------------------------------------------------------------
Dim lQGroupId As Long
Dim lTag As eNodeTag
Dim oElement As CRFElement
Dim lVerticalScrollPosition As Long     ' Use to restore vertical scrolling position
Dim lHorizontalScrollPosition As Long     ' Use to restore horizontal scrolling position


'Debug.Print Timer & " DragDrop " & vSource.Name & ", " & vX & "," & vY
    
    On Error GoTo ErrHandler
    
    If meFormAccessMode = sdReadOnly Then Exit Sub
    
    If UpdateMode = gsUPDATE And tabCRF.Tabs.Count > 0 Then
        If TypeOf vSource Is TreeView _
         And vSource.Name = "trvDataList" Then
         
            'get the Node tag from frmDataList method SlectedNodeTag
            lTag = vSource.Container.SelectedNodeTag
            
            
            ' Drag is from the Question List
            'REM 12/07/02 - Added Select Case Statment to choose between Questions and Groups
            Select Case lTag
            Case eNodeTag.QuestionTag 'Is it a question
                ' Check for DItem use - NCJ 18 Oct 99
                If vSource.Container.CanBeDropped Then    ' Can it be dropped here?
                    AddDataItemToCRF vSource, vX, vY
                End If
            Case eNodeTag.QGroupTag 'is it a Question Group
                ' For now we don't support dragging groups from other studies
                If vSource.Container.ClinicalTrialId = Me.ClinicalTrialId _
                 And vSource.Container.VersionId = Me.VersionId Then
                    ' Ask frmDataList to give us the GroupId
                    lQGroupId = vSource.Container.SelectedQGroupID
                    ' We must set up the global mouse click values required by DropEFormGroupOnCRF
                    gsgMouseDownOnFormX = vX
                    gsgMouseDownOnFormY = vY
                    Call DropEFormGroupOnCRF(lQGroupId)
                End If
            End Select
    
        ElseIf TypeOf vSource Is TabStrip Then
            ' Drag the tab strip
            mCopyCRFPage vSource
         
        ElseIf vSource.Container = Me.picCRFPage Then
            ' They dragged a CRFElement
            If vSource.Name = "picCRFElement" Then
                ' Switch off the border that we added when starting the drag
                vSource.BorderStyle = 0
            End If
            
            ' NCJ 14 May 03 - Don't do anything if it was a click rather than a drag,
            ' i.e. the thing hasn't moved

            ' We ignore the local X,Y and use the Left,Top of where the control now is
            ' NCJ 14 May 03 - Don't update eForm if nothing changed
            If MoveCRFElement(vSource, AlignToGrid(vSource.Left), AlignToGrid(vSource.Top)) Then
       
                ' NCJ 7 Oct 02 - Redraw if a group (to force resizing)
                ' NCJ 14 Feb 03 - Do a refresh whatever the element
                Set oElement = CRFElementById(mnCRFElementId)
                ' RS 14/01/2003: Preserve horizontal & vertical scroll position
                lVerticalScrollPosition = Me.vsbCRFPage.Value
                lHorizontalScrollPosition = Me.hsbCRFPage.Value
                If oElement.QGroupID > 0 Then
                    ' Draw whole eForm if a group (do this because of unsolved update funnies)
                    Call RefreshMe
                Else
                    ' Otherwise just do the element itself
                    Call RebuildCurrentElement
                End If
                Me.vsbCRFPage.Value = lVerticalScrollPosition
                Me.hsbCRFPage.Value = lHorizontalScrollPosition
                Set oElement = Nothing
            
            End If
            
            mblnMoveCaptionAndElement = False
        End If
    End If
    
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "DragDrop")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub DragOver(oSource As Control, _
                     ByVal vX As Single, _
                     ByVal vY As Single, _
                     ByVal vState As Integer)
'---------------------------------------------------------------------
' This gets called while something is being dragged around on the eForm.
' We must make sure that captions and elements stick together
' if combined caption/control movement is required,
' and make sure all the bits and pieces of a data entry field stick together.
' Actually the work gets done in SnapToGrid.
' NB vX and vY are coords on the picCRFPage
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    If UpdateMode = gsREAD Then
        oSource.DragIcon = LoadResPicture("Nodrop", vbResIcon)
    ElseIf TypeOf oSource Is TabStrip Then
        oSource.DragIcon = LoadResPicture(gsCRF_PAGE_LABEL, vbResIcon)
    Else
        'Changed by Mo Morris 4/5/99 SR 877
        'For some reason the DragOver can get triggered without mouse movement after the SnapToGrid
        'routine has decided to move the control by a grid increment (e.g. 200 twips for Large grid).
        'When this occurs the Vx or Vy co-ordinates tend to be 10 OR 5 twips on from the previous position and
        'trigger a diffrent snap position. This gets into an alternating loop that makes the control
        'flicker. The following line prevents these flickers occurring, but does not explain why the
        'different Vx and Vy values are being triggered.
        If (Abs(mnPrevX - vX) > 10) Or (Abs(mnPrevY - vY) > 10) Then
'Debug.Print "SnapToGrid in DragOver " & vX & "," & vY
            SnapToGrid oSource, vX, vY
            mnPrevX = vX
            mnPrevY = vY
            If vX > frmCRFPage.Width - vsbCRFPage.Width - 100 _
            Or vX < 100 _
            Or vY > frmCRFPage.Height - hsbCRFPage.Height - 100 _
            Or vY < 100 Then
'Debug.Print "MoveCRFElement in DragOver " & vX & "," & vY
                MoveCRFElement oSource, vX, vY
            End If
        End If
    End If
    
Exit Sub
ErrLabel:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, _
                                 "DragOver(" & oSource.Name & ")", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub AddDataItemToCRF(ByVal oSource As Control, _
                            ByVal vX As Single, _
                            ByVal vY As Single)
'---------------------------------------------------------------------
' Add a question to the eForm at the given position
' Assume NOT a gourp member
'---------------------------------------------------------------------
Dim lDataItemId As Long
Dim nNewCRFElementId As Integer
Dim sCaption As Variant
Dim oForm As Form
Dim sglCaptionX As Single
Dim sglCaptionY As Single
Dim oElement As CRFElement
Dim lTag As eNodeTag

    On Error GoTo ErrHandler
    '  Copy the data item if its being dragged from another trial
    '  or the library
    If oSource.Container.ClinicalTrialId <> Me.ClinicalTrialId _
    Or oSource.Container.VersionId <> Me.VersionId Then
        lDataItemId = CopyDataItem(oSource.Container.ClinicalTrialId, _
            oSource.Container.VersionId, _
            oSource.Container.SelectedDataItemId, _
            Me.ClinicalTrialId, _
            Me.VersionId)
    Else
        lDataItemId = oSource.Container.SelectedDataItemId
    End If
    
    'SPR 557    ATN 9/11/98
    '   Check to see if user has cancelled the operation and exit
    If lDataItemId = -1 Then
        Exit Sub
    End If
    
    '  Check if auto caption is turned on
    ' NCJ 29 Nov 01 - Calculate the CaptionXY coords here
    If frmMenu.mnuOAutoCaption.Checked = True Then
        sCaption = oSource.Container.SelectedDataItemName
        sglCaptionY = AlignToGrid(vY)
        ' Shift the caption over to the left of the field
        sglCaptionX = AlignToGrid(vX - CaptionWidth(CStr(sCaption)))
    Else
        ' RS 28/08/2002: Error using Null as a string value, corrected as described by NCJ
        ' sCaption = Null
        sCaption = ""
        sglCaptionX = 0
        sglCaptionY = 0
    End If
    
    
    
        
    '  Insert a new CRF element
    Set oElement = moEForm.NewCRFElement(lDataItemId, _
                AlignToGrid(vX), AlignToGrid(vY), sglCaptionX, sglCaptionY, _
                 sCaption)
    
    'check to see if the new element is within the bottom 5,000 twips of the form
    'if so extent the form
    If vY > picCRFPage.Height - 5000 Then
        'extend the form length
        picCRFPage.Height = vY + 5000
        'extend the scroll bar max value setting
        'note that the max value is divided by 10 to accomodate long forms
        vsbCRFPage.Max = picCRFPage.Height / 10
        Call DisplayGrid    'SDM 24/02/00 Ensure that the screen is constant
    End If
    
    '  Display the CRF element
    Call moEForm.BuildNewCRFElement(oElement)
    
    ' NCJ 19 Jun 06 - Mark study as changed
    Call frmMenu.MarkStudyAsChanged
    
    '  Update all the data lists
    For Each oForm In Forms
        If oForm.Name = "frmDataList" Then
            ' Is it our study?
            If oForm.ClinicalTrialId = Me.ClinicalTrialId Then
                Call oForm.AddCRFToDataItem(lDataItemId, CRFPageId, _
                            oElement.CRFelementID, tabCRF.SelectedItem.Caption, _
                            oSource.Container.SelectedDataItemName)
            End If
        End If
    Next
            
    Me.SetFocus
    
    Set oElement = Nothing
    Set oForm = Nothing

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "AddDataItemToCRF")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub mCopyCRFPage(ByVal vSource As Control)
'---------------------------------------------------------------------
Dim lNewCRFPageId As Long

    On Error GoTo ErrHandler
    
    lNewCRFPageId = CopyCRFPage(vSource.Parent.ClinicalTrialId, vSource.Parent.VersionId, _
                vSource.Parent.CRFPageId, Me.ClinicalTrialId, Me.VersionId)
    
    ' NCJ 15 Jan 02 - Refresh ALL question lists in target study
    Call frmMenu.RefreshQuestionLists(Me.ClinicalTrialId)
        
    Me.RefreshCRF
    Me.SetFocus
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "mCopyCRFPage")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'---------------------------------------------------------------------
Private Function MoveCRFElement(ByVal vSource As Control, _
                            ByVal vX As Single, _
                            ByVal vY As Single) As Boolean
'---------------------------------------------------------------------
' A CRFElement has been dragged to a new position
' so we must update the DB with the new values
' NCJ 14 May 03 - Return TRUE if it moved, or FALSE if it stayed the same
'---------------------------------------------------------------------
Dim sSQL As String
Dim oElement As CRFElement
Dim sSQLWhere As String

    On Error GoTo ErrHandler
    
    MoveCRFElement = False
    
    ' Make sure we don't try to move the tree view
    If vSource.Name = "trvDataList" Then Exit Function
    
    If mnCRFElementId = 0 Then Exit Function
    
    Set oElement = moEForm.CRFElementById(mnCRFElementId)
  
    ' We don't move group items
    If oElement.OwnerQGroupID > 0 Then Exit Function
    
    ' NCJ 14 May 03 - Don't do anything if nothing's changed!
    If vX = oElement.ElementX And vY = oElement.ElementY Then Exit Function
     
    'Changed by Mo Morris 30/4/99
    'check to see if the new element is within the bottom 5,000 twips of the form
    'if so extend the form
    If vY > picCRFPage.Height - 5000 Then
        'extend the form length
        picCRFPage.Height = vY + 5000
        'extend the scroll bar max value setting
        'note that the max value is divided by 10 to accommodate long forms
        vsbCRFPage.Max = picCRFPage.Height / 10
        'Mo Morris 23/8/01, SR4426
        Call DisplayGrid
    End If

    ' NCJ 3 Dec 01 - Use oElement's own routine to move itself
    Call oElement.Relocate(vX, vY, _
                        (vSource.Name = "lblCaption"), mblnMoveCaptionAndElement)
    MoveCRFElement = True
    
    Set oElement = Nothing
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "MoveCRFElement")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Function

'---------------------------------------------------------------------
Private Function AlignToGrid(ByVal sglCoord As Single) As Single
'---------------------------------------------------------------------
' Align a coordinate to the closest grid point
'---------------------------------------------------------------------
Dim sglExtraBit As Integer
Dim sglGridPoint As Single
    
    If GridSize > 1 Then
        ' There is a grid
        sglExtraBit = Abs(sglCoord Mod GridSize)
        sglGridPoint = sglCoord \ GridSize       ' Integer division
        
        If sglExtraBit < (GridSize / 2) Then
            ' Round down to closest grid point
            AlignToGrid = sglGridPoint * GridSize
        Else
            ' Round up to next grid point
            If sglCoord >= 0 Then
                AlignToGrid = (sglGridPoint + 1) * GridSize
            Else
                ' It's a negative coordinate
                AlignToGrid = (sglGridPoint - 1) * GridSize
            End If
        End If
    Else
        ' No grid so leave unchanged
        AlignToGrid = sglCoord
    End If
    
End Function

'---------------------------------------------------------------------
Public Sub DisplayGrid()
'---------------------------------------------------------------------
'   SDM 22/02/00 Displays a grid on the eForm
'---------------------------------------------------------------------
Dim blnGridDisplay As Boolean
Dim lngHori As Long
Dim lngVert As Long
Dim lngPichdc As Long
Dim lngPicWidth As Long
Dim lngPicHeight As Long
Dim lngGridStepWidth As Long
Dim lngGridStepHeight As Long
Dim lngMapMode As Long
Dim udtPoints As POINTAPI
    
    If GridSize = 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    'Make calculations so they are only done once.
    lngPicWidth = picCRFPage.ScaleWidth / Screen.TwipsPerPixelX
    lngPicHeight = picCRFPage.ScaleHeight / Screen.TwipsPerPixelY
    lngGridStepWidth = GridSize / Screen.TwipsPerPixelX
    lngGridStepHeight = GridSize / Screen.TwipsPerPixelY
    lngPichdc = picCRFPage.hDc
    blnGridDisplay = GridDisplay
    'Take a copy of map mode and then change to TWIPS
    lngMapMode = GetMapMode(lngPichdc)
    Call SetMapMode(lngPichdc, MM_TWIPS)
    'Draw dots on page
    For lngHori = 0 To lngPicWidth Step lngGridStepWidth
        For lngVert = 0 To lngPicHeight Step lngGridStepHeight
            'Convert from DeviceCoordinates to LogicalCoordinates
            udtPoints.X = lngHori
            udtPoints.Y = lngVert
            Call DPtoLP(lngPichdc, udtPoints, 1)
            'Paint dots
            If blnGridDisplay Then
                Call SetPixelV(lngPichdc, udtPoints.X, udtPoints.Y, picCRFPage.ForeColor)
            Else
                Call SetPixelV(lngPichdc, udtPoints.X, udtPoints.Y, picCRFPage.BackColor)
            End If
        Next lngVert
    Next lngHori
    
    'Restore map mode
    Call SetMapMode(lngPichdc, lngMapMode)
    Screen.MousePointer = vbDefault
    
End Sub

'---------------------------------------------------------------------
Public Property Get GridDisplay() As Boolean
'---------------------------------------------------------------------
'   SDM 22/02/00 For use with DisplayGrid
'---------------------------------------------------------------------
    GridDisplay = mbDisplayGrid
    
End Property

'---------------------------------------------------------------------
Public Property Let GridDisplay(ByVal bNewValue As Boolean)
'---------------------------------------------------------------------
'   SDM 22/02/00 For use with DisplayGrid
'---------------------------------------------------------------------
    mbDisplayGrid = bNewValue
    
End Property

'---------------------------------------------------------------------
Public Sub Renumber()
'---------------------------------------------------------------------
' This renumbers all CRFElements on the eForm based on X,Y positions
' NCJ 11 Dec 01 - It works OK for Question Groups, too
' MLM 16/09/02: Assign FieldOrder of 0 to form and visit dates.
'---------------------------------------------------------------------
Dim sResponse As String
Dim sSQL As String
Dim rsOrderedList As ADODB.Recordset
Dim sSQLUpdate As String
Dim nNewNumberOrder As Integer

         'Select all the data items on the current form and sort them in the order
         'of their Y and then their X co-ordinates on the form
         'Changed by Mo Morris 10/12/99
         'test for previously hidden data items with a FieldOrder of 0 added to SQL statement
         ' NCJ 11 Dec 99 - Added DataItemId > 0
         ' NCJ 30 Nov 01 - Added QGroupID > 0 and OwnerQGroupID = 0
         ' MLM 16/09/02: Added ElementUse = 0, i.e. don't bother loading the date element if it doesn't need renumbering
         sSQL = "Select * " _
             & " FROM CRFElement " _
             & " Where ClinicalTrialId = " & Me.ClinicalTrialId _
             & " AND VersionId = " & Me.VersionId _
             & " AND CRFPageid = " & frmCRFDesign.CRFPageId _
             & " AND ( DataItemId > 0 OR QGroupID > 0 )" _
             & " AND ( OwnerQGroupID = 0 )" _
             & " AND ((FieldOrder > 0) OR (Hidden = 0) OR (ElementUse = 0)) " _
             & " ORDER BY Y, X"
         
         Set rsOrderedList = New ADODB.Recordset
         rsOrderedList.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
                     
         'initialise new number order variable
         nNewNumberOrder = 0
         'read through each data item record to update the FieldOrder
         While Not rsOrderedList.EOF
             'changed by Mo Morris 10/12/99, hidden dataitems are not re-numbered
             sSQLUpdate = "Update CRFElement "
             'MLM 16/09/02: Assign FieldOrder of 0 to form and visit date qs
             If rsOrderedList!Hidden = False And rsOrderedList!ElementUse = 0 Then
                 nNewNumberOrder = nNewNumberOrder + 1
                 sSQLUpdate = sSQLUpdate _
                     & " SET FieldOrder = " & nNewNumberOrder
             Else
                 ' it's a hidden data item so set its Fieldorder to 0
                 sSQLUpdate = sSQLUpdate _
                     & " SET FieldOrder = 0 "
             End If
             sSQL = sSQLUpdate _
                     & " WHERE ClinicalTrialId = " & rsOrderedList![ClinicalTrialId] _
                     & " AND VersionId = " & rsOrderedList![VersionId] _
                     & " AND CRFPageId = " & rsOrderedList![CRFPageId] _
                     & " AND CRFElementId = " & rsOrderedList![CRFelementID]
            
             MacroADODBConnection.Execute sSQL
             
             ' NCJ 11 Dec 01 - We must also reset the field orders for all of a group's elements
             If rsOrderedList!QGroupID > 0 Then
                sSQL = sSQLUpdate _
                     & " WHERE ClinicalTrialId = " & rsOrderedList![ClinicalTrialId] _
                     & " AND VersionId = " & rsOrderedList![VersionId] _
                     & " AND CRFPageId = " & rsOrderedList![CRFPageId] _
                     & " AND OwnerQGroupId = " & rsOrderedList![QGroupID]
                MacroADODBConnection.Execute sSQL
             End If
             
             rsOrderedList.MoveNext
         Wend
         
         rsOrderedList.Close
         Set rsOrderedList = Nothing
         
         'Rebuild the form based on the re-numbered FieldOrder
         Call RefreshMe
         
         'Refresh data list, because it might be displaying data item in Field(Number)Order
         ' NCJ 15 Jan 02 - Refresh ALL question lists
         ' frmDataList.RefreshDataList
         Call frmMenu.RefreshQuestionLists(Me.ClinicalTrialId)

End Sub

'---------------------------------------------------------------------
Private Function CaptionWidth(ByVal sCaption As String) As Single
'---------------------------------------------------------------------
' Get the width of a text caption using this form's default font
'---------------------------------------------------------------------

    On Error Resume Next
        
    Me.FontName = Me.DefaultFontName
    Me.Font.Charset = 1
    Me.FontSize = Me.DefaultFontSize
    Me.FontBold = Me.DefaultFontBold
    Me.FontItalic = Me.DefaultFontItalic
    
    ' Add some extra characters for good measure
    CaptionWidth = TextWidth(sCaption & String(6, "_"))

End Function

'---------------------------------------------------------------------
Public Function CanAddQGroup(oQGroup As QuestionGroup) As Boolean
'---------------------------------------------------------------------
' Can we add this question group to this eForm?
' Only if none of its questions are already on the eForm
' RS 28/08/2002: Made function public
'---------------------------------------------------------------------
Dim bCanAddGroup As Boolean
Dim i As Integer

    bCanAddGroup = True
    If oQGroup.Questions.Count > 0 Then
        For i = 1 To oQGroup.Questions.Count
            If moEForm.QuestionIDExists(CLng(oQGroup.Questions(i))) Then
                bCanAddGroup = False
                Exit For
            End If
        Next i
    End If
    CanAddQGroup = bCanAddGroup

End Function

'---------------------------------------------------------------------
Public Function QuestionIDExists(ByVal lDataItemId As Long) As Boolean
'---------------------------------------------------------------------
' Does question exist on this eForm already?
'---------------------------------------------------------------------

    QuestionIDExists = moEForm.QuestionIDExists(lDataItemId)

End Function

'---------------------------------------------------------------------
Public Function QuestionCodeExists(ByVal sDataItemCode As String) As Boolean
'---------------------------------------------------------------------
' Does question exist on this eForm already?
'---------------------------------------------------------------------

    QuestionCodeExists = moEForm.QuestionCodeExists(sDataItemCode)

End Function

'---------------------------------------------------------------------
Public Sub RemoveCRFElement(nCRFElementID As Integer)
'---------------------------------------------------------------------
' Remove the question CRFElement with this ID
' This is called from frmDataList when a question is being deleted
' THIS DOES NOT WORK IF THE QUESTION IS A GROUP MEMBER...
'---------------------------------------------------------------------
Dim oElement As CRFElement

    Set oElement = moEForm.CRFElementById(nCRFElementID)
    Call moEForm.RemoveCRFElement(oElement)
    ' NCJ 19 Jun 06 - Mark study as changed
    Call frmMenu.MarkStudyAsChanged

End Sub

'---------------------------------------------------------------------
Public Sub DeleteQuestion(oElement As CRFElement)
'---------------------------------------------------------------------
' NCJ 11 Aug 03 - Created routine which removes AND deletes a question from this eForm
' (3.0 Roche bug fix no. 1944)
'---------------------------------------------------------------------

    Call moEForm.DeleteCRFElement(oElement)

End Sub

'---------------------------------------------------------------------
Public Sub RefreshEFormGroup(ByVal lQGroupId As Long)
'---------------------------------------------------------------------
' Reload the eForm group objects
' This should be called e.g. when a Question Group definition has changed
' The refresh only happens
' if that group is used on this eForm,
' and in this case the appropriate CRF Element is redrawn
' NCJ 14 May 03 - Preserve scroll position around the rebuild (Bug 1448)
'---------------------------------------------------------------------
Dim lHscroll As Long
Dim lVscroll As Long

    On Error GoTo ErrLabel
    
    ' NCJ 14 May 03 - Save Scroll Position
    lHscroll = hsbCRFPage.Value
    lVscroll = vsbCRFPage.Value
    
    Call moEForm.ReloadEFormGroup(lQGroupId)
    
    hsbCRFPage.Value = lHscroll
    vsbCRFPage.Value = lVscroll
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmCRFDesign.RefreshEFormGroup"
    
End Sub

'---------------------------------------------------------------------
Private Sub vsbQuGroup_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
'---------------------------------------------------------------------
' Dragging over a group scroll bar
' Need to convert local X and Y to CRFpage coordinates
'---------------------------------------------------------------------
Dim sglTop As Single
Dim sglLeft As Single

    Call moEForm.GetControlTopLeft(vsbQuGroup(Index), sglTop, sglLeft)
    DragOver Source, sglLeft + X, sglTop + Y, State

End Sub

'---------------------------------------------------------------------
Private Sub hsbQuGroup_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
'---------------------------------------------------------------------
' Dragging over a group horizontal scroll bar
' Need to convert local X and Y to CRFpage coordinates
'---------------------------------------------------------------------
Dim sglTop As Single
Dim sglLeft As Single

    Call moEForm.GetControlTopLeft(hsbQuGroup(Index), sglTop, sglLeft)
    DragOver Source, sglLeft + X, sglTop + Y, State

End Sub

'---------------------------------------------------------------------
' NEW FUNCTIONS FOR LASSO FUNCTION
'---------------------------------------------------------------------

'---------------------------------------------------------------------
Private Sub picCRFPage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' If currently dragging the lasso, update the lasso size/position
'---------------------------------------------------------------------

    On Error GoTo Errorlabel

    ' NCJ 11 May 06
    If meFormAccessMode = sdReadOnly Then Exit Sub
    
    If moMultiSelect.Working Then
        Exit Sub
    End If

    If moMultiSelect.LassoDragMode And Button = 1 Then
                
        ' UPDATE LASSO SIZE AND POSITION
        '
        moMultiSelect.UpdateLasso CLng(X), CLng(Y)
            
    End If
    
Exit Sub
Errorlabel:
    If MACROErrorHandler("frmCRFDesign", Err.Number, Err.Description, "picCRFPage_MouseMove", Err.Source) = Retry Then
        Resume
    End If

End Sub

'---------------------------------------------------------------------
Private Sub picCRFPage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' If currently dragging the lasso, end the lasso operation and
' build a collection of all CRFelements/Captions that are in the
' selection area
'---------------------------------------------------------------------

    On Error GoTo Errorlabel
    
    If moMultiSelect.Working Then
        Exit Sub
    End If
    
    ' RS 22/10/2002 IGNORE MOUSEUP if CONTROL KEY IS PRESSED
    If moMultiSelect.ControlKeyPressed Then
        Exit Sub
    End If


    If moMultiSelect.LassoDragMode And Button = 1 Then
        
        ' END OF LASSO DRAG OPERATION: HIDE LASSO AND INCLUDE CRFELEMENTS
        shpLasso.Visible = False
        moMultiSelect.UpdateSelection
        
    End If

Exit Sub
Errorlabel:
    If MACROErrorHandler("frmCRFDesign", Err.Number, Err.Description, "picCRFPage_MouseUp", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub shpDragHandle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' The SelectBox is only visible when MultiSelectMode is active (after
' a lasso has been dragged).
' Left Button:  Start Drag operation selected Items
' Right Button: Show context menu for multiselect operations
'---------------------------------------------------------------------

    On Error GoTo Errorlabel
    
    If Button = vbRightButton Then
        moMultiSelect.Command
        Exit Sub
    End If
    
    ' Start MultiSelect Drag Operation
    moMultiSelect.StartDrag X, Y

Exit Sub
Errorlabel:
    If MACROErrorHandler("frmCRFDesign", Err.Number, Err.Description, "shpDragHandle_MouseDown", Err.Source) = Retry Then
        Resume
    End If

End Sub

'---------------------------------------------------------------------
Private Sub shpDragHandle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' User is movingthe selected controls. Move the placeholder controls
'---------------------------------------------------------------------
    
    On Error GoTo Errorlabel
    
    If Button = vbLeftButton Then
        moMultiSelect.Drag X, Y
    End If
    
Exit Sub
Errorlabel:
    If MACROErrorHandler("frmCRFDesign", Err.Number, Err.Description, "shpDragHandle_MouseMove", Err.Source) = Retry Then
        Resume
    End If

End Sub

'---------------------------------------------------------------------
Private Sub shpDragHandle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' End of move operation. Move all selected elements and rebuild the form
'---------------------------------------------------------------------
    
    On Error GoTo Errorlabel

    If Button = vbLeftButton And moMultiSelect.Dragging Then
        ' End of Drag
        
        ' Update CRFelements
        moMultiSelect.EndDrag X, Y
        
        ' Redraw eForm
       ' RefreshMe
    End If
    
Exit Sub
Errorlabel:
    If MACROErrorHandler("frmCRFDesign", Err.Number, Err.Description, "shpDragHandle_MouseUp", Err.Source) = Retry Then
        Resume
    End If

End Sub

'---------------------------------------------------------------------
Public Property Get PasteBufferAvailable() As Boolean
'---------------------------------------------------------------------
' NCJ 29 May 03
' Are there elements in the paste buffer available for pasting?
' Use this for enabling/disabling "Paste" menu item
'---------------------------------------------------------------------

    PasteBufferAvailable = moMultiSelect.PasteBufferAvailable

End Property

'---------------------------------------------------------------------
Public Sub PasteSelection()
'---------------------------------------------------------------------
' User has selected Paste from CRFPage context menu
'---------------------------------------------------------------------

    moMultiSelect.PasteBuffer
    
End Sub
