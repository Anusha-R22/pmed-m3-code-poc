VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3BA5E907-B7F4-4504-90CD-12033A3F93F8}#1.1#0"; "MACROOptionGroup30.ocx"
Begin VB.Form frmEFormDataEntry 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "eForm"
   ClientHeight    =   5445
   ClientLeft      =   6255
   ClientTop       =   9675
   ClientWidth     =   10095
   ControlBox      =   0   'False
   Icon            =   "frmEFormDataEntry.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5445
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBlank 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   330
      ScaleHeight     =   315
      ScaleWidth      =   375
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComCtl2.FlatScrollBar hsbCRFPage 
      Height          =   240
      Left            =   750
      TabIndex        =   12
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   423
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1245185
   End
   Begin VB.Frame fraCRFPage 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   420
      TabIndex        =   3
      Top             =   360
      Width           =   9555
      Begin VB.PictureBox picCRFPage 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   4665
         Left            =   60
         ScaleHeight     =   4605
         ScaleWidth      =   8685
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   180
         Width           =   8745
         Begin VB.CommandButton cmdExpand 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   4560
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   3600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.PictureBox picBackButton 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   5580
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   34
            Top             =   4200
            Width           =   375
         End
         Begin VB.PictureBox picCommentStatus 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   0
            Left            =   1800
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   2400
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picLab 
            BackColor       =   &H00EEEEEE&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   0
            ScaleHeight     =   315
            ScaleWidth      =   8535
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   960
            Width           =   8535
            Begin VB.Label cmdChangeLab 
               BackColor       =   &H00EEEEEE&
               Caption         =   "Change laboratory"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   4320
               MouseIcon       =   "frmEFormDataEntry.frx":0442
               MousePointer    =   99  'Custom
               TabIndex        =   32
               Top             =   0
               Width           =   1545
            End
            Begin VB.Label lblLabName 
               BackStyle       =   0  'Transparent
               Caption         =   "Lab Name"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00AAAAAA&
               Height          =   195
               Left            =   1170
               TabIndex        =   31
               ToolTipText     =   "Laboratory name"
               Top             =   15
               Width           =   2505
            End
            Begin VB.Label lblLab 
               BackStyle       =   0  'Transparent
               Caption         =   "Laboratory:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00AAAAAA&
               Height          =   195
               Left            =   60
               TabIndex        =   30
               Top             =   15
               Width           =   1005
            End
         End
         Begin VB.PictureBox picHeaderBar 
            BackColor       =   &H00EEEEEE&
            BorderStyle     =   0  'None
            Height          =   750
            Left            =   0
            ScaleHeight     =   750
            ScaleWidth      =   8535
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   0
            Width           =   8535
            Begin VB.TextBox txtVisitDate 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00AAAAAA&
               Height          =   340
               Left            =   1200
               TabIndex        =   26
               Text            =   "01/02/2034"
               Top             =   360
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.CommandButton cmdCalendar 
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
               Left            =   3750
               Picture         =   "frmEFormDataEntry.frx":074C
               Style           =   1  'Graphical
               TabIndex        =   25
               TabStop         =   0   'False
               ToolTipText     =   "Click to display a calendar"
               Top             =   0
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.Label lblEFormDate 
               BackStyle       =   0  'Transparent
               Caption         =   "eForm date:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00AAAAAA&
               Height          =   375
               Left            =   4320
               TabIndex        =   36
               Top             =   450
               Width           =   1095
            End
            Begin VB.Label lblVisitDate 
               BackStyle       =   0  'Transparent
               Caption         =   "Visit date:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00AAAAAA&
               Height          =   375
               Left            =   60
               TabIndex        =   35
               Top             =   450
               Width           =   975
            End
            Begin VB.Label lblEForm 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "eForm:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00AAAAAA&
               Height          =   255
               Left            =   4320
               TabIndex        =   28
               Top             =   75
               Width           =   4065
            End
            Begin VB.Label lblVisit 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Visit:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00AAAAAA&
               Height          =   255
               Left            =   60
               TabIndex        =   27
               Top             =   75
               Width           =   4185
            End
         End
         Begin VB.PictureBox cmdNextButton 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   7200
            Picture         =   "frmEFormDataEntry.frx":084E
            ScaleHeight     =   180
            ScaleWidth      =   210
            TabIndex        =   23
            Top             =   4080
            Width           =   210
         End
         Begin VB.PictureBox picNoteStatus 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   0
            Left            =   1800
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   2880
            Visible         =   0   'False
            Width           =   240
         End
         Begin MACROOptionGroup30.OptionGroup optCRFGroup 
            Height          =   375
            Index           =   0
            Left            =   3570
            TabIndex        =   18
            Top             =   1350
            Visible         =   0   'False
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   661
         End
         Begin VB.PictureBox picQuGroup 
            Height          =   1995
            Index           =   0
            Left            =   5160
            ScaleHeight     =   1935
            ScaleWidth      =   3375
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1380
            Visible         =   0   'False
            Width           =   3435
            Begin MSComCtl2.FlatScrollBar hsbQuGroup 
               Height          =   240
               Index           =   0
               Left            =   90
               TabIndex        =   22
               Top             =   1680
               Width           =   2805
               _ExtentX        =   4948
               _ExtentY        =   423
               _Version        =   393216
               Arrows          =   65536
               Orientation     =   1245185
            End
            Begin MSComCtl2.FlatScrollBar vsbQuGroup 
               Height          =   1095
               Index           =   0
               Left            =   2880
               TabIndex        =   16
               Top             =   240
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   1931
               _Version        =   393216
               Orientation     =   1245184
            End
            Begin VB.PictureBox picFixed 
               Height          =   1395
               Index           =   0
               Left            =   60
               ScaleHeight     =   1335
               ScaleWidth      =   2775
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   240
               Visible         =   0   'False
               Width           =   2835
               Begin VB.PictureBox picHeaders 
                  Height          =   255
                  Index           =   0
                  Left            =   180
                  ScaleHeight     =   195
                  ScaleWidth      =   2355
                  TabIndex        =   21
                  TabStop         =   0   'False
                  Top             =   60
                  Width           =   2415
               End
               Begin VB.PictureBox picScroller 
                  Height          =   795
                  Index           =   0
                  Left            =   120
                  ScaleHeight     =   735
                  ScaleWidth      =   2415
                  TabIndex        =   15
                  TabStop         =   0   'False
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   2475
               End
            End
         End
         Begin VB.CommandButton cmdPopupList 
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
            Left            =   900
            Picture         =   "frmEFormDataEntry.frx":0BBE
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   1320
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdAttachment 
            Caption         =   "Attach file ..."
            Height          =   375
            Index           =   0
            Left            =   1560
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   1860
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.OptionButton optGraphical 
            Caption         =   "optGraphical"
            Height          =   375
            Index           =   0
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   1830
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtCRFElement 
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
            TabIndex        =   0
            Top             =   1350
            Visible         =   0   'False
            Width           =   825
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   4350
            Top             =   2520
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            FontName        =   "MS Sans Serif"
         End
         Begin MSMask.MaskEdBox mskCRFElement 
            Height          =   375
            Index           =   0
            Left            =   3360
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1890
            Visible         =   0   'False
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   661
            _Version        =   393216
            HideSelection   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "-"
         End
         Begin MSACAL.Calendar calCRFElement 
            Height          =   1995
            Left            =   2160
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   2490
            Visible         =   0   'False
            Width           =   2055
            _Version        =   524288
            _ExtentX        =   3625
            _ExtentY        =   3519
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
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
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
         Begin MSComctlLib.ListView lvwPopupList 
            Height          =   435
            Left            =   2160
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   767
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Value"
               Object.Width           =   4940
            EndProperty
         End
         Begin VB.PictureBox picCRFElement 
            BackColor       =   &H80000014&
            BorderStyle     =   0  'None
            FillColor       =   &H00FFFFFF&
            Height          =   615
            Index           =   0
            Left            =   120
            ScaleHeight     =   615
            ScaleWidth      =   1815
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   2430
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Image imgDictionary 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   0
            Left            =   6420
            Stretch         =   -1  'True
            Top             =   3540
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Image imgSDVStatus 
            Height          =   195
            Index           =   0
            Left            =   840
            Stretch         =   -1  'True
            Top             =   3600
            Width           =   435
         End
         Begin VB.Image imgChangeCount 
            Height          =   240
            Index           =   0
            Left            =   1170
            Stretch         =   -1  'True
            Top             =   3270
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Label lblNRCTC 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "H4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   0
            Left            =   1410
            TabIndex        =   11
            Top             =   3360
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.Image imgStatus 
            Height          =   240
            Index           =   0
            Left            =   810
            Stretch         =   -1  'True
            Top             =   3240
            Width           =   240
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
            Left            =   120
            MouseIcon       =   "frmEFormDataEntry.frx":11D8
            TabIndex        =   2
            Top             =   1380
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00000000&
            Index           =   0
            Visible         =   0   'False
            X1              =   30
            X2              =   9270
            Y1              =   3930
            Y2              =   3930
         End
      End
      Begin VB.PictureBox picHeaderExt 
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         Height          =   750
         Left            =   9120
         ScaleHeight     =   750
         ScaleWidth      =   195
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   120
         Width           =   195
      End
   End
   Begin MSComCtl2.FlatScrollBar vsbCRFPage 
      Height          =   4305
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   7594
      _Version        =   393216
      Orientation     =   1245184
   End
End
Attribute VB_Name = "frmEFormDataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'----------------------------------------------------------------------------------------'
'   File:       frmEFormDataEntry.frm
'   Copyright:  InferMed Ltd. 2001-2007. All Rights Reserved
'   Author:     Nicky Johns, August 2001
'   Purpose:    eForm data entry for Data Management in MACRO 2.2.
'               Based on original frmCRFDataEntry from MACRO 2.1
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
' REVISIONS:
'   NCJ 7 Aug 01 - Initial development based on original frmCRFDataEntry
'   ZA Aug 01 - Work on LostFocus vs. Validate events.
'   NCJ 30-31 Aug 01 - Decided to use LostFocus events
'           Further implementation of functionality
'   NCJ 18-20 Sep 01 - Making option & push buttons work
'   NCJ 21 Sep 01 - cmdClear, ChangeStatus
'   TA 21/9/01: Calls refresh schedule form when unloading
'   TA 24/09/2001: Response can be passed for when this is called from DataBrowser.
'   TA 25/9/01: Tidied up saving and allowed saving to be called externally
' NCJ 26/9/01 - Worked on right-mouse (status image) menu
' NCJ 1 Oct 01 - Check Arezzo tasks in SaveResponses
' NCJ 2 Oct 01 - Got Units dialog working
' NCJ 3 Oct 01 - Labs and lab test questions
' TA 04/10/2001 - Changes for authorising questions
' NCJ 4 Oct 01 - Removed unused tabCRF
' TA 05/10/01: CommandArrowDownChangeStatus changed from checkbox to special control
' TA 08/10/01: moFormBuilder.SetPageSize called explicitly after form is shown to prevent form sizing problem
'   NCJ 10 Oct 01 - Call ShowLabCode when changing Lab
' DPH 12/10/2001 - Stop user creating/saving/opening eForms if no permission to do so
'                - Reject Messages call frmWarnOverrule rather than MessageBox
' NCJ 15 Oct 01 - Removed unused lst and lstIndex combo boxes
'               Rationalised tab stops
' NCJ 23 Oct 01 - Incorporated new OptionGroup control
' ZA 23 Oct 01 - New IMedControls added (instead of OptionGroup control)
' NCJ 24 Oct 01 - Fixed operation of new optCRFGroup
' MACRO 3.0
' NCJ 1 Nov 01 - Changed MACROUSBS22 to MACROUSBS30
' NCJ 7 Nov 01 - Patched in ATO changes for Form_KeyDown
' NCJ 7-14 Nov 01 - Major changes required for support of RGQs
'03/12/2001: VTRACK Changes  - Text boxes have windows background
' TA 5/12/2001: Check for unvalidated data before allowing rightclick on image status
' NCJ 2 Jan 02 - Sorted out refocussing after viewing Audit Trail
' NCJ 7 Jan 02 - Include optional Response in UpdateResponses
' NCJ 8 Jan 02 - Make sure group status updated when any member response's status updated
' TA 14/1/02 - Added routines ready for scrolling to questions that get the focus - not yet implemented because of changes since 2.2
' MLM 23/01/02: Rolled forward changes from 2.2:
'               Give attached files unique filenames and make them read only
'               Do horizontal scrolling.
' NCJ 1 Feb 02 - Added ScrollIfOffScreen calls
' NCJ 11 Mar 02 - Added Response Rpt No to Display
'
' NCJ 8-9 July 02 - Brought forward all 2.2 changes since 1 Nov 01
'                   EXCEPT Multi-User and Sequetial Data Entry
' NCJ 12 Jul 02 - CBB 2.2.19/43 Don't ask for RFC when changing status of "new" responses
' NCJ 5-12 Aug 02 - Changes for revalidation and RFCs
' MLM 30/08/02: Check for form date in SaveResponses instead of Display()
' MLM 30/08/02: txtCRFElement_Got/LostFocus changed use new background colours and enabling.
' MLM 12/09/02: In many places, used EFormItems instead of Elements, Responses and row numbers.
'   Don't find Responses from moEFI because the response might belong to the visit eform;
'   find from the EFormItem instead.
' MLM 16/09/02: Added revalidation for visit eforms.
' NCJ 16/9/02 - Do revalidation AFTER displaying eForm in Display
' NCJ 17 Sept 02 - Do not store EFI as object but instead store its EFormTaskId
'               Replaced moEFI with form property EFI
' NCJ 18 Sept 02 - Changes to cope with multiuser and new locking model
'               Do not ask user to confirm Cancel or confirm Save unless they did something (see mbUserDidSomething)
' NCJ 23 Sept 02 - Check for requested eForms with locked VisitEFIs in SetUpNewEFI
'   TA 26/09/02: Changes for New UI - no title bar, not maximised etc
' NCj 26 Sept 02 - Multi User bug fixing
' RS 01/10/2002 - Added Timezone support
' TA 02/10/2002: IMEDNow used instead of Now in NewMIMessage
'   TA 04/10/2002: Removed dtatiem panel and added consequently removed functionality to
'       thE right click on the image status
' TA 07/10/2002: Reorganised some of display code - aiming for subject specific stuff eFormBuilder
'                   and general stuff in frmeFormDataEntry
' NCJ 9-10 Oct 02 - Added RQG horizontal scrolling ; tidied up right-mouse menu
' TA 15/10/2002: More changes to the User Interface
' TA 15/10/2002: Added enum and method for eFormActions
' NCJ 18 Oct 02 - Corrections to right mouse menu; added call to ChangeDoneSDVsToPlanned in SaveResponses
' NCJ 30 Oct 02 - Ensure Delete key works!
' NCJ 31 Oct 02 - Show warning popups for derived questions
'   TA 01/11/2002 - Changed to new icons
' NCJ 8-11 Nov 02 - Added hotlink handling
' NCJ 27 Nov 02 - Corrected bug in message in GoToEFI; force resize in SetupNewEFI (to get scroll bars right)
' NCJ 3 Dec 02 - Moved setting of eForm width from Display to eFormBuilder.InitEFI
' TA 03/01/2001: Show MIMessage lists for single question from eForm
' NCJ 6 Jan 03 - Do not worry about blank visit/form dates (just don't save) if user hasn't touched eform
' TA 17/10/2003: do not jump to hidden questions if launched from databrowser
' NCJ 17 Jan 03 - Offer to move to first form of next visit
'               Added cmdGoBack to go back to last eForm visited
'   TA 19/01/2003: Subject locking now done when creating and editing mimessages
'   NCJ 20 Jan 03 - Fixed bug in txtCRFElement_Change to only set DirtyFlag for enterable fields
'   NCJ 21 Jan 03 - Changed GoBack button to picBackButton
' MLM 23/01/03: Made picHeaderBar higher, and added form and visit date labels.
' NCJ 24 Jan 03 - Changed hourglass during revalidation;
'               don't prompt about missing form dates etc. in ClosedSuccessfully if they haven't actually done anything
' RS 05/02/03 - Adjust height of active dropdown item (if any) in ScrollIfOffScreen
' NCJ 10 Feb 03 - Fixed bug to do with mbDirtyField in ValidateProposedValue
'               When moving to a new eForm in GoToEFI, check the user is allowed to open it
' RS  11 Feb 03 - Corrected bug in dropdown resize when page scrolled vertically (if control offscreen)
' TA    17/02/2003: added new SDV icons
' RS    18/02/2003: Truncate strings that do not fit in the textbox on the eForm. Make sure that the original
'                   value is restored when needed in gotfocus and lostfocus events. (See moFormBuilder.TruncateTextBoxContent)
' NCJ 18 Feb 03 - Preserve scroll bar position around ElementRightMouseMenu
' NCJ 19 Feb 03 - Sorted out bugs in ScrollIfOffScreen; added RefreshResponses for Toby
' NCJ 20 Feb 03 - Implemented automatic horizontal scrolling
' NCJ 5 Mar 03 - Fixed yet another bug to do with mbDirtyField
' NCJ 24 Mar 03 - Disable user intervention during SetUpNewEFI and SaveResponses
'               Refresh Subject properties before doing Revalidation
' NCJ 26 Mar 03 - Added saving of patient data when AREZZO tasks are done in DoSaveResponses
' MLM 27/03/03: 3.0 bugs 782, 1375: Modified ActiveControlIsAButton and renamed to FocusToNextFieldOK
'   3.0 bug 1360: Set mbMouseDownInNextField = False in optCRFGroup_LostFocus, optGraphical_lostfocus so that FocusToAppropriateField fires correctly.
' NCJ 31 Mar 03 - Sorted out eForm cancellation when there's no visit/form date (BUG 1431)
' NCJ 10 Apr 03 - Bug 1532 - Extra argument in SaveResponses to sort out bug from previous bug fix
' NCJ 23 Apr 03 - Implemented F1 key for MACRO 3.0 Help
' NCJ 6 May 03 - Format calendar control dates according to LocalDateFormat (Bug 1673)
' NCJ 22 May 03 - Changed adjustment for popup lists in ScrollIfOffScreen (may still be problematic...!)
' NCJ 18 Jun 03 - Added check for RegisterSubject permission
' NCJ 2nd Oct 03 - Added Randomisation Code conditionally compiled if RANDOM = 1
' NCJ 7 Jan 04 - Removed conditional compilation from Randomisation code
' NCJ 30 Jun 04 - Make sure we check an existing lab's site association in AskLabCode
'               and call AskLabCode before revalidating
' NCJ 8 July 04 - Always update responses after revalidation (bug 2324)
' NCJ 9 Nov 04 - Added in conditionally compiled code for ORAMA
' NCJ 9 Nov 04 - Bug 2362 - Don't accept F7 etc. while we're showing the audit trail or question def
' NCJ 15 Nov 04 - Bug 2463 - Reset mbDirtyField in cmdAttachment_GotFocus
' NCJ 22 Nov 04 - Issue 2451 - Do not check SDV status when changing Done SDVs in DoSaveResponses
' NCJ 13 Jan 05 - Fix for bug 2495 - Some rolling back of stuff which fixed 2324
' NCJ 18 Jan 05 - Issue 2502 - Allow keyboard selection of sub lists of categories (list filtering)
' MLM 04/07/05: bug 2464: When launching the MIMessage browser, specify the visit, eForm and question cycle that we want.
' ic 19/07/2005 added clinical coding
' NCJ 11 Jan 06 - Issue 2619 - Must check eForm is complete before doing Registration
' NCJ 26 Jan 06 - Bug 2636 - Evaluate Hotlinks AFTER saving responses
' NCJ 11 Jul 06 - Bug 2761 - Check for null responses in ClosedSuccessfully
' ic 21/11/2006 - bug 2843 - pass the current username in to setcodingstatus and setcodingdetails
' NCJ 18 Dec 06 - Bug 2850 - Deal with empty drop-down lists
' NCJ 29 Dec 06 - Bug 2861 - Avoid creating MIMsg for non-existent responses (i.e. in last empty row of RQG)
' NCJ 28 Mar 07 - Bug 2793 - Check current field when user presses f10
' NCJ 28 Mar 07 - Bug 2896 - Keep dirty flag set for rejected text fields
' ic 28/07/2005 - removed clinical coding 'enter cc response' permission
'-------------------------------------------------------------------------------------------

Option Explicit
Option Compare Binary
Option Base 0

' Possible randomisation results
Private Enum RandomisationResult
    rrSuccess = 0
    rrCantStratify = 1
    rrNoTreatments = 2
    rrError = 3
End Enum

'actions that can be performed in an eForm (used by eFormAction method)
' NCJ 8 Nov 02 - Added eaGotoHotlink
' NCJ 17 Jan 03 - Added eaGoBack (i.e. where we've just come from)
Public Enum eEFormAction
    eaCancel = 0
    eaReturn = 1
    eaPrint = 2
    eaPrevious = 3
    eaNext = 4
    eaSave = 5
    eaGotoThisEFormInVisit = 6
    eaGotoEFormInThisVisit = 7
    eaGotoHotlink = 8
    eaGoBack = 9
End Enum

' Store which pushbutton they clicked on
Private mnMouseDownIndex As Integer

Private mbMouseDownOnPopupList As Boolean
Private mbMouseDownOnCalendar As Boolean
Private mbMouseDownInNextField As Boolean

' NCJ 18 Jan 05 - Remember when user presses arrow keys in dropdown
Private mbArrowInDropDown As Boolean

' NCJ 17 Feb 00 SRs 3003,3004,3012 etc.
' Store when we're changing forms etc. and don't allow other events to interrupt us
Private mbDontInterrupt As Boolean
Private mbDuringLoad As Boolean

Private mbDirtyField As Boolean
' The element we're validating (to avoid doing the same one twice)
Private mnValidatingResponse As Integer

' MLM 21/05/02
Private mbCRFPageClicked As Boolean

' Our "form builder"
Private moFormBuilder As EFormBuilder

Private moEForm As eFormRO
'Private moEFI As EFormInstance
Private mlEFormTaskId As Long
Private mcolRFCs As Collection
Private moUser As MACROUser
Private WithEvents moSubject As StudySubject
Attribute moSubject.VB_VarHelpID = -1

Private moLastSelectedControl As Control

' NCJ 18 Sept 02 - Remember if the user did something since the last Save
Private mbUserDidSomething As Boolean

' NCJ 17 Jan 03 - Remember Taskid of last eForm (so we can return to it)
Private mlPreviousEformTaskId As Long
Private msPreviousEformName As String

'events to notify external listeners that this form has had its EFI loaded/changed/unloaded
    ' TA/NCJ 19 Feb 03 - Don't bother with EFIUnload event
'Public Event EFIUnLoad(lStudyId As Long, sSite As String, lSubjectId As Long, lEFIId As Long)
Public Event EFILoad(lStudyId As Long, sSite As String, lSubjectId As Long, lEFIId As Long)
Public Event EFIChange(lStudyId As Long, sSite As String, lSubjectId As Long, lEFIId As Long)
Public Event Unload()

'store whether subject label has changed
Private mbSubjectLabelChanged As Boolean

' Store whether we've sorted out the lab
Private mbLabIsSorted As Boolean

Public Event SubjectLabelChanged()

' NCJ 19 Jan 05 - Store "enable list filtering" setting
Private mbEnableListFiltering As Boolean

'ic 27/07/2005 holds form-wide collection of installed dictionaries
Private moDictionaries As MACROCCBS30.Dictionaries

'---------------------------------------------------------------------
Public Function Display(oUser As MACROUser, oEFI As EFormInstance, _
                            lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long, _
                            ofrmeFormTop As frmWebBrowser, ofrmeFormLh As frmWebBrowser, _
                            Optional lResponseTaskId As Long = -1, _
                            Optional nResponseRptNo As Integer = 1, _
                            Optional bMaximised As Boolean = True) As Boolean
'---------------------------------------------------------------------
' Display all the elements on the eForm for the given EFI.
' Response lResponseTaskId to get the focus (-1 for first available question).
' NCJ 11 Mar 02 - Added optional nResponseRptNo
' MLM 30/08/02: Removed call to CheckFormDate (this will be done when saving the form)
' MLM 10/09/02: Focus to the form or visit date sometimes
' MLM 13/09/02: Return FALSE if oEFI in a VisitEFormInstance.
' NCJ 3 Dec 02 - eFormWidth is now set up in eFormBuilder.InitEFI instead of here
'---------------------------------------------------------------------
Dim oItem As EFormItem
Dim oResponse As Response
'Dim lEFormWidth As Long

'MLM 10/09/02:
Dim bFocussed As Boolean

' NCJ 3 Dec 02 - See eFormBuilder.InitEFI
''25/10/2001: VTRACK Changes ref 1.2,1.3
'#If VTRACK = 1 Then
'    lEFormWidth = glLANDSCAPE_WIDTH
'#Else
'    ' NCJ 4 Nov 02 - We read the width that's been set in SD
'    lEFormWidth = oEFI.eForm.eFormWidth
'    If lEFormWidth = NULL_LONG Then
'        lEFormWidth = glPORTRAIT_WIDTH
'    End If
'#End If
    
    
    'initiise the flag to tell whether subject label has changed
    mbSubjectLabelChanged = False
    
    
    'MLM 13/09/02: Return FALSE if oEFI in a VisitEFormInstance.
    If Not oEFI.VisitInstance.VisitEFormInstance Is Nothing Then
        If oEFI.eForm.EFormId = oEFI.VisitInstance.VisitEFormInstance.eForm.EFormId Then
            
            Display = False
            Exit Function
        End If
    End If
    
    Load Me
    
    Me.Visible = False
        
    HourglassOn
            
    'za added
    Set moUser = oUser
    
    With oEFI.eForm.Study
        Set moSubject = .Subject
        Set mcolRFCs = .RFCs
    End With
    
    If Not moFormBuilder Is Nothing Then
        Call moFormBuilder.Terminate
    End If
    
    With Me
        .WindowState = vbNormal
        .Left = lLeft
        .Top = lTop
        .Width = lWidth
        .Height = lHeight

        .BackColor = eMACROColour.emcBackground

        .picCRFPage.Top = 0
        .picCRFPage.Left = 0
        .picCRFPage.BorderStyle = 0
                 
        'set form width (now done in eFormBuilder.InitEFI)
'        .picCRFPage.Width = lEFormWidth

    End With
    
    Set moFormBuilder = New EFormBuilder
    Call moFormBuilder.Init(ofrmeFormTop, ofrmeFormLh)

     
    ' eForm specific
    ' Set up eForm but don't do revalidation until later
    If Not SetUpNewEFI(oEFI, False) Then
        ' We couldn't get this eForm's responses
        Display = False
        HourglassOff
        Unload Me
        Exit Function
' EXIT FUNCTION HERE
    End If

    'customized load event notification has to be here rather than in the form_load event
    ' because we don't have an moEFI in the load event
    'notify external listeners that this form has laoded with an EFI
    RaiseEvent EFILoad(oEFI.eForm.Study.StudyId, _
            oEFI.VisitInstance.Subject.Site, oEFI.VisitInstance.Subject.PersonId, oEFI.EFormTaskId)
        
        
    Me.Show vbModeless 'this should force as resize
    
    HourglassOff
    
    ' NCJ 16 Sept 02 - Let the form be displayed and do the revalidation
    DoEvents
    Call RevalidateMe(False)

    Display = True
    
    ' Set the focus if required
    'Focus is set:
    '1. to the element specified when this function was called, if there is one
    If lResponseTaskId > -1 Then
        ' NCJ 11 Mar 02 - Include given RptNo
        Set oResponse = EFI.Responses.ResponseByResponseId(lResponseTaskId, nResponseRptNo)
        If Not oResponse Is Nothing Then
            If Not oResponse.Element.Hidden Then
                'TA 17/10/2003: only show if not hidden
                Set oItem = moFormBuilder.EFormItemByIndex(moFormBuilder.IndexForElementRow(oResponse.Element, nResponseRptNo))
                Call moFormBuilder.ResetFormFocus(oItem)
                Set oResponse = Nothing
                Set oItem = Nothing
                Exit Function
            End If
        End If
    End If
    '2. to the first other enterable element, if there is one or to visit or eformdata if relevant
    Call moFormBuilder.FocusToNextField(Nothing, 1)
    '5. otherwise, it will be the 'next form' button that has the focus

    Set oResponse = Nothing

    'ebusre we are on top
    ofrmeFormLh.ZOrder
    ofrmeFormTop.ZOrder
    Me.ZOrder
Exit Function

ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Display", Err.Source) = Retry Then
        Resume
    End If

End Function

'---------------------------------------------------------------------
Private Function SetUpNewEFI(oEFI As EFormInstance, Optional bRevalidate As Boolean = True) As Boolean
'---------------------------------------------------------------------
' Initialise for new EFI
' including loading the Responses
' and drawing the new eForm elements
' NCJ 16 Sept 02 - Added bRevalidate which says whether we want revalidation to happen
' NCJ 18 Sept 02 - Return result of LoadResponses
' NCJ 23 Sept 02 - Check for requested eForms with locked VisitEFIs
'new hourglass form displayed
' NCJ 24 Mar 03 - Use new ToggleUserIntervention while we're setting up the form
'---------------------------------------------------------------------
Dim sLockErrMsg As String
Dim sEFILockToken As String
Dim sVEFILockToken As String
Dim enLoadResponsesResult As eLoadResponsesResult

    On Error GoTo ErrLabel
        
    HourglassOn
    
    frmHourglass.Display "Processing eForm", True
    
    ' Don't let them click anywhere and mess things up
    Call ToggleUserIntervention(False)
    
    SetUpNewEFI = False
    
    ' We don't need to hold onto the EFILock Token here as the EFI does
    ' And same for VisitEFI token
    enLoadResponsesResult = moSubject.LoadResponses(oEFI, sLockErrMsg, sEFILockToken, sVEFILockToken)
    
    If enLoadResponsesResult = lrrCouldNotLockForSave Then
        ' Couldn't get at the eForm
        UnloadfrmHourglass
        ' TODO give more explanation?
        DialogWarning sLockErrMsg
        
    ' NCJ 23 Sept 02 - Check if eForm is requested
    ' but its Visit eForm is locked
    ElseIf Not VisitEFIIsAvailable(oEFI) Then
        ' Message already displayed
        UnloadfrmHourglass
        ' Must remove the responses because we're not going to let them go any further
        Call moSubject.RemoveResponses(oEFI, True)
    Else

        Set moEForm = oEFI.eForm
        ' NCJ 17 Sept 02 - Store eFormTaskId rather than EFI itself
        ' Set moEFI = oEFI
        mlEFormTaskId = oEFI.EFormTaskId
        
        Set moLastSelectedControl = Nothing
        
        ' User hasn't done anything yet
        mbUserDidSomething = False
        
        ' NCJ 30 Jun 04 - Lab isn't checked yet
        mbLabIsSorted = False
        
        Call moFormBuilder.InitEFI(moUser, Me, moSubject, oEFI)
        
        ' NCJ 21 Jan 03 - Always enable the GoBack button
        picBackButton.Enabled = True
        If mlPreviousEformTaskId > 0 Then
            picBackButton.TooltipText = "Go back to " & msPreviousEformName
        Else
            picBackButton.TooltipText = "Return"
        End If
        
        ' NCJ 21 Feb 03 - Must also reset the picCRFPage
        ' in case it was scrolled to an unsuitable position on the previous eForm
        With Me
            .picCRFPage.Top = 0
            .picCRFPage.Left = 0
            .vsbCRFPage.Value = 0
            .hsbCRFPage.Value = 0
        End With
        
        ' NCJ 27 Nov 02 - Force a resize to get the scroll bars right
        ' in case this eForm is a different size from the last one
        Call Form_Resize
        
         ' NCJ 13 Jan 05 - Moved to after Revalidation (bug 2495)
'        ' Finally set the focus to the first available element
'        Call moFormBuilder.FocusToNextField(Nothing, 1)
        
        UnloadfrmHourglass
        
        If bRevalidate Then
            ' Let the eForm be shown (must call DoEvents) and then do revalidation
            DoEvents
            ' NCJ 24 Jan 03 - Call Hourglass Suspend/Resume instead of Off/On
            Call HourglassSuspend
            Call RevalidateMe(False)
            Call HourglassResume
        End If
        
        ' NCJ 13 Jan 05 - Moved here from before Revalidation (bug 2495)
        ' Finally set the focus to the first available element
        Call moFormBuilder.FocusToNextField(Nothing, 1)
        
        'let's always store this
        Call goUser.UserSettings.SetSetting(SETTING_LAST_USED_EFORM, oEFI.GetAllSubjectsKey)
        
        SetUpNewEFI = True
    End If
    
    ' Now let them back in
    Call ToggleUserIntervention(True)
    
    HourglassOff

Exit Function
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "SetUpNewEFI", Err.Source) = Retry Then
        ' Let them back in
        Call ToggleUserIntervention(True)
        Resume
    End If

End Function

'---------------------------------------------------------------------
Private Function VisitEFIIsAvailable(oEFI As EFormInstance) As Boolean
'---------------------------------------------------------------------
' NCJ 23 Sept 02
' Returns FALSE if EFI is requested
' and there's no visit date
' and the visit date is being edited by someone else,
' otherwise returns TRUE.
' Assume VisitEForm's Responses have already been loaded
'---------------------------------------------------------------------
Dim oVEFI As EFormInstance
Dim sMsg As String

    VisitEFIIsAvailable = True
    
    ' Don't bother checking for non-Requested eForms (assume they have a valid Visit Date)
    If oEFI.Status <> eStatus.Requested Then Exit Function
    
    ' If there's a visit date then we're OK
    If oEFI.VisitInstance.VisitDate <> 0 Then Exit Function
    
    Set oVEFI = oEFI.VisitInstance.VisitEFormInstance
    If Not oVEFI Is Nothing Then
        ' There is a Visit eForm
        If Not oVEFI.eForm.EFormDateElement Is Nothing Then
            ' There is a visit date - if it's not derived, see if the Visit EFI is read-only
            If oVEFI.ReadOnly And (oVEFI.eForm.EFormDateElement.DerivationExpr = "") Then
                'TA 23/0/12003: hidehourglass form - macro3.0 bugs and issues no. 679
                UnloadfrmHourglass
                sMsg = "Another user is entering the visit date and new eForms cannot be opened in this visit"
                DialogWarning sMsg
                VisitEFIIsAvailable = False
            End If
        End If
        
        Set oVEFI = Nothing
    End If

End Function

'---------------------------------------------------------------------
Private Sub calCRFElement_Click()
'---------------------------------------------------------------------
' Click on the calendar
' Store the date in the associated text field and validate it
' NCJ 6 May 03 - Format date according to Response's LocalDateFormat
'---------------------------------------------------------------------
Dim oResponse As Response
Dim oItem As EFormItem

    ' Get the current field
    Set oItem = moFormBuilder.CurrentItem

    mbMouseDownOnCalendar = True
    ' Hide the calendar
    calCRFElement.Visible = False
    ' Get the associated Response
    Set oResponse = oItem.Response
    ' Store the selected date in the text field
    ' NCJ 6 May 03 - Use oResponse.LocalDateFormat
    ' Need to replace "/" with "\/" to prevent VB from converting "/" to the machine's regional separator
    txtCRFElement(oItem.ControlIndex).Text = Format(calCRFElement.Value, _
                                Replace(oResponse.LocalDateFormat, "/", "\/"))
    Call txtCRFElement_LostFocus(oItem.ControlIndex)

    Set oResponse = Nothing
    Set oItem = Nothing

End Sub

'---------------------------------------------------------------------
Private Sub calCRFElement_LostFocus()
'---------------------------------------------------------------------
' Calendar control lost focus
'---------------------------------------------------------------------
Dim oItem As EFormItem

    ' Get the current field
    Set oItem = moFormBuilder.CurrentItem
    
    calCRFElement.Visible = False
    Call txtCRFElement_LostFocus(oItem.ControlIndex)

    ' NCJ 2/4/01 - Check that we're at the front (and it's not a supplementary dialog)
    If Not mbMouseDownOnCalendar And Screen.ActiveForm.Name = Me.Name Then
        Call moFormBuilder.ResetFormFocus(oItem)
    End If
    mbMouseDownOnCalendar = False

    Set oItem = Nothing

End Sub

'---------------------------------------------------------------------
Private Sub cmdAttachment_Click(Index As Integer)
'---------------------------------------------------------------------
' Click on multimedia button
' MLM 23/01/02: Rolled forward from 2.2: Save files to Documents folder
'               with unique file names and read only.
'---------------------------------------------------------------------
Dim sMediaFile As String
Dim oItem As EFormItem
Dim bStayOnField As Boolean

    'TA 09/04/2002 restart timeout timer at appropriate intervals
    Call RestartSystemIdleTimer

    Set oItem = moFormBuilder.EFormItemByIndex(Index)

    ' NCJ 7 Jun 02 (CBB2.2.12/41) Assume we want to stay put
    bStayOnField = True
    
    ' Are they viewing it or changing it?
    If cmdAttachment(Index).Caption = gsATTACH_FILE Then
       ' They want to change their attachment
       Call ChangeAttachment(oItem)
    Else
        ' They want to view their attachment
        If Right(cmdAttachment(Index).Tag, 3) = "gif" _
        Or Right(cmdAttachment(Index).Tag, 3) = "jpg" _
        Or Right(cmdAttachment(Index).Tag, 3) = "bmp" Then
            frmImageViewer.ShowImage gsDOCUMENTS_PATH & cmdAttachment(Index).Tag
        Else
            ShowDocument Me.hWnd, gsDOCUMENTS_PATH & cmdAttachment(Index).Tag
        End If
    
    End If
    
    Set oItem = Nothing

End Sub

'---------------------------------------------------------------------
Private Sub cmdCalendar_Click(Index As Integer)
'---------------------------------------------------------------------
' The calendar button - display the calendar control
' NCJ 6 May 03 - Bug fix for "local" date formats (Bug 1673)
'---------------------------------------------------------------------
Dim oResponse As Response
Dim oTextBox As TextBox
Dim oItem As EFormItem
Dim dtDate As Date

    ' Pick up the relevant element
    Set oItem = moFormBuilder.EFormItemByIndex(Index)
    Set oTextBox = txtCRFElement(Index)

    Set oResponse = oItem.Response
    
    ' Fill in the Question Information Panel
    Call moFormBuilder.ElementGotFocus(oItem)
    
    ' NCJ 20/12/00 SR4088 Store calendar's text value here
    oTextBox.Tag = oTextBox.Text
    
    ' NCJ 6 May 03 - Changed the way we display the current date value
    ' to ensure that we handle the element's formatting
    If oResponse.Value <> "" Then
        dtDate = moSubject.Arezzo.ArezzoDateToDouble(oResponse.ArezzoValue)
    Else
        dtDate = Now
    End If
    
'    If oTextBox.Text <> "" Then
'        calCRFElement.Value = oTextBox.Text
'    Else
        ' Set Y,M,D individually to prevent tomorrow being displayed
        calCRFElement.Year = DatePart("yyyy", dtDate)
        calCRFElement.Month = DatePart("m", dtDate)
        calCRFElement.Day = DatePart("d", dtDate)
'    End If
    
    ' NCJ 4 Nov 02 - No longer pass oElement to DisplayCalendar
    Call moFormBuilder.DisplayCalendar(Index)
    calCRFElement.SetFocus

    Set oResponse = Nothing
    Set oItem = Nothing

End Sub

'---------------------------------------------------------------------
Private Sub ChangeAttachment(oItem As EFormItem)
'---------------------------------------------------------------------
' Change the attachment
'---------------------------------------------------------------------
Dim nIndex As Integer
Dim sMediaFile As String
Dim bStayOnField As Boolean
Dim oFS As New Scripting.FileSystemObject
 
 
    If oItem Is Nothing Then Exit Sub
        
    nIndex = oItem.ControlIndex
    
    If oItem.Element.DataType = eDataType.Multimedia Then
        'omly multimedia
        On Error Resume Next
        CommonDialog1.CancelError = True
        CommonDialog1.Flags = cdlOFNExplorer + cdlOFNFileMustExist + cdlOFNFileMustExist
        CommonDialog1.ShowOpen
        If Err.Number <> cdlCancel Then
            'construct new file name
            sMediaFile = CreateGUID
            'add the same extension as the file the user selected, if there was one (so that it can still be displayed correctly)
            If InStr(CommonDialog1.FileTitle, ".") > 0 Then
                sMediaFile = sMediaFile & Mid(CommonDialog1.FileTitle, InStrRev(CommonDialog1.FileTitle, "."))
            End If
            'MLM 18/06/03: Only validate the file name if the file could be copied to the documents folder
            On Error Resume Next
            oFS.CopyFile CommonDialog1.FileName, gsDOCUMENTS_PATH & sMediaFile
            If Err.Number <> 0 Then
                'copy failed
                DialogError "The selected file could not be attached due to the following error:" & vbCrLf & _
                    Err.Number & ": " & Err.Description & "." & vbCrLf & _
                    "Please contact your MACRO systems administrator."
                Err.Clear
                On Error GoTo 0
            Else
                Err.Clear
                On Error GoTo 0
                If ValidateProposedValue(sMediaFile, oItem) Then
                    'FileCopy CommonDialog1.FileName, gsDOCUMENTS_PATH & sMediaFile
                    'MLM 18/4/02: Don't set file read-only as this spoils re-importing a file
                    'SetAttr gsDOCUMENTS_PATH & sMediaFile, vbReadOnly
                    cmdAttachment(nIndex).Tag = sMediaFile
                    ' Change the buton label
                    cmdAttachment(nIndex).Caption = gsVIEW_FILE
                    ' Move on to next element only if no Warning
                    bStayOnField = (oItem.Response.Status = eStatus.Warning)
                End If
            End If
        End If
        ' Now focus to where we want to be
        Call FocusToAppropriateField(oItem, bStayOnField)
        On Error GoTo 0
    End If

End Sub

'---------------------------------------------------------------------
Private Sub cmdChangeLab_Click()
'---------------------------------------------------------------------
' Change the Lab for this eForm
' NCJ 28 May 03 - Changed message wording for 1 lab, and added message for NO labs
'---------------------------------------------------------------------
Dim sMsg As String
Dim oLabs As clsLabs

    Set oLabs = New clsLabs
    Call oLabs.Load(moSubject.Site)
    If oLabs.Count = 1 Then
        sMsg = "There is only one laboratory available for this site."
        Call DialogInformation(sMsg, cmdChangeLab.Caption)
    ElseIf oLabs.Count = 0 Then
        sMsg = "There are no laboratories available for this site."
        Call DialogError(sMsg, cmdChangeLab.Caption)
    End If

    If oLabs.Count >= 1 Then
        Call ChangeLab
    End If
    Set oLabs = Nothing

End Sub

'---------------------------------------------------------------------
Private Sub ClearQuestion(oItem As EFormItem)
'---------------------------------------------------------------------
' Clear the current question
' Only clear if not already empty
' NCJ 28 Feb 03 - Must ensure question has a response and is enterable!
'---------------------------------------------------------------------

    If oItem Is Nothing Then Exit Sub
    
    If oItem.Response Is Nothing Then Exit Sub
    
    If Not oItem.Response.Enterable Then Exit Sub
    
    If oItem.Response.Value = "" Then Exit Sub
    
    ' Pass empty string as new value
    Call ValidateProposedValue("", oItem)
    
    ' NCJ 20/3/00 - Set focus back to same question
    DoEvents
    Call moFormBuilder.ResetFormFocus(oItem)


End Sub

'---------------------------------------------------------------------
Private Sub imgDictionary_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' ic 15/07/2005 thesaurus button - display browser
' revisions
' ic 21/11/2006 bug 2843 - pass the current username in to setcodingdetails
'---------------------------------------------------------------------
Dim oItem As EFormItem
Dim oDictionary As MACROCCBS30.Dictionary
Dim oTextBox As TextBox
Dim oResponse As Response
Dim sNewResponse As String
Dim sNewCodingDetails As String
Dim sText As String
Dim sError As String
Dim bSuccess As Boolean
Dim sRFC As String

    On Error GoTo ErrHandler
    
    'check that its a left click
    If Button <> vbLeftButton Then Exit Sub
    
    'get the item
    Set oItem = moFormBuilder.EFormItemByIndex(Index)
    Set oTextBox = txtCRFElement(Index)
    Set oResponse = oItem.Response
    
    'check for unvalidated data first
    If ThereIsInvalidData Then
        Exit Sub
    End If
    
    
    If (oResponse.CodingStatus = eCodingStatus.csDoNotCode) Then
        'dont allow questions marked as 'do not code' to be coded
        'Call DialogInformation("This question has been marked as 'Do not code', no coding can be added while in this state.")
        Exit Sub
        
    ElseIf (oResponse.Value = "") Then
        'dont allow responses where the verbatim has not been entered to be coded
        Call DialogInformation("A response without a value cannot be coded")
        
    Else
        HourglassOn
        
        If (moDictionaries Is Nothing) Then
            'get the dictionary for this question
            Set moDictionaries = New MACROCCBS30.Dictionaries
            Call moDictionaries.Init(SecurityDatabasePath) 'moUser.CurrentDBConString)
        End If
        
        'get the current response value and coded value
        sNewResponse = oResponse.Value
        sNewCodingDetails = oResponse.CodingDetails
    
        
        If (oResponse.CodingStatus = eCodingStatus.csAutoEncoded _
            Or oResponse.CodingStatus = eCodingStatus.csCoded _
            Or oResponse.CodingStatus = eCodingStatus.csValidated) _
            And Shift <> 2 Then
            
            'this question is coded - display the coded value
            'get the dictionary for this element
            Set oDictionary = moDictionaries.DictionaryFromVersion(oResponse.DictionaryName, oResponse.DictionaryVersion)
        
            If Not (oDictionary Is Nothing) Then
                'display the code
                If oDictionary.ToText(oResponse.CodingDetails, sText, sError) Then
                    Call DialogInformation(sText, oResponse.DictionaryName & " " & oResponse.DictionaryVersion)
                Else
                    Call DialogWarning("MACRO plugin '" & oItem.Element.DictionaryId & "' encountered errors :" & vbCrLf & sError)
                End If
            Else
                Call DialogWarning("The dictionary specified for this question was not found.")
            End If
    
        
        Else
            'this question is not coded - display the browser
            'get the dictionary for this element
            Set oDictionary = moDictionaries.DictionaryFromId(oItem.Element.DictionaryId)
        
            If Not (oDictionary Is Nothing) Then
                'encode by passing to the dictionary plugin
                Call oDictionary.Code(sNewResponse, sNewCodingDetails)
                
                'allow a maximum length for the coded value
                If (Len(sNewCodingDetails) < gnMAX_CODE_LENGTH) Then
                    'see if user has coded question
                    If (oResponse.CodingDetails <> sNewCodingDetails) Then
                        bSuccess = True
                        If (oResponse.RequiresCodingRFC) Then
                            HourglassOff
                            'Display a RFC window and ask user to select one from there
                            frmRFC.lblRFC.Caption = "Please enter or choose the reason for changing the code of " & oItem.Element.Name
                            frmRFC.DisplayRFC mcolRFCs, GetCoordForForm(oItem)
                            sRFC = frmRFC.GetRFC
                            If Len(sRFC) < 1 Then
                                bSuccess = False
                            End If
                            HourglassOn
                        Else
                            sRFC = ""
                        End If
                    
                        If bSuccess Then
                            'set the new coding status and refresh the image on the coding button
                            Call oResponse.SetCodingDetails(oDictionary.Name, oDictionary.Version, sNewCodingDetails, sRFC, _
                            goUser.UserName, goUser.UserNameFull)
                            Call moFormBuilder.SetClinicalCodingImage(CInt(eCodingStatus.csCoded), imgDictionary(Index))
                        End If
                    End If
                Else
                    Call DialogWarning("MACRO plugin '" & oItem.Element.DictionaryId & "' coded value too long")
                End If
            Else
                Call DialogWarning("The dictionary specified for this question was not found.")
            End If
        End If
        Set oDictionary = Nothing
        
    
        'set the value, if item modifiable and changed
        If Not oTextBox.Locked And oTextBox.Text <> sNewResponse Then
            ' Store the changed value in the text field and revalidate by calling the LostFocus event
            txtCRFElement(Index).Text = sNewResponse
            Call txtCRFElement_LostFocus(Index)
        End If
        
        'set focus back to where we were
        If Not oItem Is Nothing Then
            'ignore erros (form might be disabled)
            On Error Resume Next
            Call moFormBuilder.ResetFormFocus(oItem)
            
        End If
        
    End If
    
    Set oItem = Nothing
    Set oTextBox = Nothing
    Set oResponse = Nothing
    HourglassOff
    Exit Sub
    
ErrHandler:
    Call DialogWarning("MACRO plugin '" & oItem.Element.DictionaryId & "' encountered errors :" & vbCrLf & Err.Description)
    HourglassOff
End Sub

'---------------------------------------------------------------------
Private Sub cmdExpand_Click(Index As Integer)
'---------------------------------------------------------------------
' The expand button - display the text viewer
'---------------------------------------------------------------------
Dim oElement As eFormElementRO
Dim oTextBox As TextBox
Dim oItem As EFormItem
Dim sNewResponse As String

    ' Pick up the relevant element
    Set oItem = moFormBuilder.EFormItemByIndex(Index)
    Set oTextBox = txtCRFElement(Index)
    Set oElement = oItem.Element
    
    ' Fill in the Question Information Panel
    Call moFormBuilder.ElementGotFocus(oItem)
    
    ' Display/Edit Expanded Response
    sNewResponse = frmExpandResponse.Display(oTextBox, oElement.Caption)
    
    ' Set the value, if item modifiable and changed
    If Not oTextBox.Locked And oTextBox.Text <> sNewResponse Then
        ' Store the changed value in the text field and revalidate by calling the LostFocus event
        txtCRFElement(Index).Text = sNewResponse
        Call txtCRFElement_LostFocus(Index)
    Else
        ' Nothing changed - just stay where we are
        FocusToAppropriateField oItem, True
    End If
    
    Set oElement = Nothing
    Set oItem = Nothing

End Sub

'---------------------------------------------------------------------
Private Sub cmdNextButton_Click()
'---------------------------------------------------------------------
' "Next eForm" button (in fact in 3.0 it's a picture box!)
' NCJ 8 Jul 02 - Copied here from 2.2
'---------------------------------------------------------------------
    
    ' Check for unvalidated data first
    If ThereIsInvalidData Then
        Exit Sub
    End If
    
    Call GotoEFI(eaNext)

End Sub

'---------------------------------------------------------------------
Private Sub cmdNextButton_GotFocus()
'---------------------------------------------------------------------
' "Next eForm" button has the focus
' so clear question focus
'---------------------------------------------------------------------

    Call moFormBuilder.NoQuestionFocus
    cmdNextButton.BorderStyle = 1
    ' NCJ 19 Feb 03 - Scroll to the button if not visible
    If Not EFI.ReadOnly Then
        Call ScrollIfOffScreen(cmdNextButton)
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdNextButton_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------
' Intercept Return or Space to simulate button click
'---------------------------------------------------------------------

    Select Case KeyCode
    Case vbKeyReturn, vbKeySpace
        cmdNextButton_Click
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdNextButton_LostFocus()
'---------------------------------------------------------------------

    cmdNextButton.BorderStyle = 0

End Sub

'---------------------------------------------------------------------
Private Sub cmdPopupList_Click(Index As Integer)
'---------------------------------------------------------------------
' Toggle the state of the popuplist
'---------------------------------------------------------------------
Dim bPopuplistDown As Boolean
    
    bPopuplistDown = moFormBuilder.PopupListVisible
    
    txtCRFElement(Index).SetFocus
    DoEvents
    
    If bPopuplistDown Then
        Call moFormBuilder.HidePopupList
    End If

End Sub

'---------------------------------------------------------------------
Private Sub RemoveComment(oItem As EFormItem)
'---------------------------------------------------------------------
' Remove all comments for this question
'---------------------------------------------------------------------
Dim nElementId As Integer
Dim oElement As eFormElementRO
Dim oResponse As Response

    On Error GoTo ErrHandler
    
    If Not moUser.CheckPermission(gsFnAddIComment) Then Exit Sub
    If oItem Is Nothing Then Exit Sub
          
    'TA ensure we have permission, a response, and unlocked
    If oItem.Response Is Nothing Then Exit Sub

          
    Set oElement = oItem.Element
    Set oResponse = oItem.Response ' moEFI.Responses.ResponseByElement(oElement, oItem.RowNo)
    

    If oResponse.LockStatus <> LockStatus.lsUnlocked Then Exit Sub
    
    
    ' Check there are comments first
    If Len(oResponse.Comments) > 0 Then
        If DialogQuestion("Are you sure you wish to delete all the comments for the question " & oElement.Name & "?") = vbYes Then
            oResponse.Comments = ""
            'call to DisplayCurrentResponse to update comment/note image
            Call moFormBuilder.DisplayCurrentResponse(oItem, True)
            ' Let the SetFocus (below) redraw the Comments appropriately
        End If
    Else
        DialogInformation "There are no comments for this question"
    End If
    
    ' Set focus back to where we came from
    Call moFormBuilder.ResetFormFocus(oItem)
    
    Set oElement = Nothing
    Set oResponse = Nothing
    
Exit Sub

ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "RemoveComment", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub ChangeUnits(oItem As EFormItem)
'---------------------------------------------------------------------
' Show the Units conversion dialog
'---------------------------------------------------------------------
Dim oElement As eFormElementRO
Dim dblResult As Double

    If oItem Is Nothing Then Exit Sub
    
    Set oElement = oItem.Element

    If frmUnitConversion.Display(oElement, dblResult) Then
        ' Set the text of the relevant text box
        txtCRFElement(oItem.ControlIndex).Text = CStr(dblResult)
        Call txtCRFElement_LostFocus(oItem.ControlIndex)
    Else
        ' Set focus back to where we came from
        Call moFormBuilder.ResetFormFocus(oItem)
    End If

    Set oElement = Nothing
    
End Sub


'---------------------------------------------------------------------
Private Sub ToggleMissingUnobtainable(oItem As EFormItem)
'---------------------------------------------------------------------
' If Missing, change to Unobtainable
' If Unobtainable, change to Missing
' Assume we only get here if it's Requested, Missing or Unobtainable
' NCJ 12 Jul 02 - Do not ask for RFC for "new" responses
' NCJ 12 Aug 02 - Let Response object decide on RFC
'---------------------------------------------------------------------
Dim oElement As eFormElementRO
Dim oResponse As Response
Dim nNewStatus As Integer
Dim bCanChangeStatus As Boolean
Dim sRFC As String
Dim sRFCCaption As String

    If oItem Is Nothing Then Exit Sub
    
    
    ' Pick up current Element
    Set oElement = oItem.Element
    Set oResponse = oItem.Response ' moEFI.Responses.ResponseByElement(oElement, oItem.RowNo)
    
    Select Case oResponse.Status
    Case eStatus.Missing, eStatus.Requested
        nNewStatus = eStatus.Unobtainable
    Case eStatus.Unobtainable
        nNewStatus = eStatus.Missing
    End Select
    
    bCanChangeStatus = True

    ' NCJ 13 Aug 02 - Use new RequiresStatusRFC
    If oResponse.RequiresStatusRFC(nNewStatus) Then
        ' Ask for RFC and set bCanChangeStatus to FALSE if they don't give one
        sRFCCaption = "Please enter the reason for changing the status of "
        frmRFC.lblRFC.Caption = sRFCCaption & oElement.Name
        frmRFC.DisplayRFC mcolRFCs, GetCoordForForm(oItem)
        sRFC = frmRFC.GetRFC
        If Len(sRFC) < 1 Then
            bCanChangeStatus = False
        End If
    End If
    

    If bCanChangeStatus Then
        ' Change the status
        ' NCJ 12 Aug 02 - Use new SetStatus call
        Call oResponse.SetStatus(nNewStatus, sRFC, "")
'        oResponse.Status = nNewStatus
        ' Update derivations & skips etc.
        Call moFormBuilder.UpdateResponses(ChangingResponse, oResponse)
        ' User has now done something
        mbUserDidSomething = True
    End If
    
    ' Display the current response
    ' (in case user typed anything before clicking this button)
    Call moFormBuilder.DisplayCurrentResponse(oItem, True)
    ' NCJ 8 Jan 02 - Also update Group status if necessary
    If Not oElement.OwnerQGroup Is Nothing Then
        'MLM 10/09/02: Find the OwnerQGroup's Item:
        Call moFormBuilder.DisplayCurrentResponse(moFormBuilder.EFormItemByIndex(moFormBuilder.IndexForElementRow(oElement.OwnerQGroup.Element, 1)))
        'Call moFormBuilder.DisplayCurrentResponse(oElement.OwnerQGroup.Element, 1)
    End If
    
    ' Set focus back to appropriate question
    DoEvents
    If bCanChangeStatus Then
        'REM 24/06/02 - Bug 2.2.16 Roche No. 19 Changed so we move to next field
        Call moFormBuilder.FocusToNextField(oElement, oItem.RowNo)
    Else
        ' Status not changed so stay where we are
        Call moFormBuilder.ResetFormFocus(oItem)
    End If
    
    Set oElement = Nothing
    Set oResponse = Nothing
    
End Sub

'---------------------------------------------------------------------
Public Sub KeyDown(KeyCode As Integer, Optional Shift As Integer = 0)
'---------------------------------------------------------------------
' Key down in the data entry form - process function keys
' Ignore keypresses if we're in the middle of changing forms
' NCJ 23 Apr 03 - Added F1 for MACRO 3.0 Help
' NCJ 28 Mar 07 - Bug 2793 - Added F10 to list of keys to validate
'---------------------------------------------------------------------

    'SDM 22/12/99 SR1774    Ensure that ctrl etc are not pressed
    If Shift > 1 Then Exit Sub
    
    ' NCJ Check we're not changing forms
    If mbDontInterrupt Then Exit Sub
    
    Select Case KeyCode
    ' NCJ 28 Mar 07 - Bugs 2793, 2896 - Added F10, F11, F12
    Case vbKeyF1, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyF6, vbKeyF7, _
        vbKeyF10, vbKeyF11, vbKeyF12
        'TA 06/02/2002: validate the current field and
                    'disallow these actions if there is invalid data
        If ThereIsInvalidData Then
            Exit Sub
        End If
    End Select
    
    'ash 5/11/2001 Added F3-F12. F7 already added. Bug Report
    Select Case KeyCode
    Case vbKeyF1
        ' NCJ 23 Apr 03 - Display Help
        Call MACROHelp(Me.hWnd, App.Title)
        Exit Sub
    Case vbKeyF3
        'Previous Form
        Call GotoEFI(eaPrevious)
        Exit Sub
    Case vbKeyF4
        'Next Form
        Call GotoEFI(eaNext)
        Exit Sub
    Case vbKeyF5
        'Print Form
        Call PrintEForm(EFI)
        Exit Sub
    Case vbKeyF6
        'Return
        Call EFIReturn
        Exit Sub
    Case vbKeyF7
        ' Save
        Call SaveResponses(eEFormAction.eaSave)
        ' NCJ 10 Jun 02 - CBB 2.2.13/63 Refresh display for all responses
        Call moFormBuilder.DisplayCurrentResponses
        Exit Sub
    Case vbKeyF8
        'Cancel
        Call EFICancel
        Exit Sub
    Case vbKeyF9
        'Clear
        Call ClearQuestion(moFormBuilder.CurrentItem)
        Exit Sub
    Case vbKeyF10
        ' Display rt. mouse menu
        Call ElementRightMouseMenu(moFormBuilder.CurrentItem, False)
        Exit Sub
    Case vbKeyF11
        ' Add comment
        Call ShowAddComment(moFormBuilder.CurrentItem)
        Exit Sub
    Case vbKeyF12
        ' Delete Comment
         Call RemoveComment(moFormBuilder.CurrentItem)
        Exit Sub
    Case vbKeyDown, vbKeyUp, vbKeyPageDown, vbKeyPageUp
        ' RS 05/02/03: Add Key scrolling to listview
        ' Scroll Listview, if visible
        If lvwPopupList.Visible And lvwPopupList.Top > 0 Then
            KeyScrollDropDown KeyCode
        End If
    Case Else
    
    End Select
End Sub

'---------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    Call KeyDown(KeyCode, Shift)
    
End Sub

'---------------------------------------------------------------------
Private Sub KeyScrollDropDown(KeyCode As Integer)
'---------------------------------------------------------------------
' Scroll the currently opened category question using keyboard
' NCJ 18 Jan 05 - Deal with list refresh (Issue 2502)
' NCJ 18 Dec 06 - Bug 2850 - Deal with empty drop-down list! (Can happen if ListFiltering is on)
'---------------------------------------------------------------------
Dim oItem As EFormItem
Dim nDelta As Integer
Dim nNewIndex As Integer

    ' NCJ 18 Dec 06 - Do nothing if no list items
    If lvwPopupList.ListItems.Count = 0 Then Exit Sub
    
    ' NCJ 18 Dec 06 - Moved this bit of code here from further down
    If lvwPopupList.SelectedItem Is Nothing Then
        ' Just select the first item
        Set lvwPopupList.SelectedItem = lvwPopupList.ListItems(0)
        lvwPopupList.SelectedItem.EnsureVisible
    Else
    
        Select Case KeyCode
            Case vbKeyDown: nDelta = 1
            Case vbKeyUp: nDelta = -1
            Case vbKeyPageDown: nDelta = 5
            Case vbKeyPageUp: nDelta = -5
        End Select
    
        nNewIndex = lvwPopupList.SelectedItem.Index + nDelta
        ' Check we're not off the top or the bottom
        If nNewIndex < 1 Then
            nNewIndex = 1
        ElseIf nNewIndex > lvwPopupList.ListItems.Count Then
            nNewIndex = lvwPopupList.ListItems.Count
        End If
    
        If lvwPopupList.DropHighlight Is Nothing Then
            nNewIndex = 1
        End If
            
    ' NCJ 18 Dec 06 - Moved this code further up
'    If lvwPopupList.SelectedItem Is Nothing Then
'        Set lvwPopupList.SelectedItem = lvwPopupList.ListItems(0)
'        lvwPopupList.SelectedItem.EnsureVisible
'    Else
        Set lvwPopupList.SelectedItem = lvwPopupList.ListItems(nNewIndex)
        Set lvwPopupList.DropHighlight = lvwPopupList.SelectedItem
        lvwPopupList.SelectedItem.EnsureVisible
        
        ' Simulate Mouseclick, except do not execute lostfocus event
        
        Set oItem = moFormBuilder.CurrentItem
        mbMouseDownOnPopupList = False
        ' NCJ 18 Jan 05 - We don't want a list refresh for Arrow key in drop-down
        mbArrowInDropDown = True
        txtCRFElement(oItem.ControlIndex).Text = lvwPopupList.SelectedItem.SubItems(1)
        txtCRFElement(oItem.ControlIndex).SelStart = 0 ' Make sure cursor is at start
        'Call txtCRFElement_LostFocus(oItem.ControlIndex)
        
    End If

End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
'---------------------------------------------------------------------
Dim lStyle As Long
Const lFormColour As Long = vbButtonFace

    mbDuringLoad = True
    mbDontInterrupt = True
    
    mbMouseDownOnPopupList = False
    mnMouseDownIndex = 0
    mnValidatingResponse = 0
    mbMouseDownInNextField = False
    mbMouseDownOnCalendar = False
    mbArrowInDropDown = False
    
    mbDirtyField = False
    mbCRFPageClicked = False
    
   ' Turn on key preview for form, so that F1 (Help) can be trapped by form
    Me.KeyPreview = True
    Me.BackColor = lFormColour
    cmdNextButton.MousePointer = vbCustom
    cmdNextButton.MouseIcon = frmImages.CursorHandPoint.MouseIcon
    picBackButton.MousePointer = vbCustom
    picBackButton.MouseIcon = frmImages.CursorHandPoint.MouseIcon
  
    ' NCJ 17 Jan 03
    mlPreviousEformTaskId = 0
    msPreviousEformName = ""
    
    ' NCJ 19 Jan 05 - Shall we enable keyboard list filtering? (Default to FALSE)
    mbEnableListFiltering = (LCase(GetMACROSetting(MACRO_SETTING_LIST_FILTER, "false")) = "true")

    mbDontInterrupt = False
    mbDuringLoad = False
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdAttachment_GotFocus(Index As Integer)
'---------------------------------------------------------------------
' Focus received by Attachment button
'---------------------------------------------------------------------
Dim oItem As EFormItem

    Set oItem = moFormBuilder.EFormItemByIndex(Index)

    ' Fill in the Question Information Panel
    Call moFormBuilder.ElementGotFocus(oItem)
    Set moLastSelectedControl = cmdAttachment(Index)
    
    Set oItem = Nothing

    ' NCJ 15 Nov 04 - Issue 2463 - Must reset mbDirtyField field here! (So we don't try to process non-existent text fields)
    mbDirtyField = False
    
    'NCJ 1 Feb 02 - scroll if off screen
    Call ScrollIfOffScreen(cmdAttachment(Index))

End Sub

'---------------------------------------------------------------------
Private Sub Form_Resize()
'---------------------------------------------------------------------
' Reset the page size after drawing all the eForm Elements
' MLM 07/12/01: Modified calculation of v/hsbCRFPage.Max so that can't scroll off picCRFPage.
' NCJ 21 May 03 - Make sure page is all visible when we remove scroll bars
'---------------------------------------------------------------------
Dim bHasVSB As Boolean
Dim bHasHSB As Boolean
    
    'TA 18/10/2001: changed to on erro resume next so after failing to move a control
    '       it will still try the next one (in particular the scroll bars)
    On Error Resume Next
    
    'SDM 27/01/00 SR2817
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With Me
    
        ' Set position and size of container controls
        .fraCRFPage.Top = 0 '.picToolbar.Top + .picToolbar.Height ' + 50
        .fraCRFPage.Left = 0
        .fraCRFPage.Height = .Height - .fraCRFPage.Top      'TA remove becasue no title bar- 450 'Fudge factor!
        .fraCRFPage.Width = .ScaleWidth - .fraCRFPage.Left
        
'        ' NCJ 21 Feb 03 - Must also reset the picCRFPage
'        ' in case it was scrolled to an unsuitable position on the previous eForm
'        .picCRFPage.Top = 0
'        .picCRFPage.Left = 0
        
        bHasVSB = (.picCRFPage.Height > .fraCRFPage.Height) Or _
            (.picCRFPage.Height > .fraCRFPage.Height - .hsbCRFPage.Height And .picCRFPage.Width > .fraCRFPage.Width)
        bHasHSB = (.picCRFPage.Width > .fraCRFPage.Width) Or _
            (.picCRFPage.Width > .fraCRFPage.Width - .vsbCRFPage.Width And .picCRFPage.Height > .fraCRFPage.Height)
        
        If bHasVSB Then
            Set .vsbCRFPage.Container = .fraCRFPage
            .vsbCRFPage.Top = 0
            If bHasHSB Then
                'note that the max value is divided by 10 to accommodate long forms
                .vsbCRFPage.Max = (.picCRFPage.Height - .fraCRFPage.Height + .hsbCRFPage.Height) / 10
                .vsbCRFPage.LargeChange = (.fraCRFPage.Height - .hsbCRFPage.Height) / .picCRFPage.Height * .vsbCRFPage.Max
                .vsbCRFPage.Height = .fraCRFPage.Height - .hsbCRFPage.Height
            Else
                .vsbCRFPage.Max = (.picCRFPage.Height - .fraCRFPage.Height) / 10
                .vsbCRFPage.LargeChange = .fraCRFPage.Height / .picCRFPage.Height * .vsbCRFPage.Max
                .vsbCRFPage.Height = .fraCRFPage.Height
            End If
            .vsbCRFPage.Left = .fraCRFPage.Width - .vsbCRFPage.Width
            '.vsbCRFPage.Left = .picCRFPage.Width + .vsbCRFPage.Width
            .vsbCRFPage.SmallChange = .vsbCRFPage.Max / 10
 '           .vsbCRFPage.Value = 0
            .vsbCRFPage.Visible = True
        Else
            .vsbCRFPage.Visible = False
            ' NCJ 21 May 03 - Must also reset picCRFPage to the top
            .picCRFPage.Top = 0
        End If
        
        If bHasHSB Then
            Set .hsbCRFPage.Container = .fraCRFPage
           .hsbCRFPage.Left = 0
            If bHasVSB Then
                .hsbCRFPage.Max = .picCRFPage.Width - .fraCRFPage.Width + .vsbCRFPage.Width
                .hsbCRFPage.LargeChange = (.fraCRFPage.Width - .vsbCRFPage.Width) / .picCRFPage.Width * .hsbCRFPage.Max
                .hsbCRFPage.Width = .fraCRFPage.Width - .vsbCRFPage.Width
            Else
                .hsbCRFPage.Max = .picCRFPage.Width - .fraCRFPage.Width
                .hsbCRFPage.LargeChange = .fraCRFPage.Width / .picCRFPage.Width * .hsbCRFPage.Max
                .hsbCRFPage.Width = .fraCRFPage.Width
            End If
            .hsbCRFPage.Top = .fraCRFPage.Height - .hsbCRFPage.Height
            .hsbCRFPage.SmallChange = .hsbCRFPage.Max / 10
  '          .hsbCRFPage.Value = 0
            .hsbCRFPage.Visible = True
        Else
            .hsbCRFPage.Visible = False
            ' NCJ 21 May 03 - Must also reset picCRFPage to the left
            .picCRFPage.Left = 0
        End If
        
        If bHasVSB And bHasHSB Then
            'MLM 10/12/01: If both scroll bars are displayed, blank out the bottom-right corner
            'of the form (between the ends of the scroll bars)
            Set .picBlank.Container = .fraCRFPage
            .picBlank.Top = .fraCRFPage.Height - .hsbCRFPage.Height
            .picBlank.Left = .fraCRFPage.Width - .vsbCRFPage.Width
            .picBlank.Height = .hsbCRFPage.Height
            .picBlank.Width = .vsbCRFPage.Width
            .picBlank.Visible = True
        Else
            .picBlank.Visible = False
        End If
            
            
         'resize code added TA 09/12/2002
        .picHeaderBar.Width = picCRFPage.Width
        .picLab.Width = picHeaderBar.Width
        
        .lblVisit.Width = picHeaderBar.Width / 2 - txtVisitDate.Width
        .txtVisitDate.Left = lblVisit.Left + lblVisit.Width
        
        .lblEForm.Left = txtVisitDate.Left + txtVisitDate.Width
        .lblEForm.Width = lblVisit.Width
        .cmdChangeLab.Left = lblEForm.Left
        .lblLab.Left = lblVisit.Left
        
        
        .picHeaderExt.Width = .fraCRFPage.Width
        .picHeaderExt.Left = 0
        .picHeaderExt.Top = .picCRFPage.Top
                
    End With

End Sub

'---------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'---------------------------------------------------------------------
' Clear down our form variables
'---------------------------------------------------------------------
'notify external listeners that this form has unloaded (and what efi was there when it did)

    RaiseEvent Unload
    
    If Not moFormBuilder Is Nothing Then
        Call moFormBuilder.Terminate
        Set moFormBuilder = Nothing
    End If
    Set moEForm = Nothing
    Set moUser = Nothing
    Set moSubject = Nothing
    Set mcolRFCs = Nothing
    Set moLastSelectedControl = Nothing

End Sub

'---------------------------------------------------------------------
Private Sub fraCRFPage_Click()
'---------------------------------------------------------------------
'Focus away from drop down so hide it
'---------------------------------------------------------------------

    Call HidePopupAndCalendar

End Sub

'---------------------------------------------------------------------
Private Sub hsbCRFPage_Change()
'---------------------------------------------------------------------
' MLM 23/01/02: Rolled forward from 2.2: do horizontal scrolling.
'---------------------------------------------------------------------
    
    picCRFPage.Left = CSng(-hsbCRFPage.Value)

End Sub

'---------------------------------------------------------------------
Private Sub hsbQuGroup_Change(Index As Integer)
'---------------------------------------------------------------------
' Change to RQG's horizontal scroll bar
'---------------------------------------------------------------------

    ' NCJ 4 Nov 02 - Hide popup and calendar before scrolling
    Call HidePopupAndCalendar
    Call moFormBuilder.QGroupHScroll(Index, hsbQuGroup(Index).Value)
    
End Sub

'---------------------------------------------------------------------
Private Sub hsbQuGroup_Scroll(Index As Integer)
'---------------------------------------------------------------------
' Change to RQG's horizontal scroll bar
'---------------------------------------------------------------------

    Call HidePopupAndCalendar
    Call moFormBuilder.QGroupHScroll(Index, hsbQuGroup(Index).Value)
    
End Sub

Private Sub imgChangeCount_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call imgStatus_MouseUp(Index, Button, Shift, X, Y)

End Sub

Private Sub imgSDVStatus_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call imgStatus_MouseUp(Index, Button, Shift, X, Y)

End Sub

'---------------------------------------------------------------------
Private Sub lblCaption_Click(Index As Integer)
'---------------------------------------------------------------------
' NCJ 8 Nov 02
' If a Hotlink, move to specified eForm
' NCJ 26 Jan 06 - Bug 2636 - Don't evaluate Hotlink here
'---------------------------------------------------------------------
Dim oItem As EFormItem
Dim leFormTaskId As Long

    ' Switch off "hand" cursor
    Screen.MousePointer = vbDefault
    ' Pick up the relevant element (may be nothing if not a Hotlink)
    Set oItem = moFormBuilder.EFormItemByIndex(Index)
    If Not oItem Is Nothing Then
        If oItem.Element.ControlType = gn_HOTLINK Then
'            leFormTaskId = Me.EFI.GetHotlinkTarget(oItem.Element.Hotlink)
'            If leFormTaskId = 0 Then
'                DialogInformation "The eForm associated with this link is not available"
'            ElseIf leFormTaskId = Me.EFI.EFormTaskId Then
'                DialogInformation "This link specifies the current eForm"
'            Else
'                ' Go to the specified eForm
'                Call eFormAction(eEFormAction.eaGotoHotlink, leFormTaskId)
'            End If
            ' NCJ 26 Jan 06 - Pass through unevaluated Hotlink
            Call eFormAction(eEFormAction.eaGotoHotlink, oItem.Element.Hotlink)
        End If
        Set oItem = Nothing
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub lblCaption_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' Display popup menu on caption right click
' (but only if there's an associated question - NCJ 8 Nov 02)
'---------------------------------------------------------------------
Dim oItem As EFormItem
    
    Set oItem = moFormBuilder.EFormItemByIndex(Index)
    If Not oItem Is Nothing Then
        Call imgStatus_MouseUp(Index, Button, Shift, X, Y)
        Set oItem = Nothing
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub lvwPopupList_ItemClick(ByVal Item As MSComctlLib.ListItem)
'---------------------------------------------------------------------
' Click on a list view item
'---------------------------------------------------------------------
Dim nElementId As Integer
Dim oItem As EFormItem

    If moFormBuilder.PopupListVisible And mbMouseDownOnPopupList = True Then
        Set oItem = moFormBuilder.CurrentItem
        mbMouseDownOnPopupList = False
        ' NCJ 18 Jan 05 - Hide list at this point (to prevent unwanted rebuilding)
        Call moFormBuilder.HidePopupList
        txtCRFElement(oItem.ControlIndex).Text = Item.SubItems(1)
        Call txtCRFElement_LostFocus(oItem.ControlIndex)
        Set oItem = Nothing
    End If

End Sub

'---------------------------------------------------------------------
Private Sub lvwPopupList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------

    mbMouseDownOnPopupList = True
    
End Sub

'---------------------------------------------------------------------
Private Sub lvwPopupList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
'SDM SR1444 26/10/99   Highlight the item that the mouse is over.
'---------------------------------------------------------------------
    
    If moFormBuilder.PopupListVisible Then
        Set lvwPopupList.DropHighlight = lvwPopupList.HitTest(X, Y)
    End If

End Sub

'---------------------------------------------------------------------
Private Sub lvwPopupList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------

    mbMouseDownOnPopupList = False

End Sub
'---------------------------------------------------------------------
Private Sub moSubject_SubjectLabelChanged(sOldLabel As String, sNewLabel As String)
'---------------------------------------------------------------------
'pass on info that subejct label has changed
'---------------------------------------------------------------------
    
    mbSubjectLabelChanged = True
    
End Sub

'---------------------------------------------------------------------
Private Sub optCRFGroup_Click(Index As Integer, OptionIndex As Integer)
'---------------------------------------------------------------------
' Click on an option button in an option button group
' (This event is also fired if they hit the Space Bar to select)
'---------------------------------------------------------------------
Dim sCatCode As String
Dim oItem As EFormItem
Dim bStayOnField As Boolean

    'TA 09/04/2002 restart timeout timer at appropriate intervals
    Call RestartSystemIdleTimer

    ' Pick up the relevant element
    Set oItem = moFormBuilder.EFormItemByIndex(Index)

    ' NCJ 7 Jun 02 (CBB2.2.12/41) Assume we want to stay put
    bStayOnField = True
    ' Also clear the Dirty field flag
    mbDirtyField = False
    
    If Me.Visible And optCRFGroup(Index).Visible Then
        ' Pick up the selected category code
        sCatCode = optCRFGroup(Index).TagValue(OptionIndex)

        If ValidateProposedValue(sCatCode, oItem) Then
            ' NCJ 7 Jun 02 - Stay where we are if Warning (CBB 2.2.12/41)
            bStayOnField = (oItem.Response.Status = eStatus.Warning)
        End If
        Call FocusToAppropriateField(oItem, bStayOnField)
    End If

    Set oItem = Nothing
    
End Sub

'---------------------------------------------------------------------
Private Sub optCRFGroup_LostFocus(Index As Integer)
'---------------------------------------------------------------------
'MLM 27/03/03: Added.
'---------------------------------------------------------------------

    mbMouseDownInNextField = False

End Sub

'---------------------------------------------------------------------
Private Sub optCRFGroup_OnFocus(Index As Integer, OptionIndex As Integer)
'---------------------------------------------------------------------
' Option button in an option group gets the focus
' Check to see if we're in the same option group
' (Assume we're not in a Question Group)
'---------------------------------------------------------------------
Dim oItem As EFormItem
Dim oElement As eFormElementRO

    ' Pick up the relevant element
    Set oItem = moFormBuilder.EFormItemByIndex(Index)
    Set oElement = oItem.Element
    
    ' NCJ 10/6/02 Also clear the Dirty field flag
    ' NCJ 5 Mar 03 - Do it here rather than inside the If...Then
    mbDirtyField = False
    
    If moFormBuilder.CurrentElementID <> oElement.ElementID Then
        ' Fill in the Question Information Panel
        Call moFormBuilder.ElementGotFocus(oItem)
        Set moLastSelectedControl = optCRFGroup(Index)
    End If
    
    Set oItem = Nothing
    Set oElement = Nothing

    'NCJ 1 Feb 02 - scroll if off screen
    Call ScrollIfOffScreen(optCRFGroup(Index))

End Sub

'---------------------------------------------------------------------
Private Sub optGraphical_KeyPress(Index As Integer, KeyAscii As Integer)
'---------------------------------------------------------------------
'REM 19/06/02
'Push the space bar on highlighted option box to move to next field (CBB 2.2.16 No. 3)
'---------------------------------------------------------------------

    If KeyAscii = vbKeySpace Then
        mnMouseDownIndex = Index
        Call optGraphical_Click(Index)
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub optGraphical_Click(Index As Integer)
'---------------------------------------------------------------------
' Click on a push button
' NCJ 20 Sept 02 - Do not process the click if we're already validating something else!
'---------------------------------------------------------------------
Dim nElementId As Integer
Dim sCatCode As String
Dim oItem As EFormItem

    'TA 09/04/2002 restart timeout timer at appropriate intervals
    Call RestartSystemIdleTimer

    ' NCJ 20 Sept 02 - Ignore if already validating
    If mnValidatingResponse > 0 Then Exit Sub
    
    ' Pick up the relevant element
    Set oItem = moFormBuilder.EFormItemByIndex(optGraphical(Index).Container.Index)

    If Me.Visible And optGraphical(Index).Visible _
     And (mnMouseDownIndex = Index) Then
        
        'reset flags
        mnMouseDownIndex = 0

        ' Pick up the selected category code
        sCatCode = optGraphical(Index).Tag
        
        If ValidateProposedValue(sCatCode, oItem) Then
            ' Move on to the next eForm element
            ' NCJ 7 Jun 02 - Stay where we are if Warning (CBB 2.2.12/41)
            Call FocusToAppropriateField(oItem, (oItem.Response.Status = eStatus.Warning))
        End If

     Else
        ' NCJ 7/4/00 SR 3159 - Could be a rogue VB event, so ensure correct value is displayed
        ' MLM 09/09/02: Pass oItem
        moFormBuilder.DisplayCurrentResponse oItem
        ' Reset flags
        mnMouseDownIndex = 0

    End If
    
    Set oItem = Nothing
    
End Sub

'---------------------------------------------------------------------
Private Sub optGraphical_GotFocus(Index As Integer)
'---------------------------------------------------------------------
' Focus received by option box
'---------------------------------------------------------------------
Dim oElement As eFormElementRO
Dim oItem As EFormItem

    optGraphical(Index).BackColor = glYELLOW
    
'    'set the colour of the notestatus image accordingly
'    picNoteStatus(optGraphical(Index).Container.Index).BackColor = optGraphical(Index).Container.BackColor
    
    ' NCJ 10/6/02 Also clear the Dirty field flag
    mbDirtyField = False
        
    ' Pick up the relevant element
    Set oItem = moFormBuilder.EFormItemByIndex(optGraphical(Index).Container.Index)
    Set oElement = oItem.Element

    ' Fill in the Question Information Panel
    Call moFormBuilder.ElementGotFocus(oItem)
    Set moLastSelectedControl = optGraphical(Index)

    Set oElement = Nothing
    Set oItem = Nothing

    'NCJ 1 Feb 02 - scroll if off screen
    Call ScrollIfOffScreen(optGraphical(Index))

End Sub

'---------------------------------------------------------------------
Private Sub optGraphical_lostfocus(Index As Integer)
'---------------------------------------------------------------------
' Push button lost focus
'---------------------------------------------------------------------
    
    optGraphical(Index).BackColor = Me.picCRFPage.BackColor
    'MLM 27/03/03:
    mbMouseDownInNextField = False

End Sub

'---------------------------------------------------------------------
Private Sub optGraphical_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' Set flag to say user actually clicked here (if it's a change)
'---------------------------------------------------------------------
Dim oItem As EFormItem

    optGraphical(Index).BackColor = glYELLOW
    
'    'set the colour of the notestatus image accordingly
'    picNoteStatus(optGraphical(Index).Container.Index).BackColor = optGraphical(Index).BackColor
    
    ' Pick up the relevant element
    Set oItem = moFormBuilder.EFormItemByIndex(optGraphical(Index).Container.Index)
    
    ' Has it changed?
    If oItem.Response.ValueCode <> optGraphical(Index).Tag Then
        mnMouseDownIndex = Index
    End If
    
    Set oItem = Nothing
    
End Sub

'---------------------------------------------------------------------
Private Sub picBackButton_Click()
'---------------------------------------------------------------------
' Go back to where we came from
'---------------------------------------------------------------------

    ' Check for unvalidated data first
    If ThereIsInvalidData Then
        Exit Sub
    End If
    
    ' Go to previous form, if any
    If mlPreviousEformTaskId > 0 Then
        Call GotoEFI(eaGoBack, mlPreviousEformTaskId)
    Else
        Call EFIReturn
    End If

End Sub

'---------------------------------------------------------------------
Private Sub picBackButton_GotFocus()
'---------------------------------------------------------------------
' "Go back" button has the focus
' so clear question focus
'---------------------------------------------------------------------

    Call moFormBuilder.NoQuestionFocus
    picBackButton.BorderStyle = 1
    ' NCJ 19 Feb 03 - Scroll to the button if not visible
    Call ScrollIfOffScreen(picBackButton)

End Sub

'---------------------------------------------------------------------
Private Sub picBackButton_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------
' Intercept Return or Space to simulate button click
'---------------------------------------------------------------------

    Select Case KeyCode
    Case vbKeyReturn, vbKeySpace
        picBackButton_Click
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub picBackButton_LostFocus()
'---------------------------------------------------------------------

    picBackButton.BorderStyle = 0

End Sub

'MLM 21/05/02: Don't hide the drop-down list, as it isn't displayed any more.
''---------------------------------------------------------------------
'Private Sub picCRFPage_Click()
''---------------------------------------------------------------------
''Focus way from drop down so hide it
''---------------------------------------------------------------------
'
'    moFormBuilder.HidePopupList
'
'End Sub

'---------------------------------------------------------------------
Private Sub picCRFPage_GotFocus()
'---------------------------------------------------------------------
' ZA 15/08/01, check if picCRFPage has got the focus. If it has
' try to figure out what was the last control that has the focus,
' and then set the focus back to that control
'---------------------------------------------------------------------

    Call FocusToLastSelectedControl

End Sub

'---------------------------------------------------------------------
Private Sub FocusToLastSelectedControl()
'---------------------------------------------------------------------
' Set focus back to last selected control
'---------------------------------------------------------------------
Dim nIndex As Integer

    If Not moLastSelectedControl Is Nothing Then
    
        nIndex = moLastSelectedControl.Index
        ' ZA 29/10/01 - check if the control is enabled.
        ' If an error dialog has appeared on top we can't
        ' set the focus back to the last selected control
        
        If moLastSelectedControl.Enabled = True Then
        
            Select Case moLastSelectedControl.Name
                Case "txtCRFElement"
                    'MLM 21/05/02: Set flag so that txtCRFElement_GotFocus() knows not to redraw drop-down list
                    mbCRFPageClicked = True
                    txtCRFElement(nIndex).SetFocus
                Case "optCRFGroup"
                    optCRFGroup(nIndex).SetFocus
                Case "cmdAttachment"
                    cmdAttachment(nIndex).SetFocus
                Case "optGraphical"
                    optGraphical(nIndex).SetFocus
                Case Else
                    'do nothing
            End Select
        End If
       
    End If

End Sub

'---------------------------------------------------------------------
Public Function ClosedSuccessfully(Optional bPromptUser As Boolean = True, _
                                    Optional ByRef lEFITaskId As Long, Optional ByRef lResponseTaskId As Long, Optional ByRef nRptNo As Integer) As Boolean
'---------------------------------------------------------------------
' Public method to try to unload form.
' This must only be called through IsDatEntryFormLoaded.
' (to check that this form is loaded).
' Asks user if they want to unload if the data has changed.
' This is called when closing a subject
'
' MLM 10/09/02: Use new SaveNeeded function, which takes the Visit EForm into consideration.
'TA 07/01/2003: Retuyn the efi and response identifiers
' ic 14/09/2005 added clinical coding
'   NCJ 11 Jul 06 - Bug 2761 - Check for null responses
'---------------------------------------------------------------------
Dim nQuestionResponse As Integer
Dim bClose As Boolean
'MLM 10/09/02: set by SaveNeeded
Dim bResponsesExist As Boolean
Dim bCancelResponses As Boolean
Dim bAsked As Boolean
Dim bSave As Boolean

    bAsked = False


    If ThereIsInvalidData Then
        ClosedSuccessfully = False
        Exit Function
    End If
    
    
    'TA 07/01/2003: Return the efi and response identifiers
    lEFITaskId = mlEFormTaskId
    ' NCJ 11 Jul 06 - Set up default values first
    '-1 are the default values for frmeFormDataEntry.display
    lResponseTaskId = -1
    nRptNo = -1
    If Not moFormBuilder.CurrentItem Is Nothing Then
        With moFormBuilder.CurrentItem
            ' NCJ 11 Jul 06 - Check for null response!
            If Not .Response Is Nothing Then
                lResponseTaskId = .Response.ResponseId
                nRptNo = .Response.RepeatNumber
            End If
        End With
'    Else
'    '-1 are the default values for frmeFormDataEntry.display
'        lResponseTaskId = -1
'        nRptNo = -1
    End If
    
    
    If gbClinicalCoding Then
        If (moUser.CheckPermission(gsFnCodeClinicalResponse) Or moUser.CheckPermission(gsFnChangeClinicalStatus) _
            Or moUser.CheckPermission(gsFnValidateClinicalCode)) Then
            'check for changes to clinical codes
            If SaveCodedTermHistoriesNeeded() Then
                If bPromptUser Then
                    bAsked = True
                    nQuestionResponse = DialogQuestion("Some data has changed - would you like to save it?")
                    If (nQuestionResponse = vbYes) Then
                        bSave = True
                    Else
                        bSave = False
                    End If
                Else
                    bSave = True
                End If
                If bSave Then Call moSubject.SaveCodedResponses(EFI)
            End If
        End If
    End If

    
    If SaveNeeded(bResponsesExist) Then
        If bPromptUser And bResponsesExist And mbUserDidSomething Then
            If gbClinicalCoding Then
                If Not bAsked Then
                    nQuestionResponse = DialogQuestion("Some data has changed - would you like to save it?", , True)
                End If
            Else
                nQuestionResponse = DialogQuestion("Some data has changed - would you like to save it?", , True)
            End If
        Else
            ' Automatic save for an empty form or no user prompt
            nQuestionResponse = vbYes
        End If
        Select Case nQuestionResponse
        Case vbYes
            ' NCJ 14 Aug 02 SaveResponses may not succeed (because of revalidation)
'            bClose = SaveResponses
            ' NCJ 24 Jan 03 - Don't give warnings etc. if user hasn't done anything
            bClose = SaveResponses(eEFormAction.eaReturn, mbUserDidSomething, bCancelResponses)
            ' Revalidation or saving failed
            ' If the user didn't do anything then don't worry about it,
            If (Not bClose) And (Not mbUserDidSomething) Then
                bClose = True
            End If
            
            If bClose Then
                'clear old responses and elements from memory, and remove eForm lock
                ' NCJ 10 Apr 03 - Cancel if SaveResponses said they should be cancelled
                If bCancelResponses Then
                    Call moSubject.CancelResponses(EFI)
                Else
                    Call moSubject.RemoveResponses(EFI, True)
                End If
            End If
            
        Case vbNo
            ' Don't save any data
            'clear old responses and elements from memory
            Call moSubject.CancelResponses(EFI)
            bClose = True
        Case vbCancel
            bClose = False
        End Select
    Else
        'no changes - clear old responses and elements from memory, and remove eForm lock
        Call moSubject.RemoveResponses(EFI, True)
        bClose = True
    End If
    

    If bClose = True Then
        Unload Me
    End If
    
    ClosedSuccessfully = bClose
    
End Function

'---------------------------------------------------------------------
Private Sub EFICancel()
'---------------------------------------------------------------------
' Cancel pressed.
' If it's an empty form we don't bother asking if they want to save
'
' MLM 10/09/02: Use new SaveNeeded function, which takes Visit EForm into consideration
' ic 14/09/2005 added clinical coding
'---------------------------------------------------------------------

Dim bResponsesExist As Boolean
Dim bAsked As Boolean
Dim bClose As Boolean
    
    bAsked = False

    HourglassOn
    
    If gbClinicalCoding Then
        'check for changes to clinical codes
        If SaveCodedTermHistoriesNeeded() Then
            If DialogQuestion("Some data has changed - are you sure you wish to close without saving?") = vbNo Then
                HourglassOff
                Exit Sub
            Else
                bAsked = True
            End If
        End If
    End If
    
    
    'If moEFI.Responses.Count > 0 And moEFI.Responses.Changed Then
    If SaveNeeded(bResponsesExist) Then
        ' Only ask for confirmation if there are responses AND the user actually did something
        If bResponsesExist And mbUserDidSomething Then
            If gbClinicalCoding Then
                'dont ask the user 'are they sure' twice
                If Not bAsked Then
                    bClose = (DialogQuestion("Some data has changed - are you sure you wish to close without saving?") = vbYes)
                Else
                    bClose = True
                End If
                
                If bClose Then
                    'clear old responses and elements from memory, and remove eForm lock
                    Call moSubject.CancelResponses(EFI)
                    Unload Me
                End If
            Else
                If DialogQuestion("Some data has changed - are you sure you wish to close without saving?") = vbYes Then

                    'clear old responses and elements from memory, and remove eForm lock
                    Call moSubject.CancelResponses(EFI)
                    Unload Me
                End If
            End If
        Else
            'clear old responses and elements from memory, and remove eForm lock
            Call moSubject.CancelResponses(EFI)
            Unload Me
        End If
    Else
        'clear old responses and elements from memory, and remove eForm lock
        Call moSubject.RemoveResponses(EFI, True)
        Unload Me
    End If
    
    
    HourglassOff
    
End Sub

'---------------------------------------------------------------------
Private Sub EFIReturn()
'---------------------------------------------------------------------
' Return pressed.
'---------------------------------------------------------------------
Dim bCancel As Boolean

    HourglassOn
    
    If SaveResponses(eEFormAction.eaReturn, True, bCancel) Then
        ' Clear old responses and elements from memory, and remove eForm lock
        ' NCJ 10 Apr 03 - May need to cancel instead
        If bCancel Then
            Call moSubject.CancelResponses(EFI)
        Else
            Call moSubject.RemoveResponses(EFI, True)
        End If
        
        HourglassOff
    
        Unload Me
    Else
        ' Revalidation stopped them
        HourglassOff
    End If
    
    
End Sub

'---------------------------------------------------------------------
Private Sub GotoEFI(enEFormAction As eEFormAction, Optional vTaskId As Variant = 0)
'---------------------------------------------------------------------
' Next/Previous pressed.
' vTaskId could be Visit Taskid or CRFPageTaskid or text Hotlink spec
' NCJ 6 Jan 03 - Don't worry about not being able to save if user hasn't done anything in eForm
' NCJ 10 Feb 03 - Check that a target eForm can be opened by user
' NCJ 10 Apr 03 - Check whether responses should be cancelled
' NCJ 26 Jan 06 - vTaskId is now variant (may be a Hotlink spec) for Bug 2636
'---------------------------------------------------------------------
Dim sFirstLast As String
Dim oNewEFI As EFormInstance
Dim sMsg As String
Dim bChangeData As Boolean
Dim bCancelResponses As Boolean
Dim leFormTaskId As Long

Const sLAST_FORM = "This is the last eForm in this visit"
Const sFIRST_FORM = "This is the first eForm in this visit"

    On Error GoTo ErrLabel
    
    HourglassOn
    
    ' DPH 12/10/2001 - Disallow form movement to new form if no write access / visit locked or frozen
    bChangeData = moUser.CheckPermission(gsFnChangeData) And Not moSubject.ReadOnly And EFI.VisitInstance.LockStatus = eLockStatus.lsUnlocked

    ' Try to save if necessary and if user can change data
    If bChangeData Then
        ' Move to previous/next form
        ' NCJ 10 Apr 03 - Added bCancelResponses
        If Not SaveResponses(enEFormAction, mbUserDidSomething, bCancelResponses) Then
            ' Revalidation or saving failed
            ' If the user didn't do anything then don't worry about it,
            ' i.e. only get out if they DID change something
            If mbUserDidSomething Then
                HourglassOff
                Exit Sub
            End If
        End If
    End If
    
    Select Case enEFormAction
    Case eEFormAction.eaPrevious
        sMsg = sFIRST_FORM
        Set oNewEFI = moSubject.GetNextForm(EFI, False, False, bChangeData)
        
    Case eEFormAction.eaNext
        sMsg = sLAST_FORM
        Set oNewEFI = moSubject.GetNextForm(EFI, True, True, bChangeData)
        If Not oNewEFI Is Nothing Then
            If oNewEFI.VisitInstance.Visit.VisitEFormByEForm(oNewEFI.eForm).Repeating _
                    And oNewEFI.Status = eStatus.Requested And oNewEFI.CycleNo <> 1 Then
                sMsg = "Would you like to open a new " & oNewEFI.Name & " form?"
                If DialogQuestion(sMsg) = vbYes Then
                    ' We'll keep this new EFI
                Else
                    ' NCJ 27 Nov 02 - Need to reset message for later in case GetNextForm returns Nothing
                    sMsg = sLAST_FORM
                    Set oNewEFI = moSubject.GetNextForm(EFI, True, False, bChangeData)
                End If
            End If
        End If
        ' NCJ 17 Jan 03 - If there's no next form, offer to go to the next visit
        If oNewEFI Is Nothing Then
            sMsg = sMsg & vbCrLf & "Would you like to move to the next visit?"
            If DialogQuestion(sMsg) = vbYes Then
                Set oNewEFI = moSubject.GetFirstFormInNextVisit(EFI, bChangeData)
                ' Reset the message in case oNewEFI is still Nothing
                sMsg = "This is the last available eForm"
            Else
                ' Don't trouble them with any more messages
                sMsg = ""
            End If
        End If
        
    Case eEFormAction.eaGotoEFormInThisVisit
        Set oNewEFI = moSubject.eFIByTaskId(CLng(vTaskId))
        ' NCJ 10 Feb 03 - Need to check that user can open the eForm
        If Not CanOpenEform(oNewEFI, moSubject, moUser, sMsg) Then
            Set oNewEFI = Nothing
        End If

    Case eEFormAction.eaGotoThisEFormInVisit
        Set oNewEFI = moSubject.GeteFormInOtherVisit(EFI, CLng(vTaskId), True, bChangeData)
        sMsg = "There are no eForms in the selected visit that you may access"
        
    Case eEFormAction.eaGotoHotlink     ' NCJ 8 Nov 02
        ' Hotlink
        Set oNewEFI = Nothing
        ' NCJ 26 Jan 06 - rewritten to evaluate Hotlink here - Bug 2636
        ' vTaskId is text Hotlink spec
        leFormTaskId = Me.EFI.GetHotlinkTarget(CStr(vTaskId))
        If leFormTaskId = 0 Then
            sMsg = "The eForm associated with this link is not available"
        ElseIf leFormTaskId = Me.EFI.EFormTaskId Then
            sMsg = "This link specifies the current eForm"
        Else
            Set oNewEFI = moSubject.eFIByTaskId(leFormTaskId)
            ' NCJ 10 Feb 03 - Need to check that user can open the eForm
            If Not CanOpenEform(oNewEFI, moSubject, moUser, sMsg) Then
                Set oNewEFI = Nothing
            End If
        End If
        
    Case eEFormAction.eaGoBack     ' NCJ 17 Jan 03
        ' Go to eForm we just came from (assume it's allowed since they've just come from there!)
        If CLng(vTaskId) > 0 Then
            Set oNewEFI = moSubject.eFIByTaskId(CLng(vTaskId))
        Else
            Set oNewEFI = Nothing
        End If
    
    End Select
    
    ' if a new form is found this happens
    If Not (oNewEFI Is Nothing) Then
        ' clear old responses and elements from memory, and remove eForm lock
        ' NCJ 10 Apr 03 - Check whether to remove or cancel
        If bCancelResponses Then
            Call moSubject.CancelResponses(EFI)
        Else
            Call moSubject.RemoveResponses(EFI, True)
        End If
        ' Set up (load responses) and draw the new one, with revalidation
        ' NCJ 17 Jan 03 - Remember where we were (in case we want to return later)
        mlPreviousEformTaskId = EFI.EFormTaskId
        msPreviousEformName = EFI.Name
        If EFI.CycleNo > 1 Then
            msPreviousEformName = EFI.Name & "[" & EFI.CycleNo & "]"
        End If
        
        If SetUpNewEFI(oNewEFI, True) Then
            'notify external listeners that this form has had its EFI changed
            RaiseEvent EFIChange(oNewEFI.eForm.Study.StudyId, _
                oNewEFI.VisitInstance.Subject.Site, oNewEFI.VisitInstance.Subject.PersonId, oNewEFI.EFormTaskId)
        Else
            ' Couldn't load the new eForm's Responses
            ' so the best thing we can do now is go back to the Schedule
            Unload Me
        End If
    Else
        ' no form found
        ' NCJ 17 Jan 03 - Message might be empty if they declined offer to go to next visit
        If sMsg > "" Then
            DialogInformation sMsg
        End If
    End If
    
    HourglassOff

Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "EFINextPrevious", Err.Source) = Retry Then
        Resume
    End If

End Sub

'---------------------------------------------------------------------
Private Function SaveResponses(ByRef enEFormAction As eEFormAction, _
                    Optional ByVal bShowUserMessage As Boolean = True, _
                    Optional ByRef bCancelled As Boolean) As Boolean
'---------------------------------------------------------------------
' NCJ 24 Mar 03 - Changed to be a wrapper around DoSaveResponses so we can toggle user intervention
' NCJ 10 Apr 03 - Added bCancelled argument to tell us if the responses should be cancelled rather than removed
' ic 09/09/2005   added clinical coding
'---------------------------------------------------------------------

    Call ToggleUserIntervention(False)
    
    SaveResponses = DoSaveResponses(enEFormAction, bShowUserMessage, bCancelled)
    
    If gbClinicalCoding Then
        If SaveResponses Then
            'if the eform save went ok, and the user has some coding permissions, save coded responses
            If (moUser.CheckPermission(gsFnCodeClinicalResponse) Or moUser.CheckPermission(gsFnChangeClinicalStatus) _
            Or moUser.CheckPermission(gsFnValidateClinicalCode)) Then
                Call moSubject.SaveCodedResponses(EFI)
            End If
        End If
    End If
    
    Call ToggleUserIntervention(True)

End Function

'---------------------------------------------------------------------
Private Function DoSaveResponses(ByRef enEFormAction As eEFormAction, _
                                ByVal bShowUserMessage As Boolean, _
                                ByRef bCancelled As Boolean) As Boolean
'---------------------------------------------------------------------
' Save this eForm's responses
' and check if Registration is required
' NCJ 1 Oct 01 - Check non-MACRO Arezzo tasks too
' NCJ 5 Aug 02 - Perform revalidation and return FALSE if revalidation throws up new errors/warnings
' MLM 30/08/02: Disallow the saving of a form if a form date is required, but not yet entered.
' NCJ 18 Sept 02 - Saving may fail if we can't get a lock or if there's been a reload
' NCJ 18 Oct 02 - Added call to ChangeDoneSDVsToPlanned
' NCJ 6 Jan 03 - Only show user messages if bShowUserMessage = TRUE
' MLM 13/02/03: Made bShowUserMessage ByVal so that it can have a different value from mbUserDidSomething
'               Added enEFormAction parameter; use to control what options are displayed to user.
'               If saving is cancelled due to a missing form or visit date, focus to the relevant element.
' NCJ 24 Mar 03 - Changed name from SaveResponses to DoSaveResponses and created new wrapper SaveResponses
' NCJ 26 Mar 03 - Added saving of patient data when AREZZO tasks are done
' NCJ 31 Mar 03 - Bug 1431 - Ensure eForm is cancelled if form/visit date is blank and we're leaving the eForm
' NCJ 10 Apr 03 - Added bCancelled argument to tell us if the responses should be cancelled rather than removed
'---------------------------------------------------------------------
Dim bSaveOK As Boolean
Dim oFormDateResponse As Response
Dim sLockErrMsg As String
Dim enSaveResponsesResult As eSaveResponsesResult
Dim bResponsesChanged As Boolean
Dim bArezzoTasksDone As Boolean
Dim oVisitEFI As EFormInstance
Dim bNeedToSave As Boolean
'MLM 14/02/03:
Dim oElement As eFormElementRO
Dim bEFormAndVisitDatesBlank As Boolean

    On Error GoTo ErrLabel
    
    bSaveOK = True
    bCancelled = False
    

    ' DPH 12/10/2001 - Disallow Save if no permissions to do so
    If moUser.CheckPermission(gsFnChangeData) And Not moSubject.ReadOnly And Not EFI.ReadOnly Then

        ' NCJ 11 Feb 03 - Check changed responses for visit EFI
        bNeedToSave = False
        Set oVisitEFI = EFI.VisitInstance.VisitEFormInstance
        If Not oVisitEFI Is Nothing Then
            bNeedToSave = oVisitEFI.Responses.Changed
        End If
        
        ' NCJ 11 Feb 03 - Check changed responses for main EFI
        bNeedToSave = bNeedToSave Or EFI.Responses.Changed
        ' NCJ 5 Aug 02 Call Revalidation and only continue if all OK
        If bNeedToSave Then
            bSaveOK = RevalidateMe(False)
        End If
        
        'MLM 30/08/02: Check if the form has had the necessary form and visit dates entered
        If bSaveOK Then
            bEFormAndVisitDatesBlank = EFormAndVisitDatesBlank(bShowUserMessage, enEFormAction, oElement)
            bSaveOK = Not bEFormAndVisitDatesBlank
        End If
        
        'MLM 13/02/03: Check mandatory Qs here, and cancel saving if user wants to
        If bSaveOK And Not EFI.Complete And bShowUserMessage And enEFormAction <> eaSave Then
            bSaveOK = (DialogQuestion("Some mandatory questions are blank. Are you sure you want to leave this eForm?") = vbYes)
        End If
        
        If bSaveOK Then
            HourglassOn
            ' NCJ 18 Oct 02 - Remember if any actual response values have changed (so we can deal with SDVs)
            bResponsesChanged = EFI.Responses.ValuesChanged
    
            ' NCJ 10 Feb 03 - Added frmHourGlass around SaveResponses if we think we need to save
            If bNeedToSave Then
                frmHourglass.Display "Saving eForm", True
            End If
            
            ' Save anyway (just in case we missed something in bNeedToSave!)
            enSaveResponsesResult = moSubject.SaveResponses(EFI, sLockErrMsg)
            
            If bNeedToSave Then
                UnloadfrmHourglass
            End If
            
            Select Case enSaveResponsesResult
            Case eSaveResponsesResult.srrSuccess
                ' OK
                ' Reset "user did something since last save"
                mbUserDidSomething = False
                ' NCJ 18 Oct 02 - Reset SDVs if necessary
                ' NCJ 22 Nov 04 - Issue 2451 - Do not check SDV status here
'                If bResponsesChanged And moSubject.SDVStatus <> ssNone Then
                If bResponsesChanged Then
                    Call ChangeDoneSDVsToPlanned(gsADOConnectString, goUser.UserName, goUser.UserNameFull, _
                                                moSubject, EFI, GetMIMsgSource)
                End If
                
                'raise evnet that the subject label has changed if it has and reset variable
                If mbSubjectLabelChanged Then
                    RaiseEvent SubjectLabelChanged
                    mbSubjectLabelChanged = False
                End If
                
            Case eSaveResponsesResult.srrNoLockForSaving
                'failed because of locking TA 15/09/2002
                DoSaveResponses = False
                HourglassOff
                If bShowUserMessage Then
                    DialogError sLockErrMsg
                End If
'EXIT FUNCTION HERE
                Exit Function
            
            Case eSaveResponsesResult.srrSubjectReloaded
                ' There's been a reload so we're going to have to start again...
                DoSaveResponses = False
                HourglassOff
                ' NCJ 6 Jan 03 - Only do this if user is interested i.e. if bShowUserMessage is TRUE
                If bShowUserMessage Then
                    DialogWarning "Another user has changed this subject's data. Please try again."
                    ' NCJ 8 Jul 04 - Use non-refresh version of Update Responses
                    'Call moFormBuilder.UpdateResponses(ReloadingSubjectData)
                    Call moFormBuilder.UpdateSkipsDerivs(ReloadingSubjectData)
                    ' NCJ 26 Sept 02 - Need to refresh ALL responses (e.g. if Visit Date has changed)
                    Call moFormBuilder.DisplayCurrentResponses
                End If
'EXIT FUNCTION HERE
                Exit Function
            End Select
    
            ' NCJ 18 Jun 03 - Check Register Subject permission
            If goUser.CheckPermission(gsFnRegisterSubject) Then
                If TimeToRegister(EFI) Then
                    ' NCJ 11 Jan 06 - Issue 2619 - Must also check eForm is complete
                    If EFI.Complete Then
                        Call RegisterSubject
                        ' NCJ 13 May 03 - This enables /disables the Task List "Register Subject"
                        Call frmMenu.EnableDisableTaskListItems
                    End If
                End If
            End If
            
            HourglassOff
            
            ' Check for Arezzo tasks
            ' NCJ 31 Jan 03 - Collect Boolean result
            bArezzoTasksDone = False
            
        ' NCJ 9 Nov 04 - Conditional for ORAMA
        #If ORAMA = 1 Then
            'Special ORAMA window
            bArezzoTasksDone = frmORAMA.Display(moSubject.Arezzo, _
                                    Me.Top + frmMenu.Top + (frmMenu.Height - frmMenu.ScaleHeight) / 2, _
                                    Me.Left + frmMenu.Left + (frmMenu.Width - frmMenu.ScaleWidth) / 2, _
                                    fraCRFPage.Width, fraCRFPage.Height)
        #Else
        
            Call DealWithArezzoTasks(moSubject.Arezzo, Nothing, bArezzoTasksDone)
            
        #End If
        
            If bArezzoTasksDone Then
                ' Need to save AREZZO state
                Call moSubject.Save     ' Ignore errors...
            End If
            
            ' NCJ 30 Sept 03 - Just testing randomisation
            ' NCJ 7 Jan 04 - Add it in for real!
            Dim colRResults As Collection
            Call RandomiseSubject(goUser.CurrentDBConString, moSubject, colRResults)
            Set colRResults = Nothing
    
        Else
            ' It wasn't OK to save (so bSaveOK = FALSE)
        End If
    End If
    
    If bEFormAndVisitDatesBlank Then
        If oElement Is Nothing Then
            ' This means we're leaving the eForm but not saving
            DoSaveResponses = True          'leave the form without having saved
            ' NCJ 31 Mar 03 Bug 1431 - But we must Cancel the responses if we're leaving!
            ' NCJ 10 Apr 03 - Don't Cancel here but just set the Boolean
            ' Call moSubject.CancelResponses(EFI)
            bCancelled = True
        Else
            'focus to the specified element, and stay on the form
            FocusToAppropriateField moFormBuilder.EFormItemByIndex(moFormBuilder.IndexForElementRow(oElement, 1)), True
            DoSaveResponses = False
        End If
    Else
        DoSaveResponses = bSaveOK
    End If
    
Exit Function
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "DoSaveResponses", Err.Source) = Retry Then
        Resume
    End If

End Function

'---------------------------------------------------------------------
Public Sub eFormAction(enEFormAction As eEFormAction, Optional vParam As Variant)
'---------------------------------------------------------------------
' The toolbar buttons
' NCJ 28 Sep 01 - Check for unvalidated fields first
'---------------------------------------------------------------------
Dim lEFITaskId As Long
Dim lVisitTaskId As Long
    
    ' Say that we want be here
    mbMouseDownInNextField = True
    
    If ThereIsInvalidData Then
        mbMouseDownInNextField = False
        Exit Sub
    End If
    
    mbMouseDownInNextField = False
    Select Case enEFormAction
        Case eEFormAction.eaCancel
            ' Cancel
            Call EFICancel
        Case eEFormAction.eaReturn
            Call EFIReturn
            
        Case eEFormAction.eaPrint
            Call PrintEForm(EFI)
            
        Case eEFormAction.eaPrevious
            Call GotoEFI(eEFormAction.eaPrevious)
        
        Case eEFormAction.eaNext
            Call GotoEFI(eEFormAction.eaNext)
        Case eEFormAction.eaSave
            Call SaveResponses(eEFormAction.eaSave)
            ' NCJ 10 Jun 02 - CBB 2.2.13/63 Refresh display for all responses
            Call moFormBuilder.DisplayCurrentResponses
    
        Case eEFormAction.eaGotoEFormInThisVisit
            lEFITaskId = vParam
            Call GotoEFI(eEFormAction.eaGotoEFormInThisVisit, lEFITaskId)
            
        Case eEFormAction.eaGotoThisEFormInVisit
            lVisitTaskId = vParam
            Call GotoEFI(eEFormAction.eaGotoThisEFormInVisit, lVisitTaskId)

        ' NCJ 11 Nov 02 - Hotlinks
        Case eEFormAction.eaGotoHotlink
'            lEFITaskId = vParam
            ' NCJ 26 Jan 06 - Pass through unevaluated Hotlink
            Call GotoEFI(eEFormAction.eaGotoHotlink, vParam)
            
    End Select
    

End Sub


'---------------------------------------------------------------------
Private Sub txtCRFElement_Change(Index As Integer)
'---------------------------------------------------------------------
' NCJ 30 Oct 02 - Set the "dirty" flag (to ensure we recognise the "Delete" key)
' NCJ 20 Jan 03 - In MACRO 3.0, must ensure field is enterable
' NCJ 18 Jan 05 - Issue 2502 - Reset popup list for category question according to what's typed
'---------------------------------------------------------------------
Dim oItem As EFormItem

    Set oItem = moFormBuilder.EFormItemByIndex(Index)
    ' Only set the dirty flag for an enterable response
    If Not oItem.Response Is Nothing Then
        If oItem.Response.Enterable Then
            mbDirtyField = True
            ' NCJ 18 Jan 05 - Rebuild popup list (but not in response to arrow keys, or if not enabled)
            If mbEnableListFiltering And oItem.Element.ControlType = gn_POPUP_LIST And Not mbArrowInDropDown Then
                Call moFormBuilder.CreatePopupSubList(oItem.Element, Trim(txtCRFElement(Index).Text))
            End If
            mbArrowInDropDown = False
    
        End If
    End If
    
    Set oItem = Nothing
    
End Sub

'---------------------------------------------------------------------
Private Sub txtCRFElement_GotFocus(Index As Integer)
'---------------------------------------------------------------------
' Focus received by text box
'
' MLM 30/08/02: Treat Locked controls like disabled ones.
'---------------------------------------------------------------------
Dim oItem As EFormItem
Dim oTextBox As TextBox

    Set oTextBox = txtCRFElement(Index)
    
    ' NB It can happen that this event arrives while
    ' there's a modal dialog on top so check for Enabled
    If Not oTextBox.Enabled Then Exit Sub
    
    Set oItem = moFormBuilder.EFormItemByIndex(Index)
    
    'MLM 21/05/02: Redraw everything for the current question unless we've just come from picCRFPage_GotFocus()
    '   because this causes problems with the drop-down list.
    If mbCRFPageClicked Then
        mbCRFPageClicked = False
    Else
        ' Fill in the Question Information Panel
        Call moFormBuilder.ElementGotFocus(oItem)
    End If
    
    If Not oTextBox.Locked Then
        oTextBox.BackColor = glYELLOW
    End If
    'set the colour of the notestatus image accordingly
    picCommentStatus(Index).BackColor = oTextBox.BackColor
    picNoteStatus(Index).BackColor = oTextBox.BackColor
    
    With oTextBox
        .Tag = .Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
    ' Remember the last selected control
    Set moLastSelectedControl = oTextBox

    Set oTextBox = Nothing
    Set oItem = Nothing

    'NCJ 1 Feb 02 - scroll if off screen
    Call ScrollIfOffScreen(txtCRFElement(Index))

End Sub

'---------------------------------------------------------------------
Private Sub txtCRFElement_LostFocus(Index As Integer)
'---------------------------------------------------------------------
' Focus lost by text box
'
' MLM 30/08/02: Set the control's background colour differently, depending on its enablement.
'---------------------------------------------------------------------
Dim oItem As EFormItem
Dim oTextBox As TextBox
Dim bStayOnField As Boolean

    'TA 09/04/2002 restart timeout timer at appropriate intervals
    
    Call RestartSystemIdleTimer

    Set oItem = moFormBuilder.EFormItemByIndex(Index)


    ' If we're already validating this one then don't do it again
    If mnValidatingResponse = oItem.Element.ElementID Then
        Exit Sub
    End If
    
    Set oTextBox = txtCRFElement(Index)
    
    ' NCJ 7 Jun 02 (CBB2.2.12/41) Assume we want to stay put
    bStayOnField = True
    
    If oTextBox.Text <> oTextBox.Tag Then
        If ValidateProposedValue(oTextBox.Text, oItem) Then
            ' We can move on OK

'03/12/2001: VTRACK Changes  - Text boxes have windows background
#If VTRACK <> 1 Then
            'MLM 30/08/02: Set appropriate back colour for new screen design
            If oTextBox.Locked Then
                 oTextBox.BackColor = glLIGHT_GREY
            Else
                oTextBox.BackColor = vbWhite ' Me.picCRFPage.BackColor
            End If
#Else
            oTextBox.BackColor = vbWindowBackground
#End If
            'set the colour of the notestatus image accordingly
            picNoteStatus(Index).BackColor = txtCRFElement(Index).BackColor
            picCommentStatus(Index).BackColor = txtCRFElement(Index).BackColor

            ' Stay on this field if Warning (NCJ 7 Jun 02, CBB 2.2.12/41)
            bStayOnField = (oItem.Response.Status = eStatus.Warning)
            
        End If
        'MLM 27/03/03: If reject or warning, ignore what user is doing and focus to current field
        'Otherwise, focus to the next field provided it doesn't hamper their navigation.
        If bStayOnField Or (FocusToNextFieldOK And Not mbMouseDownInNextField) Then
            Call FocusToAppropriateField(oItem, bStayOnField)
        End If
    
    Else
        ' Nothing's changed
        ' Reset Dirty flag (in case it was left set by a Reject)
        mbDirtyField = False
        
'03/12/2001: VTRACK Changes  - Text boxes have windows background
#If VTRACK <> 1 Then
        'MLM 30/08/02: Set appropriate back colour for new screen design
        If oTextBox.Locked Then
            oTextBox.BackColor = glLIGHT_GREY
        Else
            oTextBox.BackColor = vbWhite ' Me.picCRFPage.BackColor
        End If
#Else
        oTextBox.BackColor = vbWindowBackground
#End If
            'set the colour of the notestatus image accordingly
            picNoteStatus(Index).BackColor = oTextBox.BackColor
            picCommentStatus(Index).BackColor = oTextBox.BackColor
        
    End If
    
    mbMouseDownInNextField = False
    
    ' Get rid of popup list and/or calendar
    Call HidePopupAndCalendar
    
    Set oItem = Nothing
    
End Sub

'---------------------------------------------------------------------
Private Sub txtCRFElement_KeyPress(Index As Integer, KeyAscii As Integer)
'---------------------------------------------------------------------
' Record that there's a dirty field on the form if not CR (see SR2260)
'---------------------------------------------------------------------
Dim oItem As EFormItem
Dim bStayOnField As Boolean
    
    ' NCJ 7 Jun 02 (CBB2.2.12/41) Assume we want to stay put
    bStayOnField = True
    
    If KeyAscii = Asc(vbCr) Then
        ' Deal with RETURN key
        KeyAscii = 0
        ' Pick up the current item
        Set oItem = moFormBuilder.EFormItemByIndex(Index)
        If txtCRFElement(Index).Text <> txtCRFElement(Index).Tag Then
            If ValidateProposedValue(txtCRFElement(Index).Text, oItem) Then

                ' NCJ 7 Jun 02 We stay where we are if we got a warning (CBB 2.2.12/41)
                bStayOnField = (oItem.Response.Status = eStatus.Warning)
'                ' We can move on OK
            End If
        Else
            ' Nothing's changed so move on
            bStayOnField = False
        End If
        Call FocusToAppropriateField(oItem, bStayOnField)

        ' Get rid of popup list and/or calendar
        Call HidePopupAndCalendar
    
        Set oItem = Nothing
    
    Else
        ' Not the RETURN key
        mbDirtyField = True
    End If


End Sub

'---------------------------------------------------------------------
Private Sub HidePopupAndCalendar()
'---------------------------------------------------------------------
' Hide the popuplist and/or calendar if showing
'---------------------------------------------------------------------
        
    ' Get rid of popup list and/or calendar
    If moFormBuilder.PopupListVisible Then
        Call moFormBuilder.HidePopupList
    End If
    
    If calCRFElement.Visible = True Then
        calCRFElement.Visible = False
    End If

End Sub

'---------------------------------------------------------------------
Private Sub txtCRFElement_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' Store that the user is clicking in a new field
' (so we won't try and jump elsewhere after validating current field)
'---------------------------------------------------------------------

    mbMouseDownInNextField = True
    
End Sub

'---------------------------------------------------------------------
Private Sub vsbCRFPage_Change()
'---------------------------------------------------------------------
' Scroll the picture box if the value of scrollbar has been changed
' Then reset focus to last selected control
'---------------------------------------------------------------------

    picCRFPage.Top = CSng(-vsbCRFPage.Value) * 10
 'TA 14/1/02: removed line below as I don't think it is needed and it mucks up scrolling
'    Call FocusToLastSelectedControl
    picHeaderExt.Top = picCRFPage.Top

    
End Sub

'---------------------------------------------------------------------
Private Sub vsbCRFPage_Scroll()
'---------------------------------------------------------------------
' Scroll the picture box continuously
'---------------------------------------------------------------------

    picCRFPage.Top = CSng(-vsbCRFPage.Value) * 10
    picHeaderExt.Top = picCRFPage.Top
   
End Sub

'---------------------------------------------------------------------
Private Function ValidateProposedValue(ByVal sValue As String, oItem As EFormItem) As Boolean
'---------------------------------------------------------------------
' Process the proposed value for the given Element and Response
' Handles all the error messages, Overrules, RFCs etc.,
' and displays the correct response value at the end
' Returns TRUE if value was accepted
' or FALSE if value was rejected
' Assume sValue is a changed value for the element
'
' MLM 09/09/02: Pass in an EFormItem instead of Element and Response.
' NCJ 8 Jan 03 - Get user full name for Authorisation
' NCJ 10 Feb 03 - Reset mbDirtyField at the end of this routine
' ic 13/09/2005 added clinical coding
' NCJ 28 Mar 07 - Bug 2896 - Keep dirty flag set for rejected text fields
'---------------------------------------------------------------------
Dim nStatus As Integer
Dim sErrMsg As String
Dim bSuccess As Boolean
Dim sOverrule As String
Dim sRFC As String
Dim sAuthorizerName As String
Dim sAuthorizerNameFull As String
Dim sRFCCaption As String
Dim bChanged As Boolean
Dim oResponse As Response
Dim bKeepDirtyFlagSet As Boolean

    On Error GoTo ErrLabel
    
    ' Store that we're processing this element
    mnValidatingResponse = oItem.Element.ElementID
    
    ' User must have done something to get here
    mbUserDidSomething = True
    
    bSuccess = True
    
    bKeepDirtyFlagSet = False
    
    sOverrule = ""
    sRFC = ""
    sAuthorizerName = ""
    
    ' If it's a labtest question, need to ascertain the Lab code
    If oItem.Element.DataType = eDataType.LabTest Then
        ' Ask for and store a lab code as necessary
        Call AskLabCode
    End If
    
    Set oResponse = oItem.Response
    nStatus = oResponse.ValidateValue(sValue, sErrMsg, bChanged)

    ' If warning, see if they want to overrule (only if value's really changed)
    If nStatus = eStatus.Warning And bChanged Then
        'Ask user for override
        ' NCJ 15 Aug 02 - New parameters
        Call frmWarnOverrule.Display(oResponse, nStatus, sValue, sErrMsg, sOverrule, GetCoordForForm(oItem))
    End If
    
    ' Now deal with Authorisation and RFC
    Select Case nStatus
        Case eStatus.Success, eStatus.Warning, eStatus.Inform, eStatus.Missing, eStatus.OKWarning
            If bChanged Then    ' only if it really changed
                'check if changes require authorisation
                If oItem.Element.Authorisation > "" Then
                    ' Here the authorization window is being called
                    ' NCJ 8 Jan 03 - Added extra sAuthorizerNameFull argument
                    'TA 20/06/2005: added clinicaltialid and site to call - cbd 2583
                    sAuthorizerName = frmTempLogin.Display(oItem.Element.Authorisation, moUser.Database.DatabaseCode, moSubject.StudyId, moSubject.Site, sAuthorizerNameFull)
                    ' Check user supplies userid and password
                    bSuccess = Not (sAuthorizerName = "")
                End If
                ' NCJ 13 Aug 02 - Use new RequiresValueRFC of Response class
'                If oElement.RequiresRFC And bSuccess Then
                If oResponse.RequiresValueRFC And bSuccess Then
                    'Display a RFC window and ask user to select one from there
                    sRFCCaption = "Please enter or choose the reason for changing "
                    frmRFC.lblRFC.Caption = sRFCCaption & oItem.Element.Name
                    frmRFC.DisplayRFC mcolRFCs, GetCoordForForm(oItem)
                    sRFC = frmRFC.GetRFC
                    If Len(sRFC) < 1 Then
                        bSuccess = False
                    End If
                End If
            End If

        Case eStatus.InvalidData
            ' Display a reject message
            ' NCJ 15 Aug 02 - New parameters
            Call frmWarnOverrule.Display(oResponse, nStatus, sValue, sErrMsg, sOverrule, GetCoordForForm(oItem))
            bSuccess = False
        Case Else
            'TA 08/04/2003: this should never happen unless there is a bug
            Call DialogError("Unexpected status value " & nStatus & ".")
    End Select
    
    ' If value is OK, confirm it, otherwise reject it
    If bSuccess Then
        If bChanged Then
            'TA 08/04/2003: only confirm if the vlaue has changed
            Call oResponse.ConfirmValue(sOverrule, sRFC, sAuthorizerName, sAuthorizerNameFull)
            ' We've now dealt successfully with this field
            ' NCJ 12 Jun 02 - CBB 2.2.15/1 Leave the dirty flag set if it's already set
            ' and we have a Warning (don't touch it if not set, e.g. not a text field)
            If mbDirtyField Then
                ' NCJ 10 Feb 03 - Store this value in a separate DirtyFlag variable
                '(Do not set directly here because mbDirtyField can get splatted by DisplayCurrentResponse)
                bKeepDirtyFlagSet = (oResponse.Status = eStatus.Warning)
            End If
            ' NCJ 9 Dec 02 - Must check for new group row here
            Call moFormBuilder.NewGroupRow(oItem.Element, oItem.RowNo)
        End If

    Else
        Call oResponse.RejectValue
        ' And leave the Dirty flag set
        ' NCJ 28 Mar 07 - Bug 2896 - Set bKeepDirtyFlagSet for text fields
        If mbDirtyField Then
            bKeepDirtyFlagSet = True
        End If
    End If

    ValidateProposedValue = bSuccess
    
    ' Display the current response in case it changed
    moFormBuilder.DisplayCurrentResponse oItem, True
    
    ' Update with derivations etc.
    ' but only if value was accepted
    If bChanged And bSuccess Then
        Call moFormBuilder.UpdateResponses(ChangingResponse, oResponse)
        ' NCJ 31 Oct 02 - Display popups for derived warnings/rejects
        If ShowDerivedWarnings Then
            ' The user did something like overruling a warning on a derived question
            ' so update the eForm
            Call moFormBuilder.DisplayCurrentResponses
        End If
        
    'ic 15/08/2005 question has been changed - reset coding status
    Call ResetCodingStatus(oItem)
        
    End If
    
    ' Store that we've finished processing this element
    mnValidatingResponse = 0

    ' NCJ 10 Feb 03 - Now reset the dirty flag as we decided earlier
    mbDirtyField = bKeepDirtyFlagSet
    
    Set oResponse = Nothing

Exit Function
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, _
                    "ValidateProposedValue(" & sValue & ")", Err.Source) = Retry Then
        Resume
    End If

End Function

'---------------------------------------------------------------------
Private Sub MoveScrollBar(ByRef KeyCode As Integer, _
                          ByVal TabIndex As Integer)
'---------------------------------------------------------------------
' Move scroll bar
'---------------------------------------------------------------------
    
  Select Case KeyCode
        Case vbKeyPageUp
            If vsbCRFPage.Value < vsbCRFPage.LargeChange Then
                vsbCRFPage.Value = 0
            Else
                vsbCRFPage.Value = vsbCRFPage.Value - vsbCRFPage.LargeChange
            End If
        Case vbKeyPageDown
            If vsbCRFPage.Value + vsbCRFPage.LargeChange > vsbCRFPage.Max Then
                vsbCRFPage.Value = vsbCRFPage.Max
            Else
                vsbCRFPage.Value = vsbCRFPage.Value + vsbCRFPage.LargeChange
            End If
        Case vbKeyUp
            If vsbCRFPage.Value < vsbCRFPage.SmallChange Then
                vsbCRFPage.Value = 0
            Else
                vsbCRFPage.Value = vsbCRFPage.Value - vsbCRFPage.SmallChange
            End If
            KeyCode = 0 'Stops the cursor moving left and right in a text box
        Case vbKeyDown
            If vsbCRFPage.Value + vsbCRFPage.SmallChange > vsbCRFPage.Max Then
                vsbCRFPage.Value = vsbCRFPage.Max
            Else
                vsbCRFPage.Value = vsbCRFPage.Value + vsbCRFPage.SmallChange
            End If
            KeyCode = 0 'Stops the cursor moving left and right in a text box
    End Select
   
End Sub

'---------------------------------------------------------------------
Private Function FocusToNextFieldOK() As Boolean
'---------------------------------------------------------------------
' Return TRUE if currently active control is a non-data-entry field
' REM 17/05/02 - added cmdNextButton
' MLM 27/03/03: Renamed and revamped. Return True if automatically focussing to the
'   next field is feasible.
'---------------------------------------------------------------------
    
    Select Case frmMenu.ActiveForm.Name
    Case "frmEFormDataEntry"
        Select Case Me.ActiveControl.Name
        ' Include names of form's button controls
        Case "cmdPopupList", "cmdAttachment", "cmdNextButton", "picBackButton"
            FocusToNextFieldOK = False
        Case Else
            FocusToNextFieldOK = True
        End Select
    Case Else
        FocusToNextFieldOK = False
    End Select

End Function

'---------------------------------------------------------------------
Private Sub imgStatus_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' Check for right click on status image to display popup menu
' NCJ 15 Nov 01 - Main code moved to ElementRightMouseMenu
'---------------------------------------------------------------------
Dim oItem As EFormItem

    On Error GoTo ErrLabel
    
    If Button <> vbRightButton Then Exit Sub
    
    ' Remember where we are
    Set oItem = moFormBuilder.CurrentItem
    
    ' Check for unvalidated data first
    If ThereIsInvalidData Then
        Exit Sub
    End If
    
    Call ElementRightMouseMenu(moFormBuilder.EFormItemByIndex(Index))
        
    ' Set focus back to where we were
    If Not oItem Is Nothing Then
        'ignore erros (form might be disabled
        On Error Resume Next
        Call moFormBuilder.ResetFormFocus(oItem)
        On Error GoTo ErrLabel
    End If
    
    Set oItem = Nothing
    
Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "imgStatus_MouseUp", Err.Source) = Retry Then
        Resume
    End If

End Sub

'---------------------------------------------------------------------
Private Sub ElementRightMouseMenu(oItem As EFormItem, Optional bUseMouseCoords As Boolean = True)
'---------------------------------------------------------------------
' Display right mouse menu for an eForm Item
' NCJ 25 Sept 02 - Take account of read-only forms
'   TA 04/10/2002: Removed dataitem panel and added consequently removed functionality to here
' NCJ 18 Feb 03 - Remember the scroll position and reset it after displaying whatever
' ic 28/07/2005 added clinical coding
'---------------------------------------------------------------------
Dim oElement As eFormElementRO
Dim sQuestion As String
Dim sItem As String
Dim sMenuItems As String
Dim sEnabled As String
Dim sOverruleReason As String
Dim bQuNotLockedorFrozen As Boolean
Dim bQuNotFrozen As Boolean
Dim oResponse As Response
Dim oMenuItems As clsMenuItems
Dim bHasResponse As Boolean
Dim sglTop As Single
Dim sglLeft As Single
Dim nHSBValue As Integer
Dim nVSBValue As Integer

    On Error GoTo ErrLabel
    
    'TA exit if this is not a question
    If oItem Is Nothing Then Exit Sub
    'this is definately not a question
    
    Set oElement = oItem.Element
    
    If oElement.QuestionId = 0 Then
        'this is not a question - exit
        Exit Sub
    End If

    Set oResponse = oItem.Response 'moEFI.Responses.ResponseByElement(oElement, oItem.RowNo)
    
    bHasResponse = Not oResponse Is Nothing
    
    If bHasResponse Then
        ' Store whether question can be accessed
        bQuNotLockedorFrozen = (oResponse.LockStatus = LockStatus.lsUnlocked)
        bQuNotFrozen = (oResponse.LockStatus <> LockStatus.lsFrozen)
    End If

    Set oMenuItems = New clsMenuItems
    Call oMenuItems.Add("QUES", "View question details...", True)
    Call oMenuItems.Add("QUDEF", "View question definition...", True)
    Call oMenuItems.Add("AUDIT", "View audit trail...", False, False)    ' Enable it later
    
    'ic 13/09/2005 add coding audit trail option
    If gbClinicalCoding Then
        Call oMenuItems.Add("CODEAUDIT", "View coding audit trail...", False, False)
        oMenuItems.KeyedItem("CODEAUDIT").Enabled = (oResponse.CodingStatus <> eCodingStatus.csEmpty)
    End If
    
    Call oMenuItems.AddSeparator
    
    Call oMenuItems.Add("WARN", "View warning...")          ' Enable it later
    If moUser.CheckPermission(gsFnMonitorDataReviewData) Then
        'TA 3/1/03: only show to monitors
        Call oMenuItems.Add("INFORM", "View inform message...")    ' Enable it later
    End If
    Call oMenuItems.AddSeparator
    
    ' Comments
    ' NCJ 25 Sept 02 - Instead disallow for read-only eForms
    Call oMenuItems.Add("ADDCOM", "New comment...", moUser.CheckPermission(gsFnAddIComment) _
                                                And Not oItem.EFI.ReadOnly _
                                                And bQuNotLockedorFrozen And bHasResponse)
    
    Call oMenuItems.Add("REMOVE", "Remove comments", moUser.CheckPermission(gsFnAddIComment) _
                                                And Not oItem.EFI.ReadOnly _
                                                And bQuNotLockedorFrozen And bHasResponse)
    Call oMenuItems.AddSeparator
    
' No MIMessages in VTRACK
#If VTRACK <> 1 Then
    Call oMenuItems.Add("NOTE", "New note...")          ' Enable it later
    Call oMenuItems.Add("VNOTE", "View notes...", False)         ' Enable it later
    Call oMenuItems.AddSeparator
    
    Call oMenuItems.Add("DISC", "New discrepancy...")   ' Enable it later
    Call oMenuItems.Add("VDISC", "View discrepancies...", False)         ' Enable it later
    Call oMenuItems.AddSeparator
        
    Call oMenuItems.Add("SDV", "New SDV mark...")       ' Enable it later
    Call oMenuItems.Add("VSDV", "View SDV mark...", False)         ' Enable it later
    
#End If

    'ic 25/07/2005 added clinical coding menu options
    If gbClinicalCoding Then
        If (oElement.DataType = eDataType.Thesaurus) Then
            Call oMenuItems.AddSeparator
            'set to validated or return to coded
            If oResponse.CodingStatus = eCodingStatus.csValidated Then
                Call oMenuItems.Add("CODESTATUSCODE", "Return coding status to Coded")
                oMenuItems.KeyedItem("CODESTATUSCODE").Enabled = moUser.CheckPermission(gsFnValidateClinicalCode)
            Else
                Call oMenuItems.Add("CODESTATUSVALID", "Set coding status to Validated")
                oMenuItems.KeyedItem("CODESTATUSVALID").Enabled = moUser.CheckPermission(gsFnValidateClinicalCode) _
                    And (oResponse.CodingStatus = eCodingStatus.csCoded _
                    Or oResponse.CodingStatus = eCodingStatus.csAutoEncoded)
            End If
                
            'set to pending or return to coded
            If oResponse.CodingStatus = eCodingStatus.csPendingNewCode Then
                Call oMenuItems.Add("CODESTATUSCODE", "Return coding status to Coded")
                oMenuItems.KeyedItem("CODESTATUSCODE").Enabled = moUser.CheckPermission(gsFnChangeClinicalStatus)
            Else
                Call oMenuItems.Add("CODESTATUSPEND", "Set coding status to Pending New Code")
                oMenuItems.KeyedItem("CODESTATUSPEND").Enabled = moUser.CheckPermission(gsFnChangeClinicalStatus) _
                    And (oResponse.CodingStatus = eCodingStatus.csCoded _
                    Or oResponse.CodingStatus = eCodingStatus.csAutoEncoded _
                    Or oResponse.CodingStatus = eCodingStatus.csValidated)
            End If
            
            'set to not coded
            Call oMenuItems.Add("CODESTATUSNOT", "Set coding status to Not Coded")
            oMenuItems.KeyedItem("CODESTATUSNOT").Enabled = moUser.CheckPermission(gsFnChangeClinicalStatus) _
                And (oResponse.CodingStatus <> eCodingStatus.csNotCoded And oResponse.CodingStatus <> eCodingStatus.csEmpty)
                
            'set to do not code
            Call oMenuItems.Add("CODESTATUSDONT", "Set coding status to Do Not Code")
            oMenuItems.KeyedItem("CODESTATUSDONT").Enabled = moUser.CheckPermission(gsFnChangeClinicalStatus) _
                And (oResponse.CodingStatus = eCodingStatus.csNotCoded Or oResponse.CodingStatus = eCodingStatus.csEmpty)
        End If
    End If

    If bHasResponse Then
        ' Enable items available for items with responses
        oMenuItems.KeyedItem("AUDIT").Enabled = moUser.CheckPermission(gsFnViewAuditTrail)
        oMenuItems.KeyedItem("WARN").Enabled = (oResponse.Status = eStatus.Warning) Or (oResponse.Status = eStatus.OKWarning)
        If moUser.CheckPermission(gsFnMonitorDataReviewData) Then
            oMenuItems.KeyedItem("INFORM").Enabled = oResponse.Status = eStatus.Inform
        End If
        
' No MIMessages in VTRACK
#If VTRACK <> 1 Then
        ' NCJ 18 Oct 02 - We allow MIMsgs on any non-frozen responses
        ' NCJ 29 Dec 06 - Bug 2861 - Check the Response's WillBeSaved property to avoid creating MIMsg for non-existent response
        oMenuItems.KeyedItem("NOTE").Enabled = bQuNotFrozen And oResponse.WillBeSaved
        oMenuItems.KeyedItem("DISC").Enabled = moUser.CheckPermission(gsFnCreateDiscrepancy) _
                                            And bQuNotFrozen And oResponse.WillBeSaved
        ' NCJ 5 Nov 02 - Only allow Create SDV if response doesn't already have an SDV
        oMenuItems.KeyedItem("SDV").Enabled = moUser.CheckPermission(gsFnCreateSDV) And bQuNotFrozen _
                                            And oResponse.SDVStatus = ssNone And oResponse.WillBeSaved
        
        oMenuItems.KeyedItem("VNOTE").Enabled = (oResponse.NoteStatus = nsHasNote)
        oMenuItems.KeyedItem("VDISC").Enabled = (oResponse.DiscrepancyStatus <> dsNone)
        oMenuItems.KeyedItem("VSDV").Enabled = (oResponse.SDVStatus <> ssNone)
        
#End If
        
        'TA 04/10/2002: Additional items needed as removed from side panel
        Call oMenuItems.AddSeparator
        Select Case oResponse.Status
            Case eStatus.Unobtainable
                Call oMenuItems.Add("MISS", "Change status to missing", oResponse.Enterable)
                Call oMenuItems.AddSeparator
             Case eStatus.Missing, eStatus.Requested
                Call oMenuItems.Add("UNOB", "Change status to unobtainable", oResponse.Enterable)
                Call oMenuItems.AddSeparator
        End Select

        ' The Clear button
        Call oMenuItems.Add("CLEAR", "Clear", (oResponse.Enterable And oResponse.Value > ""))
        ' The Change Attachment button
        If oElement.DataType = eDataType.Multimedia And oResponse.Enterable And oResponse.Value > "" Then
            Call oMenuItems.Add("ATTACH", "Change Attachment...", True)
        End If
        ' The Units button
        If oElement.Unit > "" And oResponse.Enterable Then
            Call oMenuItems.Add("UNIT", "Units...", True)
        End If

    End If
    
    ' NCJ 18 Feb 03 - Remember scroll position
    nVSBValue = vsbCRFPage.Value
    nHSBValue = hsbCRFPage.Value

    If bUseMouseCoords Then
        sItem = frmMenu.ShowPopUpMenu(oMenuItems)
    Else
        sItem = frmMenu.ShowPopUpMenu(oMenuItems, Me.Left + oItem.Element.ElementX + 2000, Me.Top + oItem.Element.ElementY)
    End If
   
    Select Case sItem
        Case "QUDEF": Call ShowQuestionDef(oItem)
        Case "AUDIT": Call ShowAuditTrail(oItem)
        Case "ADDCOM": Call ShowAddComment(oItem)
        Case "REMOVE": Call RemoveComment(oItem)
        Case "QUES": Call frmQuestionDetails.Display(oResponse, GetCoordForForm(oItem), moUser)
        Case "WARN": Call ShowWarning(oItem)
        Case "INFORM": Call ShowInform(oItem)
        Case "NOTE": Call NewMIMessage(oItem, MIMsgType.mimtNote)
        Case "DISC": Call NewMIMessage(oItem, MIMsgType.mimtDiscrepancy)
        Case "SDV": Call NewMIMessage(oItem, MIMsgType.mimtSDVMark)
        Case "MISS", "UNOB": Call ToggleMissingUnobtainable(oItem)
        Case "ATTACH": Call ChangeAttachment(oItem)
        Case "CLEAR": Call ClearQuestion(oItem)
        Case "UNIT": Call ChangeUnits(oItem)
        Case "VNOTE": Call ViewMIMessage(oItem, MIMsgType.mimtNote)
        Case "VDISC": Call ViewMIMessage(oItem, MIMsgType.mimtDiscrepancy)
        Case "VSDV": Call ViewMIMessage(oItem, MIMsgType.mimtSDVMark)
        
        'ic 26/07/2005 added clinical coding
        Case "CODEAUDIT": Call CodingAudit(oItem)
        Case "CODESTATUSVALID": Call ChangeCodingStatus(oItem, csValidated)
        Case "CODESTATUSCODE": Call ChangeCodingStatus(oItem, csCoded)
        Case "CODESTATUSPEND": Call ChangeCodingStatus(oItem, csPendingNewCode)
        Case "CODESTATUSNOT": Call ChangeCodingStatus(oItem, csNotCoded)
        Case "CODESTATUSDONT": Call ChangeCodingStatus(oItem, csDoNotCode)

    End Select
        
    'MLM 30/05/02: Redraw the screen if the pop-up menu was for the current question
    '   AND the user actually selected something from it.
    ' NCJ 18 Feb 03 - Reset the scroll bars (after allowing the ElementGotFocus to do its worst)
    If sItem <> "" And (Not moFormBuilder.CurrentItem Is Nothing) Then
        picCRFPage.AutoRedraw = False
        If oItem.ControlIndex = moFormBuilder.CurrentItem.ControlIndex Then
            Call moFormBuilder.ElementGotFocus(oItem)
        End If
        DoEvents    ' Allow scrolling back to focussed question first
        ' NCJ 18 Feb 03 - Now reset scroll position
        vsbCRFPage.Value = nVSBValue
        hsbCRFPage.Value = nHSBValue
        picCRFPage.AutoRedraw = True
    End If
    
    Set oElement = Nothing
    Set oResponse = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmEFormDataEntry.ElementRightMouseMenu"
    
End Sub

'---------------------------------------------------------------------
Private Sub CodingAudit(oItem As EFormItem)
'---------------------------------------------------------------------
' 'ic 13/09/2005
' Show the coding audit trail for the given oElement
'---------------------------------------------------------------------
'reference to AuditTrail form
Dim ofrmAuditTrail As frmWebNonMDI
Dim sglTop As Single
Dim sglLeft As Single
Dim sglHeight As Single

    On Error GoTo ErrLabel
    
    If oItem Is Nothing Then Exit Sub
    
    Call moFormBuilder.GetItemTopLeftHeight(oItem, sglTop, sglLeft, sglHeight)
    
    Set ofrmAuditTrail = New frmWebNonMDI
    
    ' NCJ 9 Nov 04 - Bug 2362 - Don't accept F7 etc. while we're showing the audit trail
    mbDontInterrupt = True
   
    With ofrmAuditTrail
        .Top = GetCoordForForm(oItem).Row
        .Left = GetCoordForForm(oItem).Col
        .Width = 8000
        .Display wdtHTML, CodingAuditTrailHTML(oItem.Response), "auto", True, "Audit Trail"
    End With

    ' NCJ 9 Nov 04 - Reset interruption flag
    mbDontInterrupt = False
    
    Set ofrmAuditTrail = Nothing
    
Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "ShowAuditTrail", Err.Source) = Retry Then
        Resume
    End If

End Sub

'---------------------------------------------------------------------
Private Sub ChangeCodingStatus(oItem As EFormItem, eNewStatus As eCodingStatus)
'---------------------------------------------------------------------
' ic 26/07/2005
' function called from popup menu to change coding status
' revisions
' ic 21/11/2006 bug 2843 - pass the current username in to setcodingstatus
'---------------------------------------------------------------------
Dim oResponse As Response
Dim oElement As eFormElementRO
Dim bOK As Boolean


    Set oElement = oItem.Element
    Set oResponse = oItem.Response

    Select Case eNewStatus
        Case eCodingStatus.csValidated, eCodingStatus.csDoNotCode, eCodingStatus.csCoded
            'status change to validated/do not code/coded - just change the status
            Call oResponse.SetCodingStatus(CInt(eNewStatus), goUser.UserName, goUser.UserNameFull)
            Call moFormBuilder.SetClinicalCodingImage(CInt(eNewStatus), imgDictionary(oItem.ControlIndex))
            
        Case eCodingStatus.csPendingNewCode
            'status change to pending new - warn the user that the existing code will be lost
            If DialogQuestion("The current coding details for this question will be overwritten - are you sure you wish to continue?") = vbYes Then
                Call oResponse.SetCodingStatus(CInt(eCodingStatus.csPendingNewCode), goUser.UserName, goUser.UserNameFull)
                Call moFormBuilder.SetClinicalCodingImage(CInt(eCodingStatus.csPendingNewCode), imgDictionary(oItem.ControlIndex))
            End If
            
        Case eCodingStatus.csNotCoded
            'status change to not coded - warn the user that the existing code, if any, will be lost
            bOK = False
            
            If oResponse.CodingStatus = eCodingStatus.csDoNotCode Then
                bOK = True
            Else
                If DialogQuestion("The current coding details for this question will be cleared - are you sure you wish to continue?") = vbYes Then
                    bOK = True
                End If
            End If
            If bOK Then
                'if user accepted, clear coding and dictionary details
                Call oResponse.SetCodingStatus(CInt(eCodingStatus.csNotCoded), goUser.UserName, goUser.UserNameFull)
                Call moFormBuilder.SetClinicalCodingImage(CInt(eCodingStatus.csNotCoded), imgDictionary(oItem.ControlIndex))
            End If
    End Select
    
    Set oElement = Nothing
    Set oResponse = Nothing
End Sub

'---------------------------------------------------------------------
Private Sub ResetCodingStatus(oItem As EFormItem)
'---------------------------------------------------------------------
' ic 15/08/2005
' sub resets a question coding status to not coded when its value changes
' revisions
' ic 21/11/2006 bug 2843 - pass the current username in to setcodingstatus
'---------------------------------------------------------------------
    If oItem.Element.DataType = eDataType.Thesaurus Then
        If oItem.Response.CodingStatus <> eCodingStatus.csDoNotCode Then
            Call oItem.Response.SetCodingStatus(CInt(eCodingStatus.csNotCoded), goUser.UserName, goUser.UserNameFull)
            Call moFormBuilder.SetClinicalCodingImage(oItem.Response.CodingStatus, imgDictionary(oItem.ControlIndex))
        End If
    End If
End Sub


'---------------------------------------------------------------------
Private Function ViewMIMessage(oItem As EFormItem, enMIMsgType As MIMsgType) As Long
'---------------------------------------------------------------------
'Show MIMessage lists for single question from eForm
'---------------------------------------------------------------------
Dim vData As Variant
Dim bCanViewMIMessage As Boolean
Dim sMsg As String
Dim ofrmMIMModal As frmViewDiscrepanciesModal
Dim sStatus As String

    bCanViewMIMessage = True
    If oItem.EFI.Responses.Changed Then
        sMsg = "Data must be saved before viewing a " & GetMIMTypeText(enMIMsgType) & vbCrLf
        sMsg = sMsg & "Would you like to save this form now?"
        If DialogQuestion(sMsg) = vbYes Then
            bCanViewMIMessage = SaveResponses(eEFormAction.eaSave)
            ' NCJ 10 Jun 02 - CBB 2.2.13/63 Refresh display for all responses
            Call moFormBuilder.DisplayCurrentResponses
        Else
            bCanViewMIMessage = False
        End If
    End If
    
    If bCanViewMIMessage Then

        
        With oItem.Response
        
        Select Case enMIMsgType
        Case MIMsgType.mimtDiscrepancy: sStatus = "111" 'three statuses of discrepancy
        Case MIMsgType.mimtNote: sStatus = "11" 'two statuses of note
        Case MIMsgType.mimtSDVMark: sStatus = "1111" 'four statuses of SDV
        End Select
        ' MLM 01/07/05: Limit MIMessages to the specific visit, eForm and question cycles of oItem
        vData = GetWinIO.GetMIMessageList(enMIMsgType, goUser, .EFormInstance.VisitInstance.Subject.StudyId, .EFormInstance.VisitInstance.Subject.Site, _
                                  .EFormInstance.VisitInstance.VisitId, .EFormInstance.VisitInstance.CycleNo, .EFormInstance.eForm.EFormId, .EFormInstance.CycleNo, oItem.Element.QuestionId, oItem.Response.RepeatNumber, "", "", .EFormInstance.VisitInstance.Subject.PersonId, _
                                    sStatus, "", "", "0001")  '001 denotes questions only
        

        End With
                                                           
            
            
        If Not IsNull(vData) Then
           Set ofrmMIMModal = frmViewDiscrepanciesModal
            Call ofrmMIMModal.DisplayModal(enMIMsgType, SubjectMIMEssage, vData, oItem.Response)
            Set ofrmMIMModal = Nothing
            
            Call moFormBuilder.DisplayCurrentResponse(oItem)
            
        Else
            MsgBox "No matching records"
        End If
    End If


End Function

'---------------------------------------------------------------------
Private Sub ShowAuditTrail(oItem As EFormItem)
'---------------------------------------------------------------------
' Show the audit trail for the given oElement
' NCJ 15/11/01 - Changed to EFormItem
' NCJ 24 Sept 02 - Use oItem.EFI instead of our EFI (e.g. for Visit dates)
'---------------------------------------------------------------------
'reference to AuditTrail form
Dim ofrmAuditTrail As frmWebNonMDI
Dim sglTop As Single
Dim sglLeft As Single
Dim sglHeight As Single

    On Error GoTo ErrLabel
    
    If oItem Is Nothing Then Exit Sub
    
    Call moFormBuilder.GetItemTopLeftHeight(oItem, sglTop, sglLeft, sglHeight)
    
    Set ofrmAuditTrail = New frmWebNonMDI
    
    ' NCJ 9 Nov 04 - Bug 2362 - Don't accept F7 etc. while we're showing the audit trail
    mbDontInterrupt = True
   
    With ofrmAuditTrail
        .Top = GetCoordForForm(oItem).Row
        .Left = GetCoordForForm(oItem).Col
        .Width = 8000
        .Display wdtHTML, AuditTrailHTML(oItem.Response), "auto", True, "Audit Trail"
    End With

    ' NCJ 9 Nov 04 - Reset interruption flag
    mbDontInterrupt = False
    
    Set ofrmAuditTrail = Nothing
    
Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "ShowAuditTrail", Err.Source) = Retry Then
        Resume
    End If

End Sub

'---------------------------------------------------------------------
Private Sub ShowAddComment(oItem As EFormItem)
'---------------------------------------------------------------------
' Allow the user to add to a question's comments
'
' MLM 09/09/02: Pass in an EFormItem instead of an Element and Response
'---------------------------------------------------------------------
Dim sNewComments As String
Dim sDItemName As String
Dim sPrevComments As String
Dim sComment As String
Dim sTimestamp As String
Dim sHeader As String
Dim nExistingLength As String
Dim oResponse As Response

    On Error GoTo ErrLabel
    
    If oItem Is Nothing Then Exit Sub

    Set oResponse = oItem.Response
    
    'TA ensure we have permission, a response, and unlocked
    If oResponse Is Nothing Then Exit Sub
    
    If Not moUser.CheckPermission(gsFnAddIComment) Then Exit Sub
    If oResponse.LockStatus <> LockStatus.lsUnlocked Then Exit Sub
    
    'TA 29/04/2003: set up new comments to be old comments so that we can check for changes
    sNewComments = oResponse.Comments
    'ASH 09/05/2002
    sPrevComments = oResponse.Comments
    sDItemName = oItem.Element.Name

    sComment = InputBox("Please enter your comment:", "Question: " & sDItemName)
    Do While sComment > ""
    '   Check for invalid characters
        If Not gblnValidString(sComment, valOnlySingleQuotes) Then
            ' Reset sNewComments
            sNewComments = sPrevComments
            DialogInformation "Comments cannot contain double or backward quotes or the | character"
        Else
            sComment = Trim(sComment)
            'Changed Mo Morris 14/1/00, format changed from "hh:mm:ss d mm yyyy"
            sTimestamp = Format(Now, "yyyy/mm/dd hh:mm:ss")
            ' The comment header - timestamp and user name
            sHeader = sTimestamp & vbNewLine & "User: " & moUser.UserName & vbNewLine
            sNewComments = sHeader & sComment & vbNewLine
            ' Add in previous comments (if any)
            If Len(sPrevComments) > 0 Then
                sNewComments = sNewComments & vbNewLine & sPrevComments
            End If
        
            'SDM 25/02/00 SR3062 Do not allow the comment field to exceed 255 characters
            If Len(sNewComments) > 255 Then
                MsgBox "The total length of comments cannot exceed 255 characters (including timestamp and user details)", vbOKOnly
                ' Reset sNewComments
                sNewComments = sPrevComments
                ' See how much room the existing comments with the new header will take
                If sPrevComments > "" Then
                    nExistingLength = Len(sPrevComments & vbNewLine & vbNewLine) + Len(sHeader)
                Else
                    ' The header with a NewLine after the comment
                    nExistingLength = Len(sHeader & vbNewLine)
                End If
                'Truncate the new comment so the user can see what is allowed
                If 255 - nExistingLength > 0 Then
                    sComment = Left(sComment, 255 - nExistingLength)
                Else
                    sComment = ""
                End If
            Else
                ' We're OK, so finish here
                Exit Do
            End If
                
        End If  ' If valid string
        
        ' Ask the user to try again with another comment
        sComment = InputBox("Please enter your comment:", "Question: " & sDItemName)
        
    Loop
    
    ' Did the user add anything?
    If sNewComments <> oResponse.Comments Then
        oResponse.Comments = sNewComments
        'call to DisplayCurrentResponse to update comment/note image
        Call moFormBuilder.DisplayCurrentResponse(oItem, True)
        ' User has now done something
        mbUserDidSomething = True
    End If

    Set oResponse = Nothing

Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "ShowAddComment", Err.Source) = Retry Then
        Resume
    End If

End Sub

'---------------------------------------------------------------------
Private Sub NewMIMessage(oItem As EFormItem, enType As MIMsgType)
'---------------------------------------------------------------------
' Create new MIMessage of given type for this CRFElement
' Must make sure data is saved before allowing new message
'
' MLM 09/09/02: Pass in an EFormItem instead of a Response.
'   Only check if the current EFI's responses have changed before prompting for a save.
' MLM 12/09/02: Make sure we add to MIMessage to the correct EFI.
' RS    30/09/2002  Added Timestamp & Timezone Parameter to CreateNewMIMessage call
'---------------------------------------------------------------------
Dim bCanCreateMessage As Boolean
Dim sMsg As String
'MLM 12/09/02:
Dim oEFI As EFormInstance

    On Error GoTo ErrLabel
    
    Set oEFI = oItem.EFI
    
    bCanCreateMessage = True
    If oEFI.Responses.Changed Then
        sMsg = "Data must be saved before creating a new " & GetMIMTypeText(enType) & vbCrLf
        sMsg = sMsg & "Would you like to save this form now?"
        If DialogQuestion(sMsg) = vbYes Then
            bCanCreateMessage = SaveResponses(eEFormAction.eaSave)
            ' NCJ 10 Jun 02 - CBB 2.2.13/63 Refresh display for all responses
            Call moFormBuilder.DisplayCurrentResponses
        Else
            bCanCreateMessage = False
        End If
    End If

    If bCanCreateMessage Then
        With moSubject
            ' Check to see if we already have an SDV for this response
            If enType = MIMsgType.mimtSDVMark And oItem.Response.SDVStatus <> ssNone Then
                DialogInformation "An SDV Mark already exists for this response"
            Else
                ' NCJ 14 Oct 02 - Include Scope, and re-ordered parameters
                Call CreateNewMIMessage(enType, MIMsgScope.mimscQuestion, _
                    IMedNow, moSubject.TimeZone.TimezoneOffset, _
                    .Site, .StudyId, .PersonId, oItem.Response, _
                    oEFI.VisitInstance.VisitId, oEFI.VisitInstance.CycleNo, _
                    oEFI.EFormTaskId, _
                    oItem.Response.ResponseId, oItem.Response.RepeatNumber, _
                    oItem.Response.Value, oItem.Response.TimeStamp, _
                    oEFI.eForm.EFormId, oEFI.CycleNo, oItem.Response.Element.QuestionId, oItem.Response.UserName, _
                    moSubject)
            End If
        End With
    End If

    Call moFormBuilder.DisplayCurrentResponse(oItem)
    
    Set oEFI = Nothing

Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "NewMIMessage", Err.Source) = Retry Then
        Resume
    End If

End Sub

'---------------------------------------------------------------------
Private Function ThereIsInvalidData() As Boolean
'---------------------------------------------------------------------
' Return whether there's unvalidated data on the form
' after allowing txtCRFElement_LostFocus a chance to validate it
' This is to solve problems reported in SR 2260
'---------------------------------------------------------------------

    If mbDirtyField Then
        If Not moFormBuilder.CurrentItem Is Nothing Then
            ' See if it gets validated in a lostfocus event
            Call txtCRFElement_LostFocus(moFormBuilder.CurrentItem.ControlIndex)
        Else
            mbDirtyField = False
        End If
    End If
    ThereIsInvalidData = mbDirtyField
    mbDirtyField = False    ' Ready for next time

End Function

'---------------------------------------------------------------------
Public Sub ResetUnvalidatedData()
'---------------------------------------------------------------------
' NCJ 28 May 03
' Reset unvalidated data on the form if any
' To be called before showing Timeout screen
'---------------------------------------------------------------------

    If mbDirtyField Then
        If Not moFormBuilder.CurrentItem Is Nothing Then
            ' Reset to what it was
            txtCRFElement(moFormBuilder.CurrentItem.ControlIndex).Text = txtCRFElement(moFormBuilder.CurrentItem.ControlIndex).Tag
        End If
    End If
    mbDirtyField = False    ' Ready for next time

End Sub

'---------------------------------------------------------------------
Private Sub ShowQuestionDef(oItem As EFormItem)
'---------------------------------------------------------------------
' Show the validation conditions for this question
'---------------------------------------------------------------------
'reference to Question definition form
Dim ofrmQuDef As frmWebNonMDI
Dim sglTop As Single
Dim sglLeft As Single
Dim sglHeight As Single

    On Error GoTo ErrLabel
    
    If oItem Is Nothing Then Exit Sub
    
    Call moFormBuilder.GetItemTopLeftHeight(oItem, sglTop, sglLeft, sglHeight)
    
    Set ofrmQuDef = New frmWebNonMDI
     
    ' NCJ 9 Nov 04
    mbDontInterrupt = True
    
    With ofrmQuDef
        .Top = GetCoordForForm(oItem).Row
        .Left = GetCoordForForm(oItem).Col
        .Width = 7000
        .Height = 4000
        .Display wdtHTML, QuestionDefHTML(oItem.Element), "auto", True, ""
    End With

    ' NCJ 9 Nov 04
    mbDontInterrupt = False
    
    Set ofrmQuDef = Nothing
    
Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "ShowAuditTrail", Err.Source) = Retry Then
        Resume
    End If

End Sub

'---------------------------------------------------------------------
Private Sub ShowWarning(oItem As EFormItem)
'---------------------------------------------------------------------
' Show the warning for this question
' and allow the user to overrule or un-overrule
' NCJ 13 Aug 02 - Must ask for RFC if appropriate
'
' MLM 10/09/02: Pass in an EFormItem instead of an Element and a Response.
'---------------------------------------------------------------------
Dim lStatus As Long
Dim sOverrule As String
Dim sRFC As String
Dim sRFCCaption As String
Dim bCanChangeOverrule As Boolean
Dim oResponse As Response

    On Error GoTo ErrLabel
    
    Set oResponse = oItem.Response
    
    sRFC = ""
    
    sOverrule = oResponse.OverruleReason
            ' NCJ 15 Aug 02 - New parameters
    lStatus = frmWarnOverrule.Display(oResponse, oResponse.Status, oResponse.Value, _
                                    oResponse.ValidationMessage, sOverrule, GetCoordForForm(oItem))
    If sOverrule <> oResponse.OverruleReason Then
        bCanChangeOverrule = True
        ' User has now done something
        mbUserDidSomething = True

        If oResponse.RequiresOverruleRFC(sOverrule) Then
            ' Ask for Reason for Change
            sRFCCaption = "Please enter the reason for changing the overrule for "
            frmRFC.lblRFC.Caption = sRFCCaption & oItem.Element.Name
            frmRFC.DisplayRFC mcolRFCs, GetCoordForForm(oItem)
            sRFC = frmRFC.GetRFC
            If Len(sRFC) < 1 Then
                bCanChangeOverrule = False
            End If
        End If
        If bCanChangeOverrule Then
            Call oResponse.SetOverruleReason(sOverrule, sRFC)
            ' Redisplay in case status has changed
            Call moFormBuilder.DisplayCurrentResponse(oItem, (oItem.Element.ElementID = moFormBuilder.CurrentElementID))
            ' NCJ 8 Jan 02 - Also update Group status if necessary
            If Not oItem.Element.OwnerQGroup Is Nothing Then
                'MLM 10/09/02: Find out the OwnerQGroup's Item:
                Call moFormBuilder.DisplayCurrentResponse(moFormBuilder.EFormItemByIndex(moFormBuilder.IndexForElementRow(oItem.Element.OwnerQGroup.Element, 1)))
            End If
        End If
    End If
    
    Set oResponse = Nothing
    
Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "ShowWarning", Err.Source) = Retry Then
        Resume
    End If

End Sub

'---------------------------------------------------------------------
Private Function ShowDerivedWarnings() As Boolean
'---------------------------------------------------------------------
' NCJ 31 Oct 02 - Show warnings and rejects resulting from refreshing of derivations
' NB Assume this immediately follows a call to moFormBuilder.UpdateResponses
' Returns TRUE if the user overruled any warnings (so we should redisplay things)
'---------------------------------------------------------------------
Dim oResponse As Response
Dim colWarnResponses As Collection
Dim colRejectResponses As Collection
Dim colMessages As Collection
Dim sOverrule As String
Dim bWarningOverruled As Boolean
Dim oItem As EFormItem  'TODO change this
    On Error GoTo ErrLabel
    
    bWarningOverruled = False
    
    ' Get the collection of warned/rejected derivations
    Call Me.EFI.DerivationWarnRejects(colWarnResponses, colRejectResponses, colMessages)
    For Each oResponse In colWarnResponses
        sOverrule = ""
        ' Show the warning
        Call frmWarnOverrule.Display(oResponse, oResponse.Status, oResponse.Value, colMessages(str(oResponse.ResponseId)), sOverrule, GetCoordForForm(oItem))
        ' Now what if they decide to overrule the warning?
        If sOverrule > "" Then
            ' Change overrule but keep the automatic reason for change
            Call oResponse.SetOverruleReason(sOverrule, oResponse.ReasonForChange)
            bWarningOverruled = True
        End If
    Next
    
    For Each oResponse In colRejectResponses
        ' Show the reject
        Call frmWarnOverrule.Display(oResponse, eStatus.InvalidData, oResponse.Value, colMessages(str(oResponse.ResponseId)), sOverrule, GetCoordForForm(oItem))
    Next
    
    ShowDerivedWarnings = bWarningOverruled
    
    Set oResponse = Nothing
    Set colWarnResponses = Nothing
    Set colRejectResponses = Nothing
    Set colMessages = Nothing

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmEFormDataEntry.ShowDerivedWarnings"

End Function

'---------------------------------------------------------------------
Private Sub ShowInform(oItem As EFormItem)
'---------------------------------------------------------------------
' Show the Inform message for this question
'---------------------------------------------------------------------
Dim sReason As String

    sReason = ""
    ' NCJ 15 Aug 02 - New parameters
    Call frmWarnOverrule.Display(oItem.Response, oItem.Response.Status, _
                                 oItem.Response.Value, oItem.Response.ValidationMessage, sReason, GetCoordForForm(oItem))
    
End Sub

'---------------------------------------------------------------------
Private Sub ChangeLab()
'---------------------------------------------------------------------
' Change the lab and revalidate all lab test questions on the form
' NB We can only get here if all lab questions are unlocked
' Ask for confirmation before changing lab
'---------------------------------------------------------------------
Dim oResponse As Response
Dim sLabCode As String
Dim sMsg As String
Dim nThisElementId As Integer
Dim oItem As EFormItem
Dim bChangedLab As Boolean
Dim bMakeChanges As Boolean

    On Error GoTo ErrLabel
    
    ' Save where we came from
    Set oItem = moFormBuilder.CurrentItem
    
    sLabCode = EFI.LabCode
    If Not frmLaboratories.Display(moSubject.Site, sLabCode) Then
        ' They didn't do anything
        Exit Sub
    End If
    
    ' Did they change lab, or is this their first selection?
    bChangedLab = (EFI.LabCode > "")
    bMakeChanges = True
    
    ' NCJ 30 Jun 03 - Changed "form" to "eForm"
    If bChangedLab Then
        sMsg = "Changing this eForm's laboratory will cause all laboratory results"
        sMsg = sMsg & vbCrLf & "on the eForm to be revalidated."
        sMsg = sMsg & vbCrLf & "Are you sure you wish to change the laboratory?"
        bMakeChanges = (DialogQuestion(sMsg) = vbYes)
    End If
    
    If bMakeChanges Then
        ' User has now done something
        mbUserDidSomething = True

        EFI.LabCode = sLabCode
        
        Call moFormBuilder.ShowLabCode
        
        ' Only revalidate if they changed lab (i.e. not for NEW lab)
        If bChangedLab Then
            ' Go through the lab test questions and revalidate the answers
            ' NCJ 12 Aug 02 - Use new revalidation
            Call RevalidateMe(True)
        End If
    End If
    
    ' Let VB do its focussing
    DoEvents
    ' Go back to where we came from
    If Not oItem Is Nothing Then
        Call moFormBuilder.ResetFormFocus(oItem)
        Set oItem = Nothing
    End If
    
Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "ChangeLab", Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Private Function AskLabCode() As String
'---------------------------------------------------------------------
' Get the lab code, but if it's empty, ask the user to select one
' Check whether current lab is actually defined
' NCJ 30 Jun 04 - Changed so that we consider the site associations too
'---------------------------------------------------------------------
Dim sLabCode As String
Dim sMsg As String
Dim oLabs As clsLabs
Dim oLab As clsLab

    On Error GoTo ErrLabel
    
    ' Have we done all this already?
    If mbLabIsSorted Then
        AskLabCode = EFI.LabCode
        Exit Function
    End If
    
    sLabCode = ""
    
    ' Get the labs for this site
    Set oLabs = New clsLabs
    Call oLabs.Load(moSubject.Site)
    
    If EFI.LabCode > "" And EFI.LabExists Then
        ' See if they're OK with the one they've got
        ' NCJ 30 Jun 04 - Must check the lab is still OK for this site
        ' See if this lab exists
        On Error Resume Next
        Set oLab = oLabs.Item(EFI.LabCode)
        On Error GoTo ErrLabel
        If Not oLab Is Nothing Then
            ' We're OK
            sLabCode = EFI.LabCode
        End If
        Set oLab = Nothing
    End If
    
    ' Tell them if we can't find their lab
    If EFI.LabCode > "" And sLabCode = "" Then
        ' They've got a lab that doesn't exist
        sMsg = "The definition for the laboratory '" & EFI.LabCode & "'"
        sMsg = sMsg & vbCrLf & "used in this eForm is no longer available."
        ' Only let them choose another if there is one!
        If oLabs.Count = 1 Then
            ' Choose the one lab for them
            Set oLab = oLabs.Item(1)
            sLabCode = oLab.Code
            Set oLab = Nothing
            sMsg = sMsg & vbCrLf & "Using new laboratory '" & sLabCode & "'"
        End If
        If oLabs.Count > 1 Then
            sMsg = sMsg & vbCrLf & "Please choose another laboratory."
        End If
        Call DialogInformation(sMsg)
        ' Ensure we remove old lab
        EFI.LabCode = sLabCode
    End If
    
    If sLabCode = "" And oLabs.Count > 0 Then
        ' Get them to choose a lab
        If frmLaboratories.Display(moSubject.Site, sLabCode) Then
            ' Set it for this form
            EFI.LabCode = sLabCode
        End If
    End If
    
    ' NCJ 10/10/01 - Must display lab code
    Call moFormBuilder.ShowLabCode
    
    Set oLabs = Nothing
    
    ' Return value
    AskLabCode = sLabCode
    
    ' Remember we've been through this already
    mbLabIsSorted = True

Exit Function
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Display", Err.Source) = Retry Then
        Resume
    End If

End Function

'---------------------------------------------------------------------
Private Sub vsbQuGroup_Change(Index As Integer)
'---------------------------------------------------------------------
' A change to a question group's vertical scroll bar
'---------------------------------------------------------------------

    ' NCJ 4 Nov 02 - Hide popup and calendar before scrolling
    Call HidePopupAndCalendar
    Call moFormBuilder.QGroupVScroll(Index, vsbQuGroup(Index).Value)

End Sub

'---------------------------------------------------------------------
Private Sub vsbQuGroup_Scroll(Index As Integer)
'---------------------------------------------------------------------
' A change to a question group's vertical scroll bar
'---------------------------------------------------------------------

    ' NCJ 4 Nov 02 - Hide popup and calendar before scrolling
    Call HidePopupAndCalendar
    Call moFormBuilder.QGroupVScroll(Index, vsbQuGroup(Index).Value)

End Sub

'---------------------------------------------------------------------
Private Sub ScrollIfOffScreen(oControl As Control)
'---------------------------------------------------------------------
' Scroll the eForm if control with focus is not visible
' NCJ 1 Feb 02 - Make sure we pick up the container for Question Group items
' RS 05/02/03 - Adjust height of active dropdown item (if any)
' NCJ 19 Feb 03 - Sorted out scrolling properly
' NCJ 20 Feb 03 - Farmed out code to routines CheckVerticalScroll and CheckHorizontalScroll
'---------------------------------------------------------------------
Dim oActualControl As Control

    'ignore all errors
    On Error GoTo ErrLabel
    
    ' NCJ 1 Feb 02 - Get container if in a question group
    Set oActualControl = oControl
    ' Is it in a question group?
    If oControl.Container.Name = "picScroller" Then
        ' picScroller sits inside picFixed inside picQuGroup
        Set oActualControl = oControl.Container.Container.Container
    End If
    
    Call CheckVerticalScroll(oActualControl.Top, oActualControl.Height)
    
    Call CheckHorizontalScroll(oActualControl.Left, oActualControl.Width)
    
    ' --------------------------------------
    ' RS 05/02/03: Resize Dropdown if needed
    ' This code is copied from method CreatePopupList in EFormBuilder
    ' --------------------------------------
    
    If lvwPopupList.Visible And lvwPopupList.Top > 0 Then
        ' NCJ 22 May 03 - Make sure we call the same code as in EFormBuilder
        Call moFormBuilder.AdjustPopupListHeight
    End If
    
''        If lvwPopupList.ListItems.Count > 0 Then
''            lvwPopupList.Height = (lvwPopupList.ListItems.Count + 1.6) * lvwPopupList.ListItems(1).Height
''        Else
''            lvwPopupList.Height = Me.TextHeight("W") * 2
''        End If
''
''        If lvwPopupList.Height > Me.fraCRFPage.Height - lvwPopupList.Top Then
''            ' RS 11/02/2003: Only resize if dropdown no longer fits due to vertical page scroll
''            If lvwPopupList.Height > Max(Me.fraCRFPage.Height - lvwPopupList.Top - 300 - Me.picCRFPage.Top, 0) Then
''                lvwPopupList.Height = Max(Me.fraCRFPage.Height - lvwPopupList.Top - 300 - Me.picCRFPage.Top, 0)
''            End If
''            'the above reduction of the listviews height will cause a vertical scrollbar to be displayed
''            'if the listviews width is not increased to allow for this a horizontal scrollbar will also appear
''            'check to see if there is enough space to increase the width
''            If lvwPopupList.Left + lvwPopupList.Width + 240 < Me.picCRFPage.ScaleWidth Then
''                lvwPopupList.Width = lvwPopupList.Width + 240
''            End If
''        End If
''    End If

Exit Sub
ErrLabel:
    ' Ignore errors
    Debug.Print "Error in ScrollIfOffScreen: " & Err.Description
End Sub

'---------------------------------------------------------------------
Private Sub CheckVerticalScroll(ByVal sglTop As Single, ByVal sglHeight As Single)
'---------------------------------------------------------------------
' NCJ 20 Feb 03 (code copied from ScrollIfOffScreen)
' Adjust the vertical scroll as needed so that the coord sglTop is visible,
' and try and make sure the height is fully visible
'---------------------------------------------------------------------
Dim lBorder As Long
Dim nScrollAmount As Integer

    On Error GoTo ErrLabel      ' We'll ignore errors
    
    If vsbCRFPage.Visible Then
        'TA 10/03/2003: only try this when there is a vertical scrollbar
    
        'set "border" to a 1/5th of the viewable area
        lBorder = Me.fraCRFPage.Height / 5
        
        ' See if the item will fit in our "border" area
        If sglHeight > lBorder Then
            ' Check it's not bigger than the viewable area
            If sglHeight >= Me.fraCRFPage.Height Then
                ' Take up 4/5 of the viewable area
                lBorder = lBorder * 4
            Else
                ' Use the item's height
                lBorder = sglHeight
            End If
        End If
        
        'if control moves into the bottom 1/5th of the visible section of a form or beyond,
        'the screen will be made to scroll
        If sglTop + lBorder + Me.picCRFPage.Top > Me.fraCRFPage.Height Then
            'set scroll bar position to that of the element in question + Border
            nScrollAmount = (sglTop + lBorder - Me.fraCRFPage.Height) / 10
            ' Check for going off the end
            If nScrollAmount < Me.vsbCRFPage.Max Then
                Me.vsbCRFPage.Value = nScrollAmount
            Else
                Me.vsbCRFPage.Value = Me.vsbCRFPage.Max
            End If
        'otherwise check the control has not moved to a non-visible section at the top of the form
        ElseIf sglTop < -Me.picCRFPage.Top Then
            'Check to see if the element in question is in the top section of the form
            If sglTop + lBorder < Me.fraCRFPage.Height Then
                'scroll to the top of the form by setting value to min (i.e. zero)
                Me.vsbCRFPage.Value = Me.vsbCRFPage.Min
            Else
                'set scroll bar position to that of the element in question + Border)
                Me.vsbCRFPage.Value = (sglTop - lBorder) / 10
            End If
        End If
    End If

Exit Sub
ErrLabel:
    ' Ignore errors

End Sub

'---------------------------------------------------------------------
Private Sub CheckHorizontalScroll(ByVal sglLeft As Single, ByVal sglWidth As Single)
'---------------------------------------------------------------------
' NCJ 20 Feb 03
' Adjust the horizontal scroll as needed so that the coord sglLeft is visible
' and try and make sure the width is fully visible
'---------------------------------------------------------------------
Dim lBorder As Long
Dim nScrollAmount As Integer

    On Error GoTo ErrLabel      ' We'll ignore errors
    
    If hsbCRFPage.Visible Then
        'TA 10/03/2003: only try this when there is a horiz scrollbar
        
        'set "border" to a 1/5th of the viewable area
        lBorder = Me.fraCRFPage.Width / 3
        
        ' See if the item will fit in our "border" area
        If sglWidth > lBorder Then
            ' Check it's not bigger than the viewable area
            If sglWidth >= Me.fraCRFPage.Width Then
                ' Take up 5/6 of the viewable area
                lBorder = lBorder * 2.5
            Else
                ' Use the item's width
                lBorder = sglWidth
            End If
        End If
        
        'if control moves into the rightmost 1/5th of the visible section of a form or beyond,
        'the screen will be made to scroll
        If sglLeft + lBorder + Me.picCRFPage.Left > Me.fraCRFPage.Width Then
            'set scroll bar position to that of the element in question + 1/5th (lBorder)
            nScrollAmount = (sglLeft + lBorder - Me.fraCRFPage.Width)
            ' Check for going off the end
            If nScrollAmount < Me.hsbCRFPage.Max Then
                Me.hsbCRFPage.Value = nScrollAmount
            Else
                Me.hsbCRFPage.Value = Me.hsbCRFPage.Max
            End If
        'otherwise check the control has not moved to a non-visible section at the left of the form
        ElseIf sglLeft < -Me.picCRFPage.Left Then
            'Check to see if the element in question is in the left section of the form
            If sglLeft + lBorder < Me.fraCRFPage.Width Then
                'scroll to the left of the form by setting value to min (i.e. zero)
                Me.hsbCRFPage.Value = Me.hsbCRFPage.Min
            Else
                'set scroll bar position to that of the element in question + 1/5th (lBorder)
                Me.hsbCRFPage.Value = (sglLeft - lBorder)
            End If
        End If
        
    End If

Exit Sub
ErrLabel:
    ' Ignore errors

End Sub

'---------------------------------------------------------------------
Private Function FocusToAppropriateField(oItem As EFormItem, _
                                        ByVal bStayPut As Boolean)
'---------------------------------------------------------------------
' NCJ 8 Jul 02 - Focus to next appropriate field where oItem is current field
' and bStayPut is whether we'd like to stay on this field
' (Reimplemented from 2.2)
'---------------------------------------------------------------------

    If bStayPut Then
        ' Try to stay where we are
        Call moFormBuilder.ResetFormFocus(oItem)
    Else
        ' Focus to "next" field
        Call moFormBuilder.FocusToNextField(oItem.Element, oItem.RowNo)
    End If
    
End Function

'---------------------------------------------------------------------
Private Function RevalidateMe(ByVal bLabChange As Boolean) As Boolean
'---------------------------------------------------------------------
' Revalidate all nonempty fields on the eForm
' checking lab test questions only if bLabChange = TRUE,
' or checking non-labtest questions only if bLabChange = FALSE
' Returns TRUE if there were no new warnings or rejects,
' or FALSE is there was a new warning or reject
'
' MLM 16/09/02: Revalidate the Visit EForm, if there is one.
'---------------------------------------------------------------------
Dim sMsg As String
Dim bNewWarnReject As Boolean
Dim bSomethingChanged As Boolean
Dim oRejectedResponse As Response
Dim enContext As eAutoRFCContext
Dim oVisitEFI As EFormInstance

    On Error GoTo ErrLabel
    
    ' Only revalidate revalidatable eForms
    If Not moFormBuilder.CanUpdateEForm Then
        RevalidateMe = True
        Exit Function
    End If
    
    ' Remember if we get any new warnings or rejections
    bNewWarnReject = False
    ' Remember if anything changed
    bSomethingChanged = False
    Set oRejectedResponse = Nothing
    
    'MLM 16/09/02: Revalidate the Visit EForm, if there is one.
    Set oVisitEFI = EFI.VisitInstance.VisitEFormInstance
    If Not oVisitEFI Is Nothing And Not bLabChange Then
        bNewWarnReject = RevalidateEFI(oVisitEFI, False, oRejectedResponse, bSomethingChanged)
    End If
    
    'revalidate the main eform
    bNewWarnReject = bNewWarnReject Or RevalidateEFI(EFI, bLabChange, oRejectedResponse, bSomethingChanged)

    If Not oRejectedResponse Is Nothing Then
        ' Something was rejected so we re-run skips and derivations
        ' NCJ 14 Aug 02 - Set the context for auto RFC generation
        If bLabChange Then
            enContext = ChangingLab
        Else
            enContext = Revalidation
        End If
        ' NCJ 8 July 04 - Call non-refreshing UpdateSkipsDerivs
        'Call moFormBuilder.UpdateSkipsDerivs(enContext, oRejectedResponse)
        ' NCJ 13 Jan 05 - Put it back to what it was (to prevent bug 2495)
         Call moFormBuilder.UpdateResponses(enContext, oRejectedResponse)
    End If
    
    Set oRejectedResponse = Nothing
    
    If bSomethingChanged Then
        ' Something changed so update statuses
    ' NCJ 8 July 04 - Update everything anyway
    ' NCJ 13 Jan 05 - Reset to conditional updating!
        Call moFormBuilder.DisplayCurrentResponses
    End If
    
    ' Return TRUE if NO new warnings/rejections
    RevalidateMe = Not bNewWarnReject
    
Exit Function
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "RevalidateMe", Err.Source) = Retry Then
        Resume
    End If

End Function

'---------------------------------------------------------------------
Private Function RevalidateEFI(ByRef oEFI As EFormInstance, ByVal bLabChange As Boolean, _
    ByRef oRejectedResponse As Response, ByRef bSomethingChanged As Boolean) As Boolean
'---------------------------------------------------------------------
' MLM 16/09/02: Added based on part of RevalidateMe.
' Returns TRUE if revalidating the efi generated rejects or new warnings.
' oRejectedResponse and bSomethingChanged may also be set.
' NCJ 24 Mar 03 - Refresh Subject properties first (in case DOB or Gender used in lab test questions)
'---------------------------------------------------------------------
Dim oResponse As Response
Dim bChangedStatus As Boolean
Dim sErrMsg As String
Dim sOverrule As String
Dim lStatus As Long

Dim oItem As EFormItem 'need to change this

    On Error GoTo ErrLabel
    
    RevalidateEFI = False

    ' NCJ 24 Mar 03 - Refresh DOB, Gender etc.
    Call oEFI.VisitInstance.Subject.RefreshProperties
    
    For Each oResponse In oEFI.Responses
        ' Only deal with "Revalidatable" responses
        If oResponse.Revalidatable(bLabChange) Then
        
            ' NCJ 5 Jul 04 - Must check the lab code here if we're revalidating
            If oResponse.Value > "" And oResponse.Element.DataType = eDataType.LabTest Then
                Call AskLabCode
            End If
            
            lStatus = oResponse.RevalidateValue(sErrMsg, bChangedStatus)
            If bChangedStatus Then
                bSomethingChanged = True
                sOverrule = ""
                'MLM 13/09/02: Make sure messages are only displayed for enterable questions:
                If (lStatus = eStatus.Warning Or lStatus = eStatus.InvalidData) And oResponse.Enterable Then
                    ' Show the Warning/Error message
                    Call frmWarnOverrule.Display(oResponse, lStatus, oResponse.Value, sErrMsg, sOverrule, GetCoordForForm(oItem))
                    RevalidateEFI = True
                    If lStatus = eStatus.InvalidData And oRejectedResponse Is Nothing Then
                        ' Remember the first rejected response
                        Set oRejectedResponse = oResponse
                    End If
                End If
                ' Confirm the change
                Call oResponse.ConfirmRevalidation(sOverrule)
            Else
                ' No change so ignore the revalidation
                Call oResponse.RejectValue
            End If      ' if changed
            
        End If      ' if we revalidated
    Next    ' oResponse
    
    Set oResponse = Nothing

Exit Function

ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "RevalidateEFI", Err.Source) = Retry Then
        Resume
    End If

End Function

'---------------------------------------------------------------------
Private Function EFormAndVisitDatesBlank(bShowUserMessage As Boolean, enEFormAction As eEFormAction, _
                ByRef oElement As eFormElementRO) As Boolean
'---------------------------------------------------------------------
' MLM 30/08/02: Added. Returns True if the eForm has an eForm or visit date
'   question that is unanswered, which should prevent the eForm from being saved.
' MLM 10/09/02: Added check for visit date. Also, only force the user to enter
'   a date if they're actually able to.
' NCJ 6 Jan 03 - Only show user messages if bShowUserMessage = TRUE
' MLM 13/02/03: If return value is True (something is missing) then return value of new parameter
'               oElement should be inspected. If it's Nothing, the current form should be
'               cancelled. Otherwise, it indicates the missing value, that should be focussed to.
'---------------------------------------------------------------------
Dim oVisitEFI As EFormInstance
Dim oResponse As Response
Dim bBlank As Boolean
Dim sMsg As String
    
    On Error GoTo ErrLabel

    bBlank = False 'assume we're ok to begin with
    
    'check visit date first
    Set oVisitEFI = EFI.VisitInstance.VisitEFormInstance
    If Not oVisitEFI Is Nothing Then
        ' Ignore if read-only
        If (Not oVisitEFI.ReadOnly) And (Not oVisitEFI.eForm.EFormDateElement Is Nothing) Then
            Set oElement = oVisitEFI.eForm.EFormDateElement
            Set oResponse = oVisitEFI.Responses.ResponseByElement(oElement)
            bBlank = (oResponse.Value = "") And oResponse.Enterable
            sMsg = "visit"
'            ' NCJ 6 Jan 03 - Only show message if bShowUserMessage is TRUE
'            If bBlank And bShowUserMessage Then
'                DialogError "The eForm cannot be saved because no visit date has been specified."
'            End If
        End If
    End If
        
    'check eForm date
    'if the visit date isn't missing, and the form has a form date, and form isn't ReadOnly
    If Not bBlank And (Not EFI.ReadOnly) And (Not moEForm.EFormDateElement Is Nothing) Then
        Set oElement = moEForm.EFormDateElement
        Set oResponse = EFI.Responses.ResponseByElement(oElement)
        bBlank = (oResponse.Value = "") And oResponse.Enterable
        sMsg = "eForm"
        ' NCJ 6 Jan 03 - Only show message if bShowUserMessage is TRUE
'        If bBlank And bShowUserMessage Then
'            DialogError "The eForm cannot be saved because no eForm date has been specified."
'        End If
    End If
    
    If bBlank Then
        'either the form or visit date is missing; decide what to do about it
        If bShowUserMessage Then
            sMsg = "This eForm cannot be saved because no " & sMsg & " date has been specified."
            If enEFormAction = eaSave Then
                'inform the user that the save will not be performed
                DialogError sMsg
            ElseIf DialogQuestion(sMsg & vbCrLf & "Do you want to cancel the changes made to this eForm?", , , vbDefaultButton2) = vbYes Then
                Set oElement = Nothing
            End If
        Else
            Set oElement = Nothing
        End If
    End If
    
    EFormAndVisitDatesBlank = bBlank
    
    Set oVisitEFI = Nothing
    Set oResponse = Nothing

Exit Function

ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "EFormAndVisitDatesBlank", Err.Source) = Retry Then
        Resume
    End If

End Function

'---------------------------------------------------------------------
Public Function SaveNeeded(Optional ByRef bResponsesExist As Boolean = False) As Boolean
'---------------------------------------------------------------------
' MLM 10/09/02: Added. Returns True if the form needs to be saved.
'   If present, bResponsesExist will be set to whether or not there are responses
'   (needed so that blank forms will save without prompting the user)
' ic 09/09/2005 added clinical coding
'---------------------------------------------------------------------

Dim bSaveNeeded As Boolean
Dim oVisitEFI As EFormInstance
Dim oEFI As EFormInstance

    Set oEFI = Me.EFI
    bResponsesExist = (oEFI.Responses.Count > 0)
    bSaveNeeded = oEFI.Responses.Changed Or Not bResponsesExist
    Set oVisitEFI = oEFI.VisitInstance.VisitEFormInstance
    If Not (bSaveNeeded Or oVisitEFI Is Nothing) Then
        'also consider the Visit EForm
        bResponsesExist = bResponsesExist Or (oVisitEFI.Responses.Count > 0)
        bSaveNeeded = oVisitEFI.Responses.Changed Or Not bResponsesExist
    End If
    
    SaveNeeded = bSaveNeeded
    
    Set oEFI = Nothing
    Set oVisitEFI = Nothing

End Function

'---------------------------------------------------------------------
Public Function SaveCodedTermHistoriesNeeded()
'---------------------------------------------------------------------
' ic 13/09/2005
' function returns true if the coded values need saving
'---------------------------------------------------------------------
    'check for changes to coded responses
    SaveCodedTermHistoriesNeeded = Me.EFI.SaveCodedTermHistoriesNeeded()
End Function

'--------------------------------------------------------------------------
Public Property Get EFI() As EFormInstance
'--------------------------------------------------------------------------
' NCJ 17 Sept 02
' The eFormInstance for this eForm,
' retrieved by EFormTaskId
'--------------------------------------------------------------------------

    If mlEFormTaskId > 0 Then
        Set EFI = moSubject.eFIByTaskId(mlEFormTaskId)
    End If

End Property

'--------------------------------------------------------------------------
Private Function GetCoordForForm(oItem As EFormItem) As clsCoords
'--------------------------------------------------------------------------
' work out position of control in relation to the screen object
'--------------------------------------------------------------------------
Dim sglTop As Single
Dim sglLeft As Single
Dim sglHeight As Single
Dim sglEFormTop As Single

    Set GetCoordForForm = New clsCoords
    If Not oItem Is Nothing Then
        Call moFormBuilder.GetItemTopLeftHeight(oItem, sglTop, sglLeft, sglHeight)
    Else
        'revalidation
    End If
    
    'work out eform top - must take into account title bat and status bar if shown
    sglEFormTop = frmMenu.Top + (frmMenu.Height - frmMenu.ScaleHeight) + Me.Top
    If frmMenu.picSymbols(0).Visible Then
        sglEFormTop = sglEFormTop - 2 * frmMenu.picSymbols(0).Height
    End If
    If frmMenu.picFunctions.Visible Then
        sglEFormTop = sglEFormTop - frmMenu.picFunctions.Height
    End If
    
    GetCoordForForm.Row = sglEFormTop + fraCRFPage.Top + picCRFPage.Top + sglTop ' + sglHeight + 120
    
    'make sure it always on screen
    If GetCoordForForm.Row < 400 Then
        GetCoordForForm.Row = 400
    End If

    If GetCoordForForm.Row > Screen.Height - 2000 Then
        GetCoordForForm.Row = Screen.Height - 2000
    End If
    
    
    GetCoordForForm.Col = frmMenu.Left + 240 '+ Me.Left + picCRFPage.Left

End Function
        
'--------------------------------------------------------------------------
Public Sub RefreshResponses()
'--------------------------------------------------------------------------
' Updates all the display of all the responses - callable from outside the eForm
'--------------------------------------------------------------------------

    Call moFormBuilder.DisplayCurrentResponses

End Sub

'--------------------------------------------------------------------------
Private Sub ToggleUserIntervention(bEnable As Boolean)
'--------------------------------------------------------------------------
' Enable or disable user intervention while we're busy
'--------------------------------------------------------------------------

    ' Don't let them click anywhere and mess things up
    Call frmMenu.ToggleUserIntervention(bEnable)
    ' Set the "don't interrupt" flag if we're disabling user intervention
    mbDontInterrupt = Not bEnable

End Sub


' **** RANDOMISATION CODE CONDITIONALLY COMPILED IF RANDOM = 1 ***************************
' **** NCJ 2nd Oct 2003 ********************************************************************
' **** NCJ 7 Jan 04 - Removed conditional compilation

'----------------------------------------------------------------------------------------'
Private Function RandomiseSubject(ByVal sDBConnection As String, _
                oSubject As StudySubject, _
                ByRef colResults As Collection) As Boolean
'----------------------------------------------------------------------------------------'
' Randomise a subject and return results in colResults
'----------------------------------------------------------------------------------------'
Dim oDBCon As ADODB.Connection
Dim sSQL As String
Dim rsRands As ADODB.Recordset
Dim sRandCode As String
Dim sStratValue As String
Dim sTreatment As String
Dim bHaveRandomised As Boolean

    On Error GoTo ErrHandler
    
    ' Have we randomised?
    bHaveRandomised = False
    Set colResults = New Collection
    
    Set oDBCon = New ADODB.Connection
    Call oDBCon.Open(sDBConnection)
    oDBCon.CursorLocation = adUseClient
    
    ' Check we have the necessary randomisation tables
    If RandomisationAvailable(oDBCon) Then
        ' Get the possible randomisation codes
        sSQL = "SELECT RandomisationCode, RandomisationCond, StratificationExpr" _
            & " FROM Randomisation WHERE " _
            & " ClinicalTrialId = " & oSubject.StudyId
        Set rsRands = New ADODB.Recordset
        rsRands.Open sSQL, oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    
        ' See if any need processing
        If rsRands.RecordCount > 0 Then
            rsRands.MoveFirst
            Do While Not rsRands.EOF
                sRandCode = rsRands!RandomisationCode
                If Not IsSubjectRandomised(oSubject, oDBCon, sRandCode) Then
                    ' See if we need to process this randomisation
                    If ShouldRandomiseSubject(oSubject, rsRands!RandomisationCond) Then
                        ' We're going to attempt a randomisation
                        bHaveRandomised = True
                        sStratValue = EvaluateStratification(oSubject, rsRands!StratificationExpr)
                        If sStratValue > "" Then
                            ' Assign treatment
                            Call AssignTreatment(oSubject, oDBCon, colResults, sRandCode, sStratValue)
                        Else
                            ' Add a "Can't stratify" error for this randomisation
                            Call IncludeResult(colResults, rrCantStratify, sRandCode, "")
                        End If
                    End If
                End If
                rsRands.MoveNext
            Loop
        End If
        rsRands.Close
        Set rsRands = Nothing
    End If
    
    Set oDBCon = Nothing
    
    RandomiseSubject = bHaveRandomised
    
Exit Function
ErrHandler:
    ' Set the results to show an error
    Call IncludeResult(colResults, rrError, "", Err.Description)
    ' Signal that there's something to look at
    RandomiseSubject = True
    
End Function

'----------------------------------------------------------------------------------------'
Private Function IsSubjectRandomised(oSubject As StudySubject, _
                                    oDBCon As ADODB.Connection, _
                                    ByVal sRandCode As String) As Boolean
'----------------------------------------------------------------------------------------'
' Returns TRUE if subject is already randomised on this randomisation
'----------------------------------------------------------------------------------------'
Dim rsRands As ADODB.Recordset
Dim sSQL As String

    IsSubjectRandomised = False
    
    On Error GoTo ErrHandler
    
    ' See if this subject has been allocated a treatment on this randomisation
    sSQL = "SELECT Treatment FROM Treatments WHERE " _
        & " ClinicalTrialId = " & oSubject.StudyId _
        & " AND RandomisationCode = '" & sRandCode & "'" _
        & " AND TrialSite = '" & oSubject.Site & "'" _
        & " AND PersonId = " & oSubject.PersonId
    Set rsRands = New ADODB.Recordset
    rsRands.Open sSQL, oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    If rsRands.RecordCount > 0 Then
        ' They've been done
        IsSubjectRandomised = True
    End If
    rsRands.Close
    
    Set rsRands = Nothing

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "frmEFormDataEntry.IsSubjectRandomised"

End Function

'----------------------------------------------------------------------------------------'
Private Function ShouldRandomiseSubject(oSubject As StudySubject, _
                                        ByVal sRandCond As String) As Boolean
'----------------------------------------------------------------------------------------'
' Returns TRUE if the randomisation condition is true
'----------------------------------------------------------------------------------------'
Dim sResult As String

    sResult = oSubject.Arezzo.EvaluateExpression(sRandCond)
    ShouldRandomiseSubject = (sResult = "true")

End Function

'----------------------------------------------------------------------------------------'
Private Function TreatmentDataItem(ByVal sRandCode As String) As String
'----------------------------------------------------------------------------------------'
' Returns the AREZZO data item name of the treatment for this randomisation
'----------------------------------------------------------------------------------------'

    TreatmentDataItem = "person:" & LCase(sRandCode) & ":treatment"
    
End Function

'----------------------------------------------------------------------------------------'
Private Function EvaluateStratification(oSubject As StudySubject, _
                                    ByVal sStratExpr As String) As String
'----------------------------------------------------------------------------------------'
' Evaluate stratification expression
' Returns empty string if no go
'----------------------------------------------------------------------------------------'
Dim sStrat As String

    sStrat = oSubject.Arezzo.EvaluateExpression(sStratExpr)
    If Not oSubject.Arezzo.ResultOK(sStrat) Then
        sStrat = ""
    End If
    EvaluateStratification = sStrat

End Function

'----------------------------------------------------------------------------------------'
Private Sub IncludeResult(colResults As Collection, _
                            ByVal enResultType As RandomisationResult, _
                            ByVal sRandCode As String, _
                            ByVal sTreatment As String)
'----------------------------------------------------------------------------------------'
' Add a result to the results collection
'----------------------------------------------------------------------------------------'
Dim sResult As String
Const sSep = "|"

    sResult = enResultType & sSep & sRandCode & sSep & sTreatment
    colResults.Add sResult

End Sub

'----------------------------------------------------------------------------------------'
Private Sub AssignTreatment(oSubject As StudySubject, _
                                oDBCon As ADODB.Connection, _
                                colResults As Collection, _
                                ByVal sRandCode As String, _
                                ByVal sStratValue As String)
'----------------------------------------------------------------------------------------'
' Get the treatment for this subject
' and add an entry to the colResults collection accordingly
' NCJ 6 Oct 03 - Add AREZZO data first, before updating Treatments table
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim oTimezone As TimeZone
Dim rsTreats As ADODB.Recordset
Dim sTreat As String

    On Error GoTo ErrHandler
    
    sTreat = ""
    ' Get the next unassigned treatment
    sSQL = "SELECT * FROM Treatments WHERE" _
        & " ClinicalTrialId = " & oSubject.StudyId _
        & " AND RandomisationCode = '" & sRandCode & "'" _
        & " AND StratificationValue = '" & sStratValue & "'" _
        & " AND TrialSite IS NULL" _
        & " ORDER BY TreatSeqNo "
    
    Set rsTreats = New ADODB.Recordset
    rsTreats.Open sSQL, oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    
    ' Hopefully there is a record
    If rsTreats.RecordCount > 0 Then
        rsTreats.MoveFirst
        ' Quickly mark this one as ours!
        rsTreats!TrialSite = oSubject.Site
        Call rsTreats.Update
        ' Haul out the treatment
        sTreat = rsTreats!Treatment
        ' Add the data to AREZZO first to make sure we can do it
        Call oSubject.Arezzo.AddData(TreatmentDataItem(sRandCode), sTreat)
        If oSubject.Save = eSaveResponsesResult.srrSuccess Then
            ' Fill in the rest of the subject details
            rsTreats!PersonId = oSubject.PersonId
            If oSubject.label > "" Then
                rsTreats!SubjectLabel = oSubject.label
            End If
            rsTreats!RandomiseDate = IMedNow
            Set oTimezone = New TimeZone
            rsTreats!RandomiseDate_TZ = oTimezone.TimezoneOffset
            Set oTimezone = Nothing
            Call rsTreats.Update
            ' Add this successful result to the collection
            Call IncludeResult(colResults, rrSuccess, sRandCode, sTreat)
        Else
            ' Can't update patient state so give up
            ' Release this treatment
            rsTreats!TrialSite = Null
            Call rsTreats.Update
            Call IncludeResult(colResults, rrError, sRandCode, sTreat)
        End If
    Else
        ' If the recordset is empty, we either have an unknown stratification value
        ' or we've run out of treatments for this value
        Call IncludeResult(colResults, rrNoTreatments, sRandCode, "")
    End If

    rsTreats.Close
    Set rsTreats = Nothing

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "frmEFormDataEntry.AssignTreatment"
    
End Sub

'----------------------------------------------------------------------------------------'
Private Function RandomisationAvailable(oDBCon As ADODB.Connection) As Boolean
'----------------------------------------------------------------------------------------'
' Are the randomisation tables available in this database?
'----------------------------------------------------------------------------------------'
Dim rsRands As ADODB.Recordset

    RandomisationAvailable = True
    
    On Error GoTo NoTable
    
    Set rsRands = New ADODB.Recordset
    ' Just try and access the table with a dummy select
    rsRands.Open "SELECT RandomisationCode FROM Randomisation WHERE 1=2", oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    rsRands.Close
    rsRands.Open "SELECT RandomisationCode FROM Treatments WHERE 1=2", oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    rsRands.Close
    Set rsRands = Nothing

Exit Function
NoTable:
    ' Error reading from tables so return FALSE
    Set rsRands = Nothing
    RandomisationAvailable = False
End Function

' ******* END OF RANDOMISATION CODE *************************************************************

