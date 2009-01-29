VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStudyDefinition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Study Definition"
   ClientHeight    =   7860
   ClientLeft      =   5220
   ClientTop       =   3555
   ClientWidth     =   9420
   Icon            =   "frmStudyDefinition.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7860
   ScaleWidth      =   9420
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6780
      TabIndex        =   20
      Top             =   7380
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   8115
      TabIndex        =   21
      Top             =   7380
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   12726
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Study Details"
      TabPicture(0)   =   "frmStudyDefinition.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraStudyDetails"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Subject Details"
      TabPicture(1)   =   "frmStudyDefinition.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Registration"
      TabPicture(2)   =   "frmStudyDefinition.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Reasons for Change"
      TabPicture(3)   =   "frmStudyDefinition.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Required Questions"
      TabPicture(4)   =   "frmStudyDefinition.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame2"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Reasons for Overrule"
      TabPicture(5)   =   "frmStudyDefinition.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame16"
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame16 
         Height          =   6675
         Left            =   -74880
         TabIndex        =   92
         Top             =   420
         Width           =   9015
         Begin VB.TextBox txtOverruleEdit 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   360
            TabIndex        =   93
            Text            =   "Text control used to edit grid"
            Top             =   1140
            Visible         =   0   'False
            Width           =   2175
         End
         Begin MSFlexGridLib.MSFlexGrid gridOverrule 
            Height          =   6375
            Left            =   120
            TabIndex        =   94
            Top             =   180
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   11245
            _Version        =   393216
            Cols            =   3
            FocusRect       =   2
         End
      End
      Begin VB.Frame Frame3 
         Height          =   6675
         Left            =   -74880
         TabIndex        =   45
         Top             =   360
         Width           =   8955
         Begin VB.Frame Frame9 
            Caption         =   "Subject Registration occurs after"
            Height          =   675
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   7035
            Begin VB.ComboBox cboVisit 
               Height          =   315
               Left            =   600
               Style           =   2  'Dropdown List
               TabIndex        =   52
               Top             =   240
               Width           =   2700
            End
            Begin VB.ComboBox cboeForm 
               Height          =   315
               Left            =   4200
               Style           =   2  'Dropdown List
               TabIndex        =   53
               Top             =   240
               Width           =   2700
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               Caption         =   "Visit"
               Height          =   255
               Left            =   120
               TabIndex        =   67
               Top             =   300
               Width           =   375
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Caption         =   "eForm"
               Height          =   255
               Left            =   3660
               TabIndex        =   66
               Top             =   300
               Width           =   435
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Subject Identifier Prefix"
            Height          =   1035
            Left            =   120
            TabIndex        =   63
            Top             =   1680
            Width           =   8715
            Begin VB.TextBox txtRegExpression 
               Height          =   315
               Index           =   0
               Left            =   1200
               TabIndex        =   57
               Top             =   240
               Width           =   6675
            End
            Begin VB.CheckBox chkBaseNumbering 
               Caption         =   "Restart numbering for new prefixes"
               Height          =   375
               Index           =   0
               Left            =   180
               TabIndex        =   59
               Top             =   600
               Width           =   3315
            End
            Begin VB.CommandButton cmdRegExpression 
               BackColor       =   &H00FFFFFF&
               Height          =   600
               Index           =   0
               Left            =   7980
               Picture         =   "frmStudyDefinition.frx":04EA
               Style           =   1  'Graphical
               TabIndex        =   58
               Top             =   180
               Width           =   600
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               Caption         =   "Expression"
               Height          =   255
               Left            =   120
               TabIndex        =   64
               Top             =   300
               Width           =   855
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Server"
            Height          =   675
            Left            =   7260
            TabIndex        =   51
            Top             =   240
            Width           =   1575
            Begin VB.CommandButton cmdRRServerSettings 
               Caption         =   "Settings..."
               Height          =   375
               Left            =   180
               TabIndex        =   54
               Top             =   210
               Width           =   1215
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Subject Identifier Numbering"
            Height          =   675
            Left            =   120
            TabIndex        =   48
            Top             =   960
            Width           =   8715
            Begin VB.TextBox txtStartNo 
               Height          =   315
               Left            =   1260
               MaxLength       =   10
               TabIndex        =   55
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txtNumberWidth 
               Height          =   315
               Left            =   5460
               MaxLength       =   3
               TabIndex        =   56
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               Caption         =   "Start Number"
               Height          =   255
               Left            =   120
               TabIndex        =   50
               Top             =   300
               Width           =   1035
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               Caption         =   "Minimum Number Width"
               Height          =   255
               Left            =   3540
               TabIndex        =   49
               Top             =   300
               Width           =   1815
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Subject Identifier Suffix"
            Height          =   1035
            Left            =   120
            TabIndex        =   46
            Top             =   2760
            Width           =   8715
            Begin VB.CheckBox chkBaseNumbering 
               Caption         =   "Restart numbering for new suffixes"
               Height          =   375
               Index           =   1
               Left            =   180
               TabIndex        =   62
               Top             =   600
               Width           =   3435
            End
            Begin VB.TextBox txtRegExpression 
               Height          =   315
               Index           =   1
               Left            =   1200
               TabIndex        =   60
               Top             =   240
               Width           =   6675
            End
            Begin VB.CommandButton cmdRegExpression 
               BackColor       =   &H00FFFFFF&
               Height          =   600
               Index           =   1
               Left            =   7980
               Picture         =   "frmStudyDefinition.frx":07F4
               Style           =   1  'Graphical
               TabIndex        =   61
               Top             =   180
               Width           =   600
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               Caption         =   "Expression"
               Height          =   255
               Left            =   180
               TabIndex        =   47
               Top             =   300
               Width           =   795
            End
         End
         Begin TabDlg.SSTab ssTabConditions 
            Height          =   2595
            Left            =   120
            TabIndex        =   68
            Top             =   3960
            Width           =   8715
            _ExtentX        =   15372
            _ExtentY        =   4577
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabHeight       =   520
            TabCaption(0)   =   "Registration Conditions"
            TabPicture(0)   =   "frmStudyDefinition.frx":0AFE
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame11(2)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Uniqueness Checks"
            TabPicture(1)   =   "frmStudyDefinition.frx":0B1A
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame11(3)"
            Tab(1).ControlCount=   1
            Begin VB.Frame Frame11 
               Height          =   2115
               Index           =   3
               Left            =   -74880
               TabIndex        =   81
               Top             =   360
               Width           =   8475
               Begin VB.CommandButton cmdUniqueExpression 
                  BackColor       =   &H00FFFFFF&
                  Height          =   600
                  Left            =   1980
                  Picture         =   "frmStudyDefinition.frx":0B36
                  Style           =   1  'Graphical
                  TabIndex        =   86
                  TabStop         =   0   'False
                  Top             =   1500
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.TextBox txtUniqueCode 
                  Height          =   315
                  Left            =   180
                  TabIndex        =   85
                  Text            =   "code"
                  Top             =   1620
                  Visible         =   0   'False
                  Width           =   675
               End
               Begin MSFlexGridLib.MSFlexGrid flxUnique 
                  Height          =   1755
                  Left            =   120
                  TabIndex        =   87
                  Top             =   240
                  Width           =   8235
                  _ExtentX        =   14526
                  _ExtentY        =   3096
                  _Version        =   393216
                  AllowBigSelection=   0   'False
               End
            End
            Begin VB.Frame Frame11 
               Height          =   2115
               Index           =   1
               Left            =   -74880
               TabIndex        =   76
               Top             =   360
               Width           =   6735
               Begin VB.TextBox txtCondition 
                  Height          =   315
                  Index           =   1
                  Left            =   900
                  TabIndex        =   79
                  Text            =   "condition"
                  Top             =   1620
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.TextBox txtCode 
                  Height          =   315
                  Index           =   1
                  Left            =   180
                  TabIndex        =   78
                  Text            =   "code"
                  Top             =   1620
                  Visible         =   0   'False
                  Width           =   675
               End
               Begin VB.CommandButton cmdCondition 
                  BackColor       =   &H00FFFFFF&
                  Height          =   600
                  Index           =   1
                  Left            =   1980
                  Picture         =   "frmStudyDefinition.frx":0E40
                  Style           =   1  'Graphical
                  TabIndex        =   77
                  Top             =   1500
                  Width           =   600
               End
               Begin MSFlexGridLib.MSFlexGrid flxConditions 
                  Height          =   1335
                  Index           =   1
                  Left            =   120
                  TabIndex        =   80
                  Top             =   180
                  Width           =   6495
                  _ExtentX        =   11456
                  _ExtentY        =   2355
                  _Version        =   393216
               End
            End
            Begin VB.Frame Frame11 
               Height          =   2115
               Index           =   0
               Left            =   -74880
               TabIndex        =   74
               Top             =   360
               Width           =   8535
               Begin MSComctlLib.ListView lvwDetails 
                  Height          =   1755
                  Index           =   5
                  Left            =   120
                  TabIndex        =   75
                  Top             =   240
                  Width           =   6975
                  _ExtentX        =   12303
                  _ExtentY        =   3096
                  View            =   3
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  FullRowSelect   =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   1
                  NumItems        =   0
               End
            End
            Begin VB.Frame Frame12 
               Height          =   2115
               Left            =   -74880
               TabIndex        =   72
               Top             =   360
               Width           =   8535
               Begin MSComctlLib.ListView lvwDetails 
                  Height          =   1755
                  Index           =   4
                  Left            =   120
                  TabIndex        =   73
                  Top             =   240
                  Width           =   6975
                  _ExtentX        =   12303
                  _ExtentY        =   3096
                  View            =   3
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  FullRowSelect   =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   1
                  NumItems        =   0
               End
            End
            Begin VB.Frame Frame10 
               Height          =   2115
               Left            =   -74880
               TabIndex        =   70
               Top             =   360
               Width           =   8535
               Begin MSComctlLib.ListView lvwDetails 
                  Height          =   1755
                  Index           =   3
                  Left            =   120
                  TabIndex        =   71
                  Top             =   240
                  Width           =   6975
                  _ExtentX        =   12303
                  _ExtentY        =   3096
                  View            =   3
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  FullRowSelect   =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   1
                  NumItems        =   0
               End
            End
            Begin VB.Frame Frame11 
               Height          =   2115
               Index           =   2
               Left            =   120
               TabIndex        =   69
               Top             =   360
               Width           =   8475
               Begin VB.CommandButton cmdRegCondition 
                  BackColor       =   &H00FFFFFF&
                  Height          =   600
                  Left            =   1980
                  Picture         =   "frmStudyDefinition.frx":114A
                  Style           =   1  'Graphical
                  TabIndex        =   83
                  TabStop         =   0   'False
                  Top             =   1500
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.TextBox txtRegCode 
                  Height          =   315
                  Left            =   180
                  TabIndex        =   82
                  Text            =   "code"
                  Top             =   1620
                  Visible         =   0   'False
                  Width           =   675
               End
               Begin MSFlexGridLib.MSFlexGrid flxReg 
                  Height          =   1755
                  Left            =   120
                  TabIndex        =   84
                  Top             =   240
                  Width           =   8235
                  _ExtentX        =   14526
                  _ExtentY        =   3096
                  _Version        =   393216
                  AllowBigSelection=   0   'False
               End
            End
         End
      End
      Begin VB.Frame Frame4 
         Height          =   6735
         Left            =   -74880
         TabIndex        =   44
         Top             =   360
         Width           =   8955
         Begin VB.Frame Frame15 
            Caption         =   "Date of Birth"
            Height          =   975
            Left            =   120
            TabIndex        =   90
            Top             =   1560
            Width           =   8715
            Begin VB.TextBox txtExpression 
               Height          =   315
               Index           =   1
               Left            =   180
               MaxLength       =   255
               TabIndex        =   16
               Top             =   360
               Width           =   7635
            End
            Begin VB.CommandButton cmdExpression 
               BackColor       =   &H00FFFFFF&
               Height          =   600
               Index           =   1
               Left            =   7920
               Picture         =   "frmStudyDefinition.frx":1454
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   240
               Width           =   600
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Gender"
            Height          =   1095
            Left            =   120
            TabIndex        =   89
            Top             =   2700
            Width           =   8715
            Begin VB.TextBox txtExpression 
               Height          =   315
               Index           =   0
               Left            =   180
               MaxLength       =   255
               TabIndex        =   18
               Top             =   480
               Width           =   7635
            End
            Begin VB.CommandButton cmdExpression 
               BackColor       =   &H00FFFFFF&
               Height          =   600
               Index           =   0
               Left            =   7920
               Picture         =   "frmStudyDefinition.frx":175E
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   360
               Width           =   600
            End
            Begin VB.Label Label9 
               Caption         =   "Arezzo expression which evaluates to 1 for Female or 2 for Male"
               Height          =   195
               Left            =   900
               TabIndex        =   91
               Top             =   180
               Width           =   4635
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "Subject Label"
            Height          =   1215
            Left            =   120
            TabIndex        =   88
            Top             =   240
            Width           =   8715
            Begin VB.CommandButton cmdTrialSubjectLabelExpression 
               BackColor       =   &H00FFFFFF&
               Height          =   600
               Left            =   7920
               Picture         =   "frmStudyDefinition.frx":1A68
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   240
               Width           =   600
            End
            Begin VB.CheckBox chkLocal 
               Caption         =   "Only to be stored locally"
               Height          =   195
               Left            =   180
               TabIndex        =   15
               Top             =   840
               Width           =   4155
            End
            Begin VB.TextBox txtTrialSubjectLabel 
               Height          =   315
               Left            =   180
               MaxLength       =   255
               TabIndex        =   13
               Top             =   360
               Width           =   7635
            End
         End
      End
      Begin VB.Frame Frame2 
         Height          =   6735
         Left            =   -74880
         TabIndex        =   42
         Top             =   360
         Width           =   9015
         Begin MSComctlLib.ListView lsvRequired 
            Height          =   6375
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   11245
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6675
         Left            =   -74880
         TabIndex        =   38
         Top             =   420
         Width           =   9015
         Begin VB.TextBox txtReasonGridEdit 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   360
            TabIndex        =   39
            Text            =   "Text control used to edit grid"
            Top             =   1140
            Visible         =   0   'False
            Width           =   2175
         End
         Begin MSFlexGridLib.MSFlexGrid gridReasons 
            Height          =   6375
            Left            =   120
            TabIndex        =   40
            Top             =   180
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   11245
            _Version        =   393216
            Cols            =   3
            FocusRect       =   2
         End
         Begin MSComctlLib.ImageList imglistSmallIcons 
            Left            =   735
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
         End
         Begin MSComctlLib.ImageList imglistLargeIcons 
            Left            =   0
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
         End
      End
      Begin VB.Frame fraStudyDetails 
         Height          =   6735
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   8955
         Begin VB.ComboBox cboSchemes 
            Height          =   315
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   3360
            Width           =   2415
         End
         Begin VB.TextBox txtStudyStatus 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   "Study Status"
            Top             =   3780
            Width           =   2415
         End
         Begin VB.CheckBox chkSingleUseDataItems 
            Caption         =   "Each question can be used once only"
            Height          =   375
            Left            =   5760
            TabIndex        =   12
            Top             =   4200
            Width           =   3015
         End
         Begin VB.ComboBox cboTrialType 
            Height          =   315
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2880
            Width           =   2415
         End
         Begin VB.TextBox txtStandardTime 
            Height          =   315
            Left            =   7200
            TabIndex        =   9
            Top             =   2880
            Width           =   1575
         End
         Begin VB.TextBox txtStandardDate 
            Height          =   315
            Left            =   7200
            TabIndex        =   8
            Top             =   2400
            Width           =   1575
         End
         Begin VB.TextBox txtTrialDescription 
            Height          =   1095
            Left            =   1380
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Text            =   "frmStudyDefinition.frx":1D72
            Top             =   720
            Width           =   7395
         End
         Begin VB.ComboBox cboPhase 
            Height          =   315
            ItemData        =   "frmStudyDefinition.frx":1D88
            Left            =   1380
            List            =   "frmStudyDefinition.frx":1D8A
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   2400
            Width           =   2415
         End
         Begin VB.TextBox txtKeywords 
            Height          =   315
            Left            =   1380
            MaxLength       =   255
            TabIndex        =   3
            Top             =   1920
            Width           =   7395
         End
         Begin VB.TextBox txtCurrentVersion 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7200
            TabIndex        =   11
            Top             =   3840
            Width           =   1575
         End
         Begin VB.TextBox txtClinicalTrialName 
            Height          =   315
            Left            =   1380
            TabIndex        =   1
            Text            =   "txtClinicalTrialName"
            Top             =   240
            Width           =   3135
         End
         Begin MSMask.MaskEdBox mskExpectedRecruitment 
            Height          =   315
            Left            =   7200
            TabIndex        =   10
            Top             =   3360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "CTC Scheme"
            Height          =   255
            Left            =   180
            TabIndex        =   37
            Top             =   3360
            Width           =   1095
         End
         Begin VB.Label lblTrialType 
            Alignment       =   1  'Right Justify
            Caption         =   "Study Type"
            Height          =   255
            Left            =   180
            TabIndex        =   35
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label lblStandardTime 
            Alignment       =   1  'Right Justify
            Caption         =   "Standard Time"
            Height          =   255
            Left            =   5880
            TabIndex        =   34
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Label lblStandardDate 
            Alignment       =   1  'Right Justify
            Caption         =   "Standard Date"
            Height          =   255
            Left            =   5760
            TabIndex        =   33
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Current Status"
            Height          =   375
            Left            =   60
            TabIndex        =   32
            Top             =   3810
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Phase"
            Height          =   375
            Left            =   60
            TabIndex        =   31
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Keywords"
            Height          =   255
            Left            =   60
            TabIndex        =   30
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Current Version"
            Height          =   255
            Left            =   5880
            TabIndex        =   29
            Top             =   3840
            Width           =   1215
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Expected Recruitment"
            Height          =   255
            Left            =   5520
            TabIndex        =   28
            Top             =   3360
            Width           =   1575
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Study Name"
            Height          =   195
            Left            =   405
            TabIndex        =   27
            Top             =   240
            Width           =   870
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            Height          =   255
            Left            =   60
            TabIndex        =   26
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.PictureBox Chart2 
         Height          =   4575
         Left            =   -74160
         ScaleHeight     =   4515
         ScaleWidth      =   7995
         TabIndex        =   24
         Top             =   960
         Width           =   8055
      End
      Begin VB.PictureBox Chart1 
         Height          =   4575
         Left            =   -74280
         ScaleHeight     =   4515
         ScaleWidth      =   8115
         TabIndex        =   23
         Top             =   960
         Width           =   8175
      End
   End
   Begin VB.TextBox txtAfterGridEditTab 
      Height          =   285
      Left            =   3000
      TabIndex        =   41
      Text            =   "text control after text for grid editing tab"
      Top             =   4320
      Width           =   2775
   End
   Begin VB.TextBox txtAfterGridTab 
      Height          =   285
      Left            =   600
      TabIndex        =   25
      Text            =   "text control after grid tab"
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label lblUserCode 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Left            =   60
      TabIndex        =   36
      Top             =   7380
      Width           =   3855
   End
End
Attribute VB_Name = "frmStudyDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998 - 2006. All Rights Reserved
'   File:       frmStudyDefinition.frm
'   Author:     Andrew Newbigging June 1997
'   Purpose:    Maintenance of basic details of a study
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   17  Joanne Lau              07/09/98    Details:
'                                           Replaced the Update table statements used to
'                                           update table DataItem by using the edit recordset
'                                           method. This prevents SQLServer crashing when
'                                           updating a column with a long text string.
'                                           txtTrialDescription_lostfocus and
'                                           txtKeywords_lostfocus.
'                                           Changed maxlength of textboxes -
'                                           txtTrialDescription,txtKeywords to accept a
'                                           limit of 255 chars. Bug id 424.
'       Mo Morris               7/1/99      SR 651
'The data validation of ExpectedRecruitment, Keywords and TrialDescription are handled by
'the lostfocus event on their respective controls. When these lostfocus subs are triggered
'by the form being closed (form unload or frmMenu.tlbMenu.button(gsTRIAL_LABEL)) and the
'contents found to be invalid, they force a re-Focus, which used to result in a crash if
'this form was no longer visible. mskExpectedRecruitment_LostFocus, txtKeywords_LostFocus
'and txtTrialDescription_LostFocus have all had a line added to check the form is visible.
'       Andrew Newbigging       29/4/99     SR 826
'   MOdified RefreshTrialDetails to set the status of the trial in the popup list
'       Mo Morris           2/5/99          SR 883
'       txtKeywords_Lostfocus changed so that Keywords can now contain Spaces as well as
'       Alpha characters, Numeric Characters and Commas.
'   NCJ 13/8/99 - Changed PfeCheckExpression to Arezzo version
'   21  Paul Norris 18/08/99    SR 884,888,492,1306,1181,1188,1497,1176
'                               Upgraded to use clsStudyDefintion class to maintain data
'   22  Paul Norris 23/08/99    Amended EnableOK() to correctly display
'                               buttons for new trials
'   23  Paul Norris 23/08/99    Amended chkSingleUseDataItems_Click() to check if rule
'                               is broken
'   24  Paul Norris 30/08/99    Commented out code to delete rows
'       Mo Morris   14/9/99     cmdTrialSubjectLabelExpression_Click changed. Call to
'                               frmExpressionbuilder replaced by call to frmCUIFunctionEditor
'   PN  20/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   PN  22/09/99    Amended Form_QueryUnload() to remove cancel statement
'   PN  28/09/99    Amended RefreshGrid() to left justify all data
'                   Added moStudyDefinition_StandardDateValid() etc to inform user of
'                   invalid data
'   WillC 11/11/99  Added the Error handlers
'   WillC 15/11/99  Added a label to show the userId
'   NCJ 13 Dec 99   Ids to Long
'   SDM 15/12/99    Required questions. SR1177
'   Mo Morris   20/12/99    Changes from word 'TRIAL' to 'STUDY'
'   NCJ 21 Dec 99   Replaced cboStatus with txtStudyStatus
'   NCJ 22 Dec 99, SR 1764  Changed all controls' background colour to "Windows Background"
'   TA 29/04/2000   subclassing removed
'   WillC SR1764  22/8/00  Need to set the backcolor of some textboxes, explicitly fixed error handlers
'   TA 15/09/2000 - additional controls for handling CTC Scheme, Gender and DOB expressions
'   TA 04/10/2000: Laboratory, CTC Scheme, Clinical Test and Clinical Test Ids replace with Codes
'   TA 17/10/2000: Redesigned layout and added Registration and subject details tabs
'   TA 21/10/2000: Registration conditions and Uniqueness checks added -
'                       the validation is done in the form though ideally we need a new class for each of these
'                       and classes to wrap the relationship between the input grid class and these new classes
'   TA 05/12/2000: Tab Indexes altered so that tabbing around the registration grid works
'   TA 06/12/2000: Minor code changes from white box test
'   ZA 20/08/2002  Added RFCDefault check box on RFC tab
'   ZA 10/09/2002  Removed RFCDefault check box and placed under OPtions menu
'   ZA 30/09/2002  Check if the user clicked inside the grid cell or not - bug fix
'   NCJ 16 Jun 03 - Corrected erronoeous message about discarding changes (MACRO 3.0 Bug 1004)
'   NCJ May/Jun 06 - MUSD - Consider access mode
'   NCJ 20 Sept 06 - Allow Expression Editor buttons even in RO/LO modes
'   NCJ 21 Sept 06 - Call CUIFunctionEditor with appropriate Access Mode
'   NCJ 26 Feb 07 - Bug 2029 - Check registration visit and eForm exist in DisplayRegistration
'------------------------------------------------------------------------------------'

Option Explicit


'constants for registration
Private Const m_TAB_REG As Integer = 0
Private Const m_TAB_unique As Integer = 1

Private Const m_INDEX_PREFIX As Integer = 0
Private Const m_INDEX_SUFFIX As Integer = 1

Private Const m_COL_CODE = 1
Private Const m_COL_AREZZO = 2

'Registration class not contained within StudyDefinition class so it can be moved to separate form
Private WithEvents moRegistration As clsRegistration
Attribute moRegistration.VB_VarHelpID = -1


'TA 15/09/2000: constants for Gender and DOB expression controls
Private Const m_INDEX_GENDER = 0
Private Const m_INDEX_DOB = 1

Private Const m_TAB_DETAILS = 0
Private Const m_TAB_SUBJECT = 1
Private Const m_TAB_REGISTRATION = 2
Private Const m_TAB_CHANGE = 3
Private Const m_TAB_REQUIRED = 4
Private Const m_TAB_OVERRULE = 5
Private Const ms_RFC_GRID_NAME = "gridReasons"
Private Const ms_OVERRULE_GRID_NAME = "gridOverrule"

Private Const mcREASONCOL = 1
Private Const mcINVISIBLECOL = 2

'link between grid and table
Private moRegInputGrid As clsInputGrid
Private moUniqueInputGrid As clsInputGrid

' variable to track if the click was a right click
Private mbIsRightMouseClick As Boolean

Private WithEvents moStudyDefinition As clsStudyDefinition
Attribute moStudyDefinition.VB_VarHelpID = -1

Private mtSelectedGridCell As SelectedGridCell
Private mnRefreshOverruleGrid As Boolean
Private msClinicalTrialName As String 'used to store form property
Private mlClinicalTrialId As Long 'used to store form property
Private mnVersionId As Integer 'used to store form property
Private msUpdateMode As String
Private mbIsLoading As Boolean
Private mbIsNew As Boolean

Private mbRefreshRFCGrid As Boolean
Private mbRefreshRFOGrid As Boolean

'ZA 30/09/2002 - Is click within the cell of grid?
Private mbClickInGridCell As Boolean

Private mtSelectedRFOCell As SelectedGridCell
Private msGridWithFocus As String

' NCJ 16 Jun 03 - Record whthere they changed anything
Private mbChanged As Boolean

'---------------------------------------------------------------------
Private Sub cboSchemes_Click()
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    If moStudyDefinition.CTCSchemeCode <> cboSchemes.Text Then
        moStudyDefinition.CTCSchemeCode = cboSchemes.Text
        mbChanged = True
    End If

End Sub

'---------------------------------------------------------------------
Private Sub cmdExpression_Click(Index As Integer)
'---------------------------------------------------------------------

    Select Case Index
    Case m_INDEX_GENDER
        frmCUIFunctionEditor.Initialise txtExpression(Index), "Expression", _
                        (frmMenu.StudyAccessMode >= sdReadWrite), "Gender"
    Case m_INDEX_DOB
        frmCUIFunctionEditor.Initialise txtExpression(Index), "Expression", _
                        (frmMenu.StudyAccessMode >= sdReadWrite), "Date of Birth"
    End Select
    
End Sub



'---------------------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------------------
'   ATN 17/12/99    SR 1763
'   Need to reset the selected item
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    frmMenu.ChangeSelectedItem "", ""
    
    'WillC SR1764  Need to set the backcolor of these, explicitly
    txtStandardDate.BackColor = vbWindowBackground
    txtStandardTime.BackColor = vbWindowBackground
    txtTrialSubjectLabel.BackColor = vbWindowBackground
    mskExpectedRecruitment.BackColor = vbWindowBackground
    
Exit Sub
ErrHandler:
    If MACROErrorHandler("frmStudyDefinition", Err.Number, Err.Description, _
                            "Form_Activate", Err.Source) = Retry Then
        Resume
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set moRegistration = Nothing

End Sub

'---------------------------------------------------------------------
Private Sub gridOverrule_Click()
'---------------------------------------------------------------------
'ZA 30/09/2002
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If Not mbIsRightMouseClick Then
        Call HandleGridClick(gridOverrule)
        msGridWithFocus = ms_OVERRULE_GRID_NAME
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "gridOverrule_Click")
        Case OnErrorAction.Retry
            Resume
    End Select

    
End Sub

'---------------------------------------------------------------------
Private Sub gridOverrule_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
'ZA 30/09/2002
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If gridOverrule.Col = mcREASONCOL Then
        If KeyAscii = vbKeyReturn Then
            Call SetupNextCellSelection(gridOverrule)
            Call gridOverrule_Click
        End If
    Else
        Call DoGridEdit(KeyAscii, gridOverrule, txtOverruleEdit)
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "gridOverrule_KeyPress")
            Case OnErrorAction.Retry
            Resume
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub GridMouseDown(Button As Integer, X As Single, Y As Single, oGrid As MSFlexGrid)
'---------------------------------------------------------------------
' Mouse down in grid
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
   
    ' NCJ 14 Jun 06 - Ignore if not read/write
    If frmMenu.StudyAccessMode < sdReadWrite Then Exit Sub
    
    ' maintain the flag for the mouse click
    If Button = vbRightButton Then
        mbIsRightMouseClick = True
    
    Else
        mbIsRightMouseClick = False
        'ZA 30/09/2002 - check if user clicked inside a cell
        With oGrid
            If (X > .ColPos(.Col) + .ColWidth(.Col)) _
              Or (Y > .RowPos(.Row) + .RowHeight(.Row)) Then
                mbClickInGridCell = False
            Else
                mbClickInGridCell = True
            End If
        End With
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "GridMouseDown")
        Case OnErrorAction.Retry
            Resume
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub gridOverrule_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
'ZA 30/09/2002
'---------------------------------------------------------------------
    
    Call GridMouseDown(Button, X, Y, gridOverrule)

End Sub

'---------------------------------------------------------------------
Private Sub gridOverrule_Scroll()
'---------------------------------------------------------------------
'ZA 30/09/2002
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    Call SetGridText(gridOverrule)
    DoEvents
    gridOverrule.Refresh

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "gridOverrule_Scroll")
        Case OnErrorAction.Retry
            Resume
    End Select
End Sub

'---------------------------------------------------------------------
Private Sub gridReasons_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
        
    Call GridMouseDown(Button, X, Y, gridReasons)

End Sub

'---------------------------------------------------------------------
Private Sub lsvRequired_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'---------------------------------------------------------------------
    'SDM 14/12/99 SR1177    Stop the check mark from being altered
    '                       by the user
    Item.Checked = Not Item.Checked
End Sub

'---------------------------------------------------------------------
Public Property Let IsNew(bIsNew As Boolean)
'---------------------------------------------------------------------

    mbIsNew = bIsNew

End Property

'---------------------------------------------------------------------
Public Property Get ClinicalTrialId() As Long
'---------------------------------------------------------------------

    ClinicalTrialId = mlClinicalTrialId ' return value

End Property

'---------------------------------------------------------------------
Public Property Let ClinicalTrialId(lTrialId As Long)
'---------------------------------------------------------------------

    mlClinicalTrialId = lTrialId ' store value

End Property

'---------------------------------------------------------------------
Public Property Get VersionId() As Integer
'---------------------------------------------------------------------

    VersionId = mnVersionId ' return value

End Property

'---------------------------------------------------------------------
Public Property Let VersionId(tmpX As Integer)
'---------------------------------------------------------------------

    mnVersionId = tmpX ' store value

End Property

'---------------------------------------------------------------------
Public Property Get ClinicalTrialName() As String
'---------------------------------------------------------------------

    ClinicalTrialName = msClinicalTrialName

End Property

'---------------------------------------------------------------------
Public Property Let ClinicalTrialName(tmpStr As String)
'---------------------------------------------------------------------

    msClinicalTrialName = tmpStr

End Property

'---------------------------------------------------------------------
Private Sub cboPhase_Click()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If mbIsLoading Then Exit Sub
    moStudyDefinition.Phase = cboPhase
    mbChanged = True
    
Exit Sub
ErrHandler:
    If MACROErrorHandler("frmStudyDefinition", Err.Number, Err.Description, _
                            "cboPhase_Click", Err.Source) = Retry Then
        Resume
    End If
        
End Sub

'---------------------------------------------------------------------
Private Sub cboTrialType_Click()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If mbIsLoading Then Exit Sub
    moStudyDefinition.TrialType = cboTrialType
    mbChanged = True
    
Exit Sub
ErrHandler:
    If MACROErrorHandler("frmStudyDefinition", Err.Number, Err.Description, _
                            "cboTrialType_Click", Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub chkSingleUseDataItems_Click()
'---------------------------------------------------------------------
Dim sMSG As String
    
    On Error GoTo ErrHandler
    
    If chkSingleUseDataItems.Value = vbChecked Then
        ' the check box was checked so check if the action can be completed
        
        If moStudyDefinition.IsSingleUseRuleBroken(ClinicalTrialId, VersionId) Then
            sMSG = "At least one question is used more than once." & vbCr
            sMSG = sMSG & "To complete this action you must ensure that all questions have a single use."
            MsgBox sMSG, vbCritical, GetApplicationTitle
            chkSingleUseDataItems.Value = vbUnchecked
            mbChanged = True
        Else
            If Not mbIsLoading Then
                moStudyDefinition.SingleUseDataItems = chkSingleUseDataItems.Value
                mbChanged = True
            End If
                
        End If
        
    Else
    '***
        If Not mbIsLoading Then
            moStudyDefinition.SingleUseDataItems = chkSingleUseDataItems.Value
            mbChanged = True
        End If
        
    End If

Exit Sub
ErrHandler:
    If MACROErrorHandler("frmStudyDefinition", Err.Number, Err.Description, _
                            "chkSingleUseDataItems_Click", Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdCancel_Click()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    ' prompt user to save if changes have been made
    If HandleSave(False) Then
        Call frmMenu.HideStudyDefinition(False)

    End If

Exit Sub
ErrHandler:
    If MACROErrorHandler("frmStudyDefinition", Err.Number, Err.Description, _
                            "cmdCancel_Click", Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdOK_Click()
'---------------------------------------------------------------------
    ' save if changes have been made
    'ZA 07/10/2002 - TODO
    mbIsNew = mbIsNew And Not HandleSave(True)
    Call frmMenu.HideStudyDefinition(True)
    frmTrialList.cmdRefresh.Value = True
    
End Sub

'---------------------------------------------------------------------
Public Function SaveChanges(Optional bSaveWithoutPrompt As Boolean = True) As Boolean
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If Not moStudyDefinition Is Nothing Then
        'ZA 07/10/2002 - TODO
        SaveChanges = HandleSave(bSaveWithoutPrompt)
    Else
        SaveChanges = True
    End If

Exit Function
ErrHandler:
    If MACROErrorHandler("frmStudyDefinition", Err.Number, Err.Description, _
                            "SaveChanges", Err.Source) = Retry Then
        Resume
    End If

End Function

'---------------------------------------------------------------------
Private Function HandleSave(bSaveWithoutPrompt As Boolean) As Boolean
'---------------------------------------------------------------------
' If bSaveWithoutPrompt = TRUE, don't give user messages
'---------------------------------------------------------------------
Dim sMSG As String
Dim bActionComplete As Boolean
    
    On Error GoTo ErrHandler
    
    sMSG = "Are you sure you wish to close without saving changes to study details?"
    
    If Not (moStudyDefinition Is Nothing) And Not (moRegistration Is Nothing) Then
        If (moStudyDefinition.IsDirty Or moRegistration.IsDirty) Then
                        
            ' changes have been made
            
            If moStudyDefinition.IsValid And moRegistration.IsValid And RegGridsAreValid Then
                ' the changes are valid
                If bSaveWithoutPrompt Then
                    ' just save
                    moStudyDefinition.Save mlClinicalTrialId, mnVersionId
                    'TA 17/11/2000: save registration details
                    SaveRegistration
                    bActionComplete = True

                Else
                    ' ask to save
'                    sMsg = "This will discard the changes you have made to this study's definition"
                    Select Case DialogQuestion(sMSG, , False)

                    'following code if we want to give the option of saving changes on a cancel
'                    Case vbYes
'                        'TA 17/11/2000: save registration details
'                        SaveRegistration
'                        moStudyDefinition.Save mlClinicalTrialId, mnVersionId
'                        ' action not cancelled
'                        bActionComplete = True

                    Case vbYes
                        ' action not cancelled
                        bActionComplete = True

                    Case vbNo
                        ' action cancelled
                        bActionComplete = False

                    End Select

                End If
        
            Else
                ' the changes are not valid
'                sMsg = "This will discard the changes you have made to this study's definition"

                If bSaveWithoutPrompt Then
                    ' Just go away without saying anything (or saving anything)
                    bActionComplete = True
                Else
                    ' Give them a prompt
                    Select Case DialogQuestion(sMSG, , False)
                    Case vbYes
                        ' action not cancelled
                        bActionComplete = True
                        
                    Case Else
                        ' action cancelled
                        bActionComplete = False
                        
                    End Select
                End If
            End If
        Else
            ' No changes to save
            bActionComplete = True
            
        End If
        
        If bActionComplete Then
            ' unload the object to free up some memory
            moStudyDefinition.Discard
            Set moStudyDefinition = Nothing
            Set moRegistration = Nothing
        End If

    Else
        bActionComplete = True
        
    End If
    
    HandleSave = bActionComplete

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmStudyDefinition.HandleSave"
    
End Function

'---------------------------------------------------------------------
Private Sub moStudyDefinition_AddedNewReason(enReasonType As eReasonType)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If Not mbIsLoading Then
        ' the last item in the grid is now valid so add a new one
        ' then refresh the grid
        Select Case enReasonType
        Case eReasonType.ReasonForChange
            mbRefreshRFCGrid = True
        Case eReasonType.ReasonForOverrule
            mbRefreshRFOGrid = True
        End Select
    End If
    
    Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "moStudyDefinition_AddedNewReason")
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
Private Sub moStudyDefinition_HasChanges()
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If moStudyDefinition.IsValid Then
        Call EnableOK(True)
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "moStudyDefinition_HasChanges")
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
Private Sub moStudyDefinition_IsValid(bValid As Boolean)
'---------------------------------------------------------------------

    Call EnableOK(bValid)
    
End Sub

'---------------------------------------------------------------------
Private Sub EnableOK(bEnable As Boolean)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If mbIsNew Then
        ' new trials will always have cancel disabled and ok enabled
        cmdOK.Enabled = True
        cmdCancel.Enabled = False
        
    Else
        cmdCancel.Enabled = True
        cmdOK.Enabled = False
        
        If bEnable And frmMenu.StudyAccessMode >= sdReadWrite Then
            If Not (moStudyDefinition Is Nothing) And Not (moRegistration Is Nothing) Then
                ' only if there are changes will OK button be enabled
                If (moStudyDefinition.IsDirty Or moRegistration.IsDirty) _
                        And moStudyDefinition.IsValid And moRegistration.IsValid _
                        And RegGridsAreValid Then
                    cmdOK.Enabled = True
                End If
            End If
        End If
    
    End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmStudyDefinition.EnableOK"
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdTrialSubjectLabelExpression_Click()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    'Mo Morris 14/9/99
    'call to frmExpressionbuilder replaced by call to frmCUIFunctionEditor
    'txtTrialSubjectLabel.SetFocus
    'frmExpressionBuilder.EditExpression txtTrialSubjectLabel
    frmCUIFunctionEditor.Initialise txtTrialSubjectLabel, "Expression", _
                    (frmMenu.StudyAccessMode >= sdReadWrite), "Subject Label"
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdTrialSubjectLabelExpression_Click")
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
On Error GoTo ErrHandler

    Me.Icon = frmMenu.Icon
    
    ' Turn on key preview for form, so that F1 (Help) can be trapped by form
    Me.KeyPreview = True
    
    ' Specify position relative to mdiForm
    Me.Top = 0
    Me.Left = 0
    'lblUserCode.Caption = moStudyDefinition.Load
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
Private Sub moStudyDefinition_StandardDateValid(bIsValid As Boolean)
'---------------------------------------------------------------------
' PN 28/09/99 created to display invalid data
' this event is raised when the data becomes valid or invalid
'---------------------------------------------------------------------
On Error GoTo ErrHandler
    Call SetBackColour(txtStandardDate, bIsValid)
    

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "moStudyDefinition_StandardDateValid")
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
Private Sub moStudyDefinition_StandardTimeValid(bIsValid As Boolean)
'---------------------------------------------------------------------
' PN 28/09/99 created to display invalid data
' this event is raised when the data becomes valid or invalid
'---------------------------------------------------------------------
    Call SetBackColour(txtStandardTime, bIsValid)
    
End Sub
'---------------------------------------------------------------------
Private Sub SetBackColour(oTextBox As TextBox, bIsValid As Boolean)
'---------------------------------------------------------------------
' PN 28/09/99 created to display invalid data
' repaint the back colour of the textbox to indicate that the data is
' invalid
'---------------------------------------------------------------------
On Error GoTo ErrHandler
    
    If bIsValid Then
        oTextBox.BackColor = vbWhite
        
    Else
        oTextBox.BackColor = vbYellow
    
    End If
    

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "SetBackColour")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

        
End Sub

Private Sub moStudyDefinition_SubjectDOBValid(bIsValid As Boolean)
'---------------------------------------------------------------------
' TA 15/9/00 created to display invalid data
' this event is raised when the data becomes valid or invalid
'---------------------------------------------------------------------
    
    Call SetBackColour(txtExpression(m_INDEX_DOB), bIsValid)
    
End Sub

Private Sub moStudyDefinition_SubjectGenderValid(bIsValid As Boolean)
'---------------------------------------------------------------------
' TA 15/9/00 created to display invalid data
' this event is raised when the data becomes valid or invalid
'---------------------------------------------------------------------
    
    Call SetBackColour(txtExpression(m_INDEX_GENDER), bIsValid)
    
End Sub

'---------------------------------------------------------------------
Private Sub moStudyDefinition_TrialSubjectLabelValid(bIsValid As Boolean)
'---------------------------------------------------------------------
' PN 28/09/99 created to display invalid data
' this event is raised when the data becomes valid or invalid
'---------------------------------------------------------------------
    
    Call SetBackColour(txtTrialSubjectLabel, bIsValid)
    
End Sub

'---------------------------------------------------------------------
Private Sub mskExpectedRecruitment_Change()
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If Not mbIsLoading Then
        Call TextChange(mskExpectedRecruitment, moStudyDefinition, "ExpectedRecruitment")
        mbChanged = True
    End If

Exit Sub
ErrHandler:
    If MACROErrorHandler("frmStudyDefinition", Err.Number, Err.Description, _
                            "mskExpectedRecruitment_Change", Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub mskExpectedRecruitment_LostFocus()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    mskExpectedRecruitment = TextLostFocus(mskExpectedRecruitment, _
                                moStudyDefinition, "ExpectedRecruitment")

Exit Sub
ErrHandler:
    If MACROErrorHandler("frmStudyDefinition", Err.Number, Err.Description, _
                            "mskExpectedRecruitment_LostFocus", Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub SSTab1_Click(PreviousTab As Integer)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If SSTab1.Tab = m_TAB_REQUIRED Then
        Call PopulateRequiredTab
    End If
    

Exit Sub
ErrHandler:
    If MACROErrorHandler("frmStudyDefinition", Err.Number, Err.Description, _
                            "SSTab1_Click", Err.Source) = Retry Then
        Resume
    End If
    
End Sub


'---------------------------------------------------------------------
Private Sub txtExpression_Change(Index As Integer)
'---------------------------------------------------------------------

'---------------------------------------------------------------------
    
    If Not mbIsLoading Then
        If Index = m_INDEX_GENDER Then
            Call TextChange(txtExpression(Index), moStudyDefinition, "SubjectGender")
        Else
            Call TextChange(txtExpression(Index), moStudyDefinition, "SubjectDOB")
        End If
        mbChanged = True
    End If

End Sub

'---------------------------------------------------------------------
Private Sub txtExpression_LostFocus(Index As Integer)
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If Index = m_INDEX_GENDER Then
        txtExpression(Index).Text = TextLostFocus(txtExpression(Index), _
                            moStudyDefinition, "SubjectGender")
    Else
        txtExpression(Index).Text = TextLostFocus(txtExpression(Index), _
                            moStudyDefinition, "SubjectDOB")
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtTrialSubjectLabel_LostFocus")
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
Private Sub txtKeywords_Change()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If Not mbIsLoading Then
        Call TextChange(txtKeywords, moStudyDefinition, "Keywords")
        mbChanged = True

    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtKeywords_Change")
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
Private Sub txtKeywords_LostFocus()
'---------------------------------------------------------------------
On Error GoTo ErrHandler
    txtKeywords = TextLostFocus(txtKeywords, moStudyDefinition, "Keywords")

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtKeywords_LostFocus")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

        
End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtOverruleEdit_Change()
'----------------------------------------------------------------------------------------'
'ZA 30/09/2002
'----------------------------------------------------------------------------------------'

    ' copy the change accross to the object
    ' let the object validate the data input
    On Error GoTo ErrHandler
    If txtOverruleEdit.Visible Then
        Call UpdateReasonsWithEdits(gridOverrule, txtOverruleEdit, moStudyDefinition.ReasonsForOverrule)
        mbChanged = True

    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtOverruleEdit_Change")
        Case OnErrorAction.Retry
            Resume
    End Select
End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtOverruleEdit_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler
    
    With gridOverrule
        Select Case KeyCode
        Case vbKeyReturn
            '.SetFocus
            DoEvents
            If mbRefreshRFOGrid Then
                Call RefreshOverruleGrid
            End If
            msGridWithFocus = ms_OVERRULE_GRID_NAME
            txtAfterGridEditTab.SetFocus
            
         End Select
     
     End With

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtOverruleEdit_KeyDown")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtOverruleEdit_KeyPress(KeyAscii As Integer)
'----------------------------------------------------------------------------------------'
'noise suppression
   
   If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtOverruleEdit_LostFocus()
'----------------------------------------------------------------------------------------'
On Error GoTo ErrHandler

    If mbRefreshRFOGrid And Me.Visible Then
        Call RefreshOverruleGrid
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtOverruleEdit_LostFocus")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtRegCode_LostFocus()
'----------------------------------------------------------------------------------------'

    Call EnableOK(True)

End Sub

'---------------------------------------------------------------------
Private Sub txtStandardDate_Change()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If Not mbIsLoading Then
        Call TextChange(txtStandardDate, moStudyDefinition, "StandardDateFormat")
        mbChanged = True
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "ReorderVisits")
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
Private Sub txtStandardDate_LostFocus()
'---------------------------------------------------------------------
On Error GoTo ErrHandler
    txtStandardDate = TextLostFocus(txtStandardDate, moStudyDefinition, "StandardDateFormat")

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtStandardDate_LostFocus")
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
Private Sub txtStandardTime_Change()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    If Not mbIsLoading Then
        Call TextChange(txtStandardTime, moStudyDefinition, "StandardTimeFormat")
        mbChanged = True
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtStandardTime_Change")
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
Private Sub txtStandardTime_LostFocus()
'---------------------------------------------------------------------
On Error GoTo ErrHandler
    txtStandardTime = TextLostFocus(txtStandardTime, moStudyDefinition, "StandardTimeFormat")
    

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtStandardTime_LostFocus")
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
Private Sub txtTrialDescription_Change()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If Not mbIsLoading Then
        Call TextChange(txtTrialDescription, moStudyDefinition, "TrialDescription")
        mbChanged = True
        
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtTrialDescription_Change")
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
Private Sub txtTrialDescription_LostFocus()
'---------------------------------------------------------------------
On Error GoTo ErrHandler
    txtTrialDescription = TextLostFocus(txtTrialDescription, moStudyDefinition, "TrialDescription")

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtTrialDescription_LostFocus")
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
Private Sub txtTrialSubjectLabel_Change()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If Not mbIsLoading Then
        Call TextChange(txtTrialSubjectLabel, moStudyDefinition, "TrialSubjectLabel")
        mbChanged = True
        
    End If
    
    ' the lost focus fires after the form is hidden.
    ' this will cause an error if the form is hidden since setting chkLocal.Value
    ' fires the chkLocal_Click event which calls moStudyDefinition.LocalTrialSubjectLabel.
    ' however, the moStudyDefinition object is destroyed when the form is hidden to free some memory
    If Me.Visible Then
        If txtTrialSubjectLabel.Text > vbNullString Then
            chkLocal.Enabled = True
        Else
            chkLocal.Value = vbUnchecked
            chkLocal.Enabled = False
        End If
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtTrialSubjectLabel_Change")
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
Private Sub txtTrialSubjectLabel_LostFocus()
'---------------------------------------------------------------------
On Error GoTo ErrHandler
    txtTrialSubjectLabel = TextLostFocus(txtTrialSubjectLabel, _
                            moStudyDefinition, "TrialSubjectLabel")

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtTrialSubjectLabel_LostFocus")
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
Private Sub chkLocal_Click()
'---------------------------------------------------------------------

    If Not mbIsLoading Then
        moStudyDefinition.LocalTrialSubjectLabel = chkLocal.Value
        mbChanged = True
    End If

End Sub

'---------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'---------------------------------------------------------------------
Dim bSaveWithoutPrompt As Boolean

    ' Only prompt for saving if user has clicked the Close box
    bSaveWithoutPrompt = (UnloadMode <> vbFormControlMenu)
    If HandleSave(bSaveWithoutPrompt) Then
        ' the user did not click cancel
        frmMenu.HideStudyDefinition
    Else
        Cancel = 1

    End If
    
End Sub

'---------------------------------------------------------------------
Public Property Get UpdateMode() As String
'---------------------------------------------------------------------

    UpdateMode = msUpdateMode

End Property

'---------------------------------------------------------------------
Public Property Let UpdateMode(tmpMode As String)
'---------------------------------------------------------------------

    msUpdateMode = tmpMode

End Property

'---------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------
On Error GoTo ErrHandler
    If KeyCode = vbKeyF1 Then               ' Show user guide
        'ShowDocument Me.hWnd, gsMACROUserGuidePath
        
        'REM 07/12/01 - New Call for MACRO Help
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
Public Sub RefreshTrialDetails()
'---------------------------------------------------------------------
' NCJ 21/12/99 - Status is now a text field
' NCJ 14 Jun 06 - Consider study access; field enabling moved to other routines
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    mbIsLoading = True
    Me.Hide
    
    Me.Caption = ClinicalTrialName & " Study Definition"
    
    txtClinicalTrialName = ClinicalTrialName
    VersionId = gnCurrentVersionId(ClinicalTrialId)
    txtCurrentVersion.Text = VersionId
    txtClinicalTrialName.Enabled = False
    txtCurrentVersion.Enabled = False
    
    SSTab1.TabVisible(m_TAB_DETAILS) = True
    SSTab1.TabVisible(m_TAB_CHANGE) = True
    SSTab1.Tab = m_TAB_DETAILS
    
    Set moStudyDefinition = New clsStudyDefinition
    moStudyDefinition.Load mlClinicalTrialId, mnVersionId
    
    'ZA 30/09/2002 - add RFC and RFO
    moStudyDefinition.ReasonsForOverrule.AddNew
    moStudyDefinition.ReasonsForChange.AddNew
    
    Call DisplayData
      
    'set up registration data
    Call DisplayRegistration
    
    'TA 29/11/00 - only allow registration editing if they have the appropriate permission
    SSTab1.TabEnabled(m_TAB_REGISTRATION) = goUser.CheckPermission(gsFnMaintainRegistration)
    
    mbIsLoading = False
    mbChanged = False

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "RefreshTrialDetails")
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
Private Sub DisplayData()
'---------------------------------------------------------------------
' Display data for study
' NCJ 21 Dec 99 - Status is now a text box rather than a combo
' NCJ 14 Jun 06 - Disable things for read-only access
'---------------------------------------------------------------------
Dim bCanEdit As Boolean

    On Error GoTo ErrHandler
    
    ' NCJ 14 Jun 06 - Decide whether they can edit
    bCanEdit = frmMenu.StudyAccessMode >= sdReadWrite
    
    mtSelectedGridCell.Col = 1
    mtSelectedGridCell.Row = 1
    'ZA 01/10/2002 - over rule grid cell
    mtSelectedRFOCell.Col = 1
    mtSelectedRFOCell.Row = 1
    
    With moStudyDefinition
        ' Phases
        Call LoadCombo(cboPhase, .Phases)
        SetComboDropdownWidth Me, cboPhase
        If .Phase <> vbNullString Then
            cboPhase = .Phase
        End If
        cboPhase.Enabled = bCanEdit
        
        ' Trial Types
        Call LoadCombo(cboTrialType, .TrialTypes)
        SetComboDropdownWidth Me, cboTrialType
        If .TrialType <> vbNullString Then
            cboTrialType = .TrialType
        End If
        cboTrialType.Enabled = bCanEdit
        
        ' CTC Schemes
        Call LoadCTCSchemeCombo
        SetComboDropdownWidth Me, cboSchemes
        If .CTCSchemeCode > "" Then
            cboSchemes.Text = .CTCSchemeCode
        Else
            cboSchemes.ListIndex = 0
        End If
        cboSchemes.Enabled = bCanEdit
        
        txtStudyStatus.Text = RemoveNull(.Status)
        txtStudyStatus.Enabled = bCanEdit
        
        txtTrialDescription = .TrialDescription
        txtTrialDescription.Enabled = bCanEdit
        mskExpectedRecruitment = .ExpectedRecruitment
        mskExpectedRecruitment.Enabled = bCanEdit
        txtKeywords = .Keywords
        txtKeywords.Enabled = bCanEdit
        
        txtStandardDate = .StandardDateFormat
        txtStandardDate.Enabled = bCanEdit
        txtStandardTime = .StandardTimeFormat
        txtStandardTime.Enabled = bCanEdit
        
        chkSingleUseDataItems.Value = .SingleUseDataItems
        chkSingleUseDataItems.Enabled = bCanEdit
        
        'Mo Morris 30/8/01 Db Audit (UserId to UserName)
        lblUserCode.Caption = "Creator: " & .UserName
        
        txtTrialSubjectLabel = .TrialSubjectLabel
        txtTrialSubjectLabel.Enabled = bCanEdit
        cmdTrialSubjectLabelExpression.Enabled = bCanEdit Or txtTrialSubjectLabel.Text > ""
        
        If .LocalTrialSubjectLabel = 1 And .TrialSubjectLabel <> vbNullString Then
            chkLocal.Value = vbChecked
        Else
            chkLocal.Value = vbUnchecked
        End If
        If .TrialSubjectLabel = vbNullString Or Not bCanEdit Then
            chkLocal.Enabled = False
        Else
            chkLocal.Enabled = bCanEdit
        End If
        
        ' NCJ 20 Sept 06 - Allow Expression Editor buttons even in RO/LO modes
        txtExpression(m_INDEX_GENDER).Text = .SubjectGender
        txtExpression(m_INDEX_GENDER).Enabled = bCanEdit
        cmdExpression(m_INDEX_GENDER).Enabled = bCanEdit Or txtExpression(m_INDEX_GENDER).Text > ""
        txtExpression(m_INDEX_DOB).Text = .SubjectDOB
        txtExpression(m_INDEX_DOB).Enabled = bCanEdit
        cmdExpression(m_INDEX_DOB).Enabled = bCanEdit Or txtExpression(m_INDEX_DOB).Text > ""
        Call EnableOK(True)

    End With
    
    Call RefreshRFCGrid
    
    'ZA 30/09/2002 -
    Call RefreshOverruleGrid
    'no RFC or Overrule grid has focus at this stage
    msGridWithFocus = ""
    
    Me.Show

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmStudyDefinition.DisplayData"
    
End Sub

'---------------------------------------------------------------------
Private Sub gridReasons_Click()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If Not mbIsRightMouseClick Then
        Call HandleGridClick(gridReasons)
        msGridWithFocus = ms_RFC_GRID_NAME
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "gridReasons_Click")
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
Private Sub HandleGridClick(ByRef oGrid As MSFlexGrid)
'---------------------------------------------------------------------
Dim nNewRow As Integer
Dim nNewCol As Integer
    
    On Error GoTo ErrHandler

    ' NCJ 14 Jun 06 - Ignore if not read/write
    If frmMenu.StudyAccessMode < sdReadWrite Then Exit Sub
    
    With oGrid
        ' save new selected cell
        nNewCol = .Col
        nNewRow = .Row
        
        ' set back to old cell position
        'ZA 01/10/2002 - check for the grid first
        If oGrid.Name = ms_OVERRULE_GRID_NAME Then
            .Col = mtSelectedRFOCell.Col
            .Row = mtSelectedRFOCell.Row
        Else
            .Col = mtSelectedGridCell.Col
            .Row = mtSelectedGridCell.Row
        End If
    
        Call SetGridText(oGrid)
        
        ' reset back to new selected cell
        .Col = nNewCol
        .Row = nNewRow
        
        ' set into edit mode
        If oGrid.Name = ms_RFC_GRID_NAME Then
            Call DoGridEdit(Asc(" "), gridReasons, txtReasonGridEdit)
        Else
            Call DoGridEdit(Asc(" "), gridOverrule, txtOverruleEdit)
        End If
        DoEvents
        .Refresh
        'ZA01/10/2002 - check for the grid
        If oGrid.Name = ms_RFC_GRID_NAME Then
            mtSelectedGridCell.Col = .Col
            mtSelectedGridCell.Row = .Row
        Else
            mtSelectedRFOCell.Col = .Col
            mtSelectedRFOCell.Row = .Row
        End If
    
    End With

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "HandleGridClick")
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
Private Sub SetGridText(ByRef oGrid As MSFlexGrid)
'---------------------------------------------------------------------
'set grid text according to user selection
'---------------------------------------------------------------------
Dim oText As TextBox

' PN change 41
    On Error GoTo ErrHandler
    
    If oGrid.Name = ms_RFC_GRID_NAME Then
        Set oText = txtReasonGridEdit
    ElseIf oGrid.Name = ms_OVERRULE_GRID_NAME Then
        Set oText = txtOverruleEdit
    End If
    
    
    If oText.Visible Then
        oGrid = oText
        oText.Visible = False
        DoEvents
        oText.Text = vbNullString
    End If
    
    
'    If txtReasonGridEdit.Visible Then
'       gridReasons = txtReasonGridEdit
'       txtReasonGridEdit.Visible = False
'       DoEvents
'       txtReasonGridEdit.Text = vbNullString
'    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "SetGridText")
        Case OnErrorAction.Retry
            Resume
    End Select
End Sub

'---------------------------------------------------------------------
Private Sub DoGridEdit(nKeyAscii As Integer, oGrid As MSFlexGrid, oTextBox As TextBox)
'---------------------------------------------------------------------
' Edit a cell in the given grid
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    ' edit a cell
    
    'ZA 30/09/2002 - only edit grid if user clicked inside the grid
    If mbClickInGridCell Then
        With oTextBox
            'use correct font
            .FontName = oGrid.FontName
            .FontSize = oGrid.FontSize
    
           'position the edit box
            .Left = oGrid.CellLeft + oGrid.Left
            .Top = oGrid.CellTop + oGrid.Top
            .Width = oGrid.CellWidth
            .Height = oGrid.CellHeight
            .Visible = True
            
            Select Case nKeyAscii
               Case 0 To Asc(" ")
                  .Text = oGrid
                  .SelStart = 1000
               Case Else
                  .Text = Chr(nKeyAscii)
                  .SelStart = 1
            End Select
            
            On Error Resume Next
            .SetFocus
            On Error GoTo 0
        End With
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "DoGridEdit")
        Case OnErrorAction.Retry
            Resume
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub GridEdit(KeyAscii As Integer)
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    ' edit a cell
    
    'ZA 30/09/2002 - only edit grid if user clicked inside the grid
    If mbClickInGridCell Then
        With txtReasonGridEdit
            'use correct font
            .FontName = gridReasons.FontName
            .FontSize = gridReasons.FontSize
    
           'position the edit box
            .Left = gridReasons.CellLeft + gridReasons.Left
            .Top = gridReasons.CellTop + gridReasons.Top
            .Width = gridReasons.CellWidth
            .Height = gridReasons.CellHeight
            .Visible = True
            
            Select Case KeyAscii
               Case 0 To Asc(" ")
                  .Text = gridReasons
                  .SelStart = 1000
               Case Else
                  .Text = Chr(KeyAscii)
                  .SelStart = 1
            End Select
            
            On Error Resume Next
            .SetFocus
            On Error GoTo 0
        End With
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "GridEdit")
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
Private Sub RefreshRFCGrid()
'---------------------------------------------------------------------
' Refresh the Reasons For Change grid
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Call RefreshGridRows(moStudyDefinition.ReasonsForChange, gridReasons, "Reason ForChange")
    mbRefreshRFCGrid = False

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "RefreshRFCGrid")
        Case OnErrorAction.Retry
            Resume
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub RefreshGridRows(oReasons As clsReasonsForChange, oGrid As MSFlexGrid, sReasonHeader As String)
'---------------------------------------------------------------------
' Refresh the given grid with the given collection of Reasons For X
'---------------------------------------------------------------------
Dim oReason As clsReasonForChange
Dim lIndex As Long
Dim nRow As Integer
Dim nCol As Integer
            
    On Error GoTo ErrHandler
    
    lIndex = 0
    With oGrid
        nRow = .Row
        nCol = .Col
        
        If oReasons.Count > 0 Then
            Call LockWindow(oGrid)
            .Clear
           
            .Row = 0
            .Col = mcREASONCOL
            .ColWidth(mcREASONCOL) = 6500
            .Text = sReasonHeader
            .Col = mcINVISIBLECOL
            .ColWidth(mcINVISIBLECOL) = 0
            
           'when a new reason is created, it is set to RFC
           'so we need to have a dummy RFO
           .Rows = oReasons.CountValid + 1
            
            For Each oReason In oReasons
                If Not oReason.IsDeleted Then
                    lIndex = lIndex + 1
                    .Row = lIndex
                    If Not oReason.IsValid Then
                        .Col = mcREASONCOL
                        .CellBackColor = g_INVALID_BACKCOLOUR
                    End If
                    
                    .Col = mcREASONCOL
                    .Text = oReason.Reason
                
                    ' PN 28/09/99 left justify all data not just alpha data
                    .CellAlignment = flexAlignLeftCenter
                
                    ' set the cross reference between grid and objects
                    .RowData(.Row) = oReason.ReasonID
                End If
            
            Next oReason
            
            'TA 28/11/2000: last row empty therefore valid
            .Row = .Rows - 1
            .Col = mcREASONCOL
            If .Text = "" Then
                .CellBackColor = vbWindowBackground
            End If
            
            If .Rows - 1 > nRow Then
                .Row = nRow
            End If
            If .Cols - 1 > nCol Then
                .Col = nCol
            End If
            
            ' allow refresh grid screen
            Call UnlockWindow
        End If

    End With
        
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmStudyDefinition.RefreshGridRows"

End Sub

'---------------------------------------------------------------------
Private Sub RefreshOverruleGrid()
'---------------------------------------------------------------------
' Refresh the Reasons For Overrule grid
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Call RefreshGridRows(moStudyDefinition.ReasonsForOverrule, gridOverrule, "Reason For Overrule")
    mbRefreshRFOGrid = False

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "RefreshOverrulegrid")
        Case OnErrorAction.Retry
            Resume
    End Select
End Sub

'---------------------------------------------------------------------
Private Sub gridReasons_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    ' NCJ 14 Jun 06 - Ignore if not read/write
    If frmMenu.StudyAccessMode < sdReadWrite Then Exit Sub
    
    If gridReasons.Col = mcREASONCOL Then
        If KeyAscii = vbKeyReturn Then
            Call SetupNextCellSelection(gridReasons)
            Call gridReasons_Click
        End If
        
    Else
        Call GridEdit(KeyAscii)
        
    End If
    mbChanged = Not mbIsLoading

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "gridReasons_KeyPress")
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
Private Sub SetupNextCellSelection(ByRef oGrid As MSFlexGrid)
'---------------------------------------------------------------------
'---------------------------------------------------------------------
'ZA 30/09/2002 - added oGrid as parameter

On Error GoTo ErrHandler
 
    Call SetGridText(oGrid)
    Call SelectNextCell(oGrid)
    
    'ZA 01/10/2002 - set up next cell based on the selected grid
    If oGrid.Name = ms_OVERRULE_GRID_NAME Then
        mtSelectedRFOCell.Col = gridOverrule.Col
        mtSelectedRFOCell.Row = gridOverrule.Row
    Else
        mtSelectedGridCell.Col = gridReasons.Col
        mtSelectedGridCell.Row = gridReasons.Row
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "SetupNextCellSelection")
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
Private Sub gridReasons_Scroll()
'---------------------------------------------------------------------
On Error GoTo ErrHandler

    Call SetGridText(gridReasons)
    DoEvents
    gridReasons.Refresh

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "gridReasons_Scroll")
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
Private Sub txtReasonGridEdit_Change()
'---------------------------------------------------------------------
    
    ' copy the change across to the object
    ' let the object validate the data input
    On Error GoTo ErrHandler
    If txtReasonGridEdit.Visible Then
        Call UpdateReasonsWithEdits(gridReasons, txtReasonGridEdit, moStudyDefinition.ReasonsForChange)
        mbChanged = True
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtreasongridEdit_Change")
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
Private Sub UpdateReasonsWithEdits(oGrid As MSFlexGrid, oText As TextBox, oReasons As clsReasonsForChange)
'---------------------------------------------------------------------
' ensure that the changes are carried accross to the categories collection
'---------------------------------------------------------------------
Dim oReason As clsReasonForChange
Dim lPos As Long
Dim bIsValid As Boolean
    
    On Error GoTo ErrHandler
    
    If Not mbIsLoading Then
        ' retrieve the correct object using the ValueID field stored
        ' in the text control tag property
        'ZA 30/09/2002 - don't proceed if reason = nothing
        If oGrid.RowData(oGrid.Row) = 0 Then
             Exit Sub
        End If
        
        Set oReason = oReasons.GetItemFromID(oGrid.RowData(oGrid.Row))
        
        bIsValid = oReason.IsValid

        On Error GoTo InvalidChar
        ' the object encapsulates the rules for an RFO/RFC
        ' so let it catch the errors
        If oGrid.Col = mcREASONCOL Then
            oReason.Reason = oText.Text
        End If
    
        On Error GoTo ErrHandler

        If Not bIsValid And oReason.IsValid Then
            ' the category was not valid but is now
            ' so repaint the back color of the grid row
            oGrid.Col = 1
            oGrid.CellBackColor = vbWhite
        End If
        
        If oReason.IsDirty Then
            Call EnableOK(True)
        End If
    End If
    
    Set oReason = Nothing
    
    Exit Sub
    
InvalidChar:
    Beep
    lPos = oText.SelStart
    If oGrid.Col = mcREASONCOL Then
        oText.Text = oReason.Reason
        oGrid.Text = oReason.Reason
    End If
    If lPos > 0 Then lPos = lPos - 1
    oText.SelStart = lPos
    
    Set oReason = Nothing

    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "UpdateReasonsWithEdits")
            Case OnErrorAction.Retry
            Resume
    End Select
End Sub

''---------------------------------------------------------------------
'Private Sub UpdateReasonsWithUserEdits(ByRef oGrid As MSFlexGrid, ByVal enReasonType As eReasonType)
''---------------------------------------------------------------------
'' ensure that the changes are carried accross to the categories collection
''---------------------------------------------------------------------
'Dim oReason As clsReasonForChange
'Dim lPos As Long
'Dim bIsValid As Boolean
'
'    On Error GoTo InvalidChar
'
'    If Not mbIsLoading Then
'        ' retrieve the correct object using the ValueID field stored
'        ' in the text control tag property
'        'ZA 30/09/2002 - don't proceed if reason = nothing
'        If oGrid.RowData(oGrid.Row) = 0 Then
'             Exit Sub
'        End If
'        If enReasonType = eReasonType.ReasonForOverrule Then
'            Set oReason = moStudyDefinition.ReasonsForOverrule.GetItemFromID(oGrid.RowData(oGrid.Row))
'        Else
'            Set oReason = moStudyDefinition.ReasonsForChange.GetItemFromID(oGrid.RowData(oGrid.Row))
'        End If
'
'        'ZA 04/10/02
'       ' Debug.Print "Reason ID = " & oReason.ReasonID, ",  Type = " & oReason.ReasonType
'
'        'Set oReason = moStudyDefinition.ReasonsForChange.GetItemFromID(gridReasons.RowData(gridReasons.Row))
'        bIsValid = oReason.IsValid
'
'        ' the object encapsulates the rules for a category item
'        ' so let it catch the errors
'        If gridReasons.Col = mcREASONCOL Then
'            If oGrid.Name = ms_RFC_GRID_NAME Then
'                oReason.Reason = txtReasonGridEdit.Text
'            Else
'                oReason.Reason = txtOverruleEdit.Text
'            End If
'        End If
'
'        If Not bIsValid And oReason.IsValid Then
'            ' the category was not valid but is now
'            ' so repaint the back color of the grid row
'            'With gridReasons
'            With oGrid
'                .Col = 1
'                .CellBackColor = vbWhite
'            End With
'
'        End If
'
'        If oReason.IsDirty Then
'            Call EnableOK(True)
'        End If
'    End If
'    Set oReason = Nothing
'    Exit Sub
'
'InvalidChar:
'    Beep
'    lPos = txtReasonGridEdit.SelStart
'    If gridReasons.Col = mcREASONCOL Then
'        txtReasonGridEdit = oReason.Reason
'        gridReasons.Text = oReason.Reason
'    End If
'    If lPos > 0 Then lPos = lPos - 1
'    txtReasonGridEdit.SelStart = lPos
'
'End Sub


'---------------------------------------------------------------------
Private Sub txtReasonGridEdit_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    With gridReasons
        Select Case KeyCode
        Case vbKeyReturn
            '.SetFocus
            DoEvents
            If mbRefreshRFCGrid Then
                Call RefreshRFCGrid
            End If
            msGridWithFocus = ms_RFC_GRID_NAME
            txtAfterGridEditTab.SetFocus
            
         End Select
     
     End With
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtReasonGridEdit_KeyDown")
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
Private Sub txtReasonGridEdit_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
   'noise suppression
   
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    Else
        mbChanged = True
    End If

End Sub

'---------------------------------------------------------------------
Private Sub txtReasonGridEdit_LostFocus()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    If mbRefreshRFCGrid And Me.Visible Then
        Call RefreshRFCGrid
    End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmStudyDefinition.txtReasonGridEdit_LostFocus"
    
End Sub

'---------------------------------------------------------------------
Private Sub txtAfterGridEditTab_GotFocus()
'---------------------------------------------------------------------
' this event fires when the grid loses focus
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    ' force the grid to get focus again
    ' set the cell selected to the one after the current cell
    'ZA 30/09/2002
    If msGridWithFocus = ms_OVERRULE_GRID_NAME Then
        Call SetupNextCellSelection(gridOverrule)
        Call gridOverrule_Click
    ElseIf msGridWithFocus = ms_RFC_GRID_NAME Then
        Call SetupNextCellSelection(gridReasons)
        Call gridReasons_Click
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtAfterGridEditTab_GotFocus")
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
Private Sub PopulateRequiredTab()
'---------------------------------------------------------------------
'SDM 14/12/99 SR1177
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsRequiredData As ADODB.Recordset
Dim liTrialType As ListItem

    On Error GoTo ErrHandler
    
    'Populate list
    lsvRequired.ListItems.Clear
    lsvRequired.ColumnHeaders.Clear
    lsvRequired.ColumnHeaders.Add , , "Questions"
    sSQL = "SELECT " & _
           "DataItem.DataItemId, " & _
           "DataItem.DataItemName " & _
           "FROM " & _
           "DataItem, " & _
           "RequiredData " & _
           "WHERE " & _
           "DataItem.DataItemId = RequiredData.DataItemId AND " & _
           "DataItem.ClinicalTrialId = 0 AND " & _
           "DataItem.VersionId = 0 AND " & _
           "RequiredData.TrialTypeId = " & moStudyDefinition.TrialTypes.Key(cboTrialType.Text)
    Set rsRequiredData = New ADODB.Recordset
    rsRequiredData.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsRequiredData.EOF Then
        Set rsRequiredData = Nothing
        Exit Sub
    End If
    rsRequiredData.MoveFirst
    Do While Not rsRequiredData.EOF
        lsvRequired.ListItems.Add , Trim(Str(rsRequiredData!DataItemId) & "K"), rsRequiredData!DataItemName
        rsRequiredData.MoveNext
    Loop
    
    'Set check marks
    sSQL = "SELECT " & _
           "DataItem.CopiedFromDataItemId " & _
           "FROM " & _
           "DataItem, " & _
           "CRFElement, " & _
           "RequiredData " & _
           "WHERE " & _
           "DataItem.CopiedFromDataItemId = RequiredData.DataItemId AND " & _
           "CRFElement.DataItemId = DataItem.DataItemId AND " & _
           "CRFElement.ClinicalTrialId = " & mlClinicalTrialId & " AND " & _
           "CRFElement.VersionId = " & mnVersionId & " AND " & _
           "DataItem.CopiedFromClinicalTrialId = 0 AND " & _
           "DataItem.CopiedFromVersionId = 0 AND " & _
           "DataItem.ClinicalTrialId = " & mlClinicalTrialId & " AND " & _
           "DataItem.VersionId = " & mnVersionId
    Set rsRequiredData = New ADODB.Recordset
    rsRequiredData.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsRequiredData.EOF Then
        lsvRequired.SelectedItem.Selected = False
        Set rsRequiredData = Nothing
        Exit Sub
    End If
    rsRequiredData.MoveFirst
    Do While Not rsRequiredData.EOF
        lsvRequired.ListItems(rsRequiredData!CopiedFromDataItemId & "K").Checked = True
        rsRequiredData.MoveNext
    Loop
    Set rsRequiredData = Nothing
    
    lsvRequired.SelectedItem.Selected = False
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cboTrialType_Click")
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
Private Sub LoadCTCSchemeCombo()
'---------------------------------------------------------------------
' load combo with cTC schemes
' top is blank and has itemdata code of 0
'---------------------------------------------------------------------
Dim rsCTC As ADODB.Recordset

    Set rsCTC = New ADODB.Recordset
    rsCTC.Open "SELECT CTCSchemeCode FROM CTCScheme ORDER by CTCSchemeCode", MacroADODBConnection
    
    cboSchemes.Clear
    cboSchemes.AddItem ""
    Do While Not rsCTC.EOF
        cboSchemes.AddItem rsCTC!CTCSchemeCode
        rsCTC.MoveNext
    Loop

End Sub

'----------------------------------------------------------------------------------------
Private Sub cboeForm_Click()
'----------------------------------------------------------------------------------------
' Select the trigger eForm for registration
'----------------------------------------------------------------------------------------

    If Not mbIsLoading Then
        moRegistration.TriggerFormId = cboeForm.ItemData(cboeForm.ListIndex)
        mbChanged = True
    End If
    cboeForm.Enabled = (cboVisit.ItemData(cboVisit.ListIndex) <> 0)
    
End Sub

'----------------------------------------------------------------------------------------
Private Sub cboVisit_Click()
'----------------------------------------------------------------------------------------
' Select the trigger visit for Registration
'----------------------------------------------------------------------------------------
Dim lVisitId As Long

    lVisitId = cboVisit.ItemData(cboVisit.ListIndex)
    
    If Not mbIsLoading Then
        moRegistration.TriggerVisitId = lVisitId
        mbChanged = True
    End If
    If lVisitId = 0 Then
        cboeForm.Clear
        cboeForm.Enabled = False
    Else
        cboeForm.Enabled = True
        LoadFormCombo lVisitId
    End If
    
End Sub

'----------------------------------------------------------------------------------------
Private Sub LoadFormCombo(lVisitId As Long)
'----------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------
Dim tblForms As clsDataTable
Dim sSQL As String

    'TA 07/04/2003: corrected SQL so onlt forms in visit are returned
    sSQL = "SELECT DISTINCT CRFPage.CRFPageCode, CRFPage.CRFPageId FROM CRFPage, StudyVisitCRFPage" _
                & " WHERE CRFPage.ClinicalTrialId = StudyVisitCRFPage.ClinicalTrialId" _
                & " AND CRFPage.CRFPageId=StudyvisitCRFPage.CRFPageId" _
                & " AND CRFPage.ClinicalTrialId = " & mlClinicalTrialId _
                & " AND VisitId = " & lVisitId _
                & " ORDER BY CRFPage.CRFPageCode"
                
    
    Set tblForms = TableFromSQL(sSQL)
    
    Call TabletoCombo(cboeForm, tblForms)
    ListCtrl_Pick cboeForm, Val(tblForms(1, 2))


End Sub

'----------------------------------------------------------------------------------------
Private Sub LoadVisitCombo()
'----------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------
Dim tblVisits As clsDataTable
Dim sSQL As String

    'TA 11/01/2001: visit name changed to visit code
    sSQL = "SELECT DISTINCT StudyVisit.VisitCode, StudyVisit.VisitId FROM StudyVisit, StudyVisitCRFPage" _
                & " WHERE StudyVisit.ClinicalTrialId = " & mlClinicalTrialId _
                & " AND StudyVisit.ClinicalTrialId = StudyVisitCRFPage.ClinicalTrialId" _
                & " AND StudyVisit.VisitId = StudyVisitCRFPage.VisitId"
    
    Set tblVisits = TableFromSQL(sSQL)
    
    'TA 4/12/00: "none" to ""
    If tblVisits.Rows > 0 Then
        tblVisits.Insert RecordBuild("", "0"), 1
    Else
        tblVisits.Add RecordBuild("", "0")
    End If
    Call TabletoCombo(cboVisit, tblVisits)
    ListCtrl_Pick cboVisit, 0

End Sub

'----------------------------------------------------------------------------------------
Private Sub chkBaseNumbering_Click(Index As Integer)
'----------------------------------------------------------------------------------------
    
    If Index = m_INDEX_PREFIX Then
        moRegistration.UsePrefix = (chkBaseNumbering(Index).Value = vbChecked)
    Else
        moRegistration.UseSuffix = (chkBaseNumbering(Index).Value = vbChecked)
    End If
    mbChanged = Not mbIsLoading

End Sub

'----------------------------------------------------------------------------------------'
Private Sub DisplayRegistration()
'----------------------------------------------------------------------------------------'
' Fill the registration tab with the current details for this study
' NCJ 26 Feb 07 - Bug 2029 - Check visit and eForm exist
'----------------------------------------------------------------------------------------'
Dim bCanEdit As Boolean

    mbIsLoading = True
    
    ' NCJ 14 Jun 06 - Find out if read-only
    bCanEdit = (frmMenu.StudyAccessMode >= sdReadWrite)
    
    Set moRegistration = New clsRegistration
    moRegistration.Load ClinicalTrialId, VersionId
    LoadVisitCombo
   
    With moRegistration
        txtStartNo.Text = .StartNumber
        txtNumberWidth.Text = .NumberWidth
        txtRegExpression(m_INDEX_PREFIX).Text = .Prefix
        If .UsePrefix Then
            chkBaseNumbering(m_INDEX_PREFIX).Value = vbChecked
        Else
            chkBaseNumbering(m_INDEX_PREFIX).Value = vbUnchecked
        End If
        txtRegExpression(m_INDEX_SUFFIX).Text = .Suffix
        If .UseSuffix Then
            chkBaseNumbering(m_INDEX_SUFFIX).Value = vbChecked
        Else
            chkBaseNumbering(m_INDEX_SUFFIX).Value = vbUnchecked
        End If
          
        Set moRegInputGrid = New clsInputGrid
        Set moRegInputGrid.mMsFlexGrid = flxReg
        Set moRegInputGrid.mTextbox = txtRegCode
        Set moRegInputGrid.mcmdButton = cmdRegCondition
        'initalise grid with button shown for column 2
        moRegInputGrid.Init .RegistrationConditions, 2, bCanEdit
        flxReg.ColWidth(1) = flxReg.Parent.TextWidth("WWWWWWWWWWWWWWW")
        flxReg.ColWidth(2) = flxReg.Width - (flxReg.ColWidth(1) + 720 + cmdRegCondition.Width)
        
        Set moUniqueInputGrid = New clsInputGrid
        Set moUniqueInputGrid.mMsFlexGrid = flxUnique
        Set moUniqueInputGrid.mTextbox = txtUniqueCode
        Set moUniqueInputGrid.mcmdButton = cmdUniqueExpression
        'initalise grid with button shown for column 2
        moUniqueInputGrid.Init .UniquenessChecks, 2, bCanEdit
        flxUnique.ColWidth(1) = flxUnique.Parent.TextWidth("WWWWWWWWWWWWWWW")
        flxUnique.ColWidth(2) = flxUnique.Width - (flxUnique.ColWidth(1) + 720 + cmdUniqueExpression.Width)
        
        ' NCJ 26 Feb 07 - Bug 2029 - Check visit/eForm exist
        If .TriggerVisitId <> 0 And .TriggerFormId <> 0 Then
            ' Visit and form previously selected - check they still exist
            If VisitCodeFromId(mlClinicalTrialId, .TriggerVisitId) <> "" And CRFPageCodeFromId(mlClinicalTrialId, .TriggerFormId) <> "" Then
                ListCtrl_Pick cboVisit, .TriggerVisitId
                ListCtrl_Pick cboeForm, .TriggerFormId
            Else
                ' Non existent visit/eForm
                ListCtrl_Pick cboVisit, 0
                .TriggerVisitId = 0
                .TriggerFormId = 0
                mbChanged = True
            End If
        End If
        
        mbIsLoading = False
      
' NCJ 26 Feb 07 - Bug 2029 - Code moved above
'        If .TriggerVisitId <> 0 Then
'            'only select eForm if visit is not 0 (none)
'            ListCtrl_Pick cboeForm, .TriggerFormId
'        End If
        
        ' NCJ 14 Jun 06 - Now disable stuff if read-only
        cboeForm.Enabled = bCanEdit
        cboVisit.Enabled = bCanEdit
        chkBaseNumbering(m_INDEX_PREFIX).Enabled = bCanEdit
        chkBaseNumbering(m_INDEX_SUFFIX).Enabled = bCanEdit
        txtStartNo.Enabled = bCanEdit
        txtNumberWidth.Enabled = bCanEdit
        txtRegExpression(m_INDEX_PREFIX).Enabled = bCanEdit
        txtRegExpression(m_INDEX_SUFFIX).Enabled = bCanEdit
        cmdRegCondition.Enabled = bCanEdit
        cmdRegExpression(m_INDEX_PREFIX).Enabled = bCanEdit
        cmdRegExpression(m_INDEX_SUFFIX).Enabled = bCanEdit
        cmdUniqueExpression.Enabled = bCanEdit
        cmdRRServerSettings.Enabled = bCanEdit
    End With

    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub moRegistration_HasChanges(bIsDirty As Boolean)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

        Call EnableOK(bIsDirty)

End Sub

'----------------------------------------------------------------------------------------'
Private Sub moRegistration_IsValid(bValid As Boolean)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'
Dim i As Long
    'disable all tabs except registration tab if registration not valid
    If SSTab1.Tab = m_TAB_REGISTRATION Then
        For i = 0 To SSTab1.Tabs - 1
            'disable all tabs except registration tab if registration not valid
            
            SSTab1.TabEnabled(i) = (Not bValid) Or (i = m_TAB_REGISTRATION)
        Next
    End If
    Call EnableOK(bValid)
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub moRegistration_ValidSuffix(bIsValid As Boolean)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    Call ColourControl(txtRegExpression(m_INDEX_SUFFIX), bIsValid)
 
End Sub

'----------------------------------------------------------------------------------------'
Private Sub moRegistration_ValidNumberWidth(bIsValid As Boolean)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    Call ColourControl(txtNumberWidth, bIsValid)
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub moRegistration_ValidPrefix(bIsValid As Boolean)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    Call ColourControl(txtRegExpression(m_INDEX_PREFIX), bIsValid)
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub moRegistration_ValidStartNumber(bIsValid As Boolean)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    Call ColourControl(txtStartNo, bIsValid)
    
End Sub


'----------------------------------------------------------------------------------------
Private Sub cmdRegExpression_Click(Index As Integer)
'----------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------

    If Index = m_INDEX_PREFIX Then
        frmCUIFunctionEditor.Initialise txtRegExpression(Index), "Expression", _
                        (frmMenu.StudyAccessMode >= sdReadWrite), "Prefix"
    Else
        frmCUIFunctionEditor.Initialise txtRegExpression(Index), "Expression", _
                        (frmMenu.StudyAccessMode >= sdReadWrite), "Suffix"
    End If


End Sub

'----------------------------------------------------------------------------------------
Private Sub cmdRegCondition_Click()
'----------------------------------------------------------------------------------------

    frmCUIFunctionEditor.Initialise txtRegCode, "Condition", _
                        (frmMenu.StudyAccessMode >= sdReadWrite), "Registration Condition"

End Sub

'----------------------------------------------------------------------------------------
Private Sub cmduniqueExpression_Click()
'----------------------------------------------------------------------------------------

    frmCUIFunctionEditor.Initialise txtUniqueCode, "Expression", _
                    (frmMenu.StudyAccessMode >= sdReadWrite), "Unique Check Expression"


End Sub

'----------------------------------------------------------------------------------------
Private Sub txtRegExpression_Change(Index As Integer)
'----------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------

    If Index = m_INDEX_PREFIX Then
        moRegistration.Prefix = txtRegExpression(Index).Text
    Else
        moRegistration.Suffix = txtRegExpression(Index).Text
    End If
    mbChanged = Not mbIsLoading

End Sub

'----------------------------------------------------------------------------------------
Private Sub txtRegCode_Change()
'----------------------------------------------------------------------------------------
' validation on registration conditions (not in class as there is no class)
'----------------------------------------------------------------------------------------
Dim lOtherCol As Long

    Select Case flxReg.Col
    Case m_COL_CODE
        Call ValidTextBox(txtRegCode, ttCode)
        lOtherCol = m_COL_AREZZO
    Case m_COL_AREZZO
        If gclmGuideline.IsValidCondition(txtRegCode.Text) Then
            txtRegCode.BackColor = vbWindowBackground
        Else
            txtRegCode.BackColor = g_INVALID_BACKCOLOUR
        End If
        lOtherCol = m_COL_CODE
    End Select
    With flxReg
        If (.Row = .Rows - 1) _
                And .TextMatrix(.Row, lOtherCol) = "" Then
            'last row empty so not invalid
           txtRegCode.BackColor = vbWindowBackground
        End If
    End With
    'if we are not loading then there are changes
    moRegistration.IsDirty = Not mbIsLoading
    mbChanged = Not mbIsLoading

End Sub

'----------------------------------------------------------------------------------------
Private Sub flxReg_LeaveCell()
'----------------------------------------------------------------------------------------
' validation on registration conditions (not in class as there is no class)
'----------------------------------------------------------------------------------------

    If frmMenu.StudyAccessMode < sdReadWrite Then Exit Sub
    
    With flxReg
        Select Case .Col
        Case m_COL_CODE
            If ValidText(.Text, ttCode) Then
                .CellBackColor = vbWindowBackground
            Else
                .CellBackColor = g_INVALID_BACKCOLOUR
            End If
        Case m_COL_AREZZO
            If gclmGuideline.IsValidCondition(.Text) Then
                .CellBackColor = vbWindowBackground
            Else
                .CellBackColor = g_INVALID_BACKCOLOUR
            End If
        End Select
        
        If .Col = m_COL_CODE Or .Col = m_COL_AREZZO Then
            If (.Row = .Rows - 1) _
                    And .TextMatrix(.Row, m_COL_CODE) = "" _
                    And .TextMatrix(.Row, m_COL_AREZZO) = "" Then
                'last row empty so not invalid
               .CellBackColor = vbWindowBackground
            End If
        End If
    End With
    
    'if we are not loading then there are changes
    moRegistration.IsDirty = Not mbIsLoading
    mbChanged = Not mbIsLoading

End Sub

'----------------------------------------------------------------------------------------
Private Sub txtUniqueCode_Change()
'----------------------------------------------------------------------------------------
' validation on Uniqueness expressions (not in class as there is no class)
'----------------------------------------------------------------------------------------
Dim lOtherCol As Long

    Select Case flxUnique.Col
    Case m_COL_CODE
        Call ValidTextBox(txtUniqueCode, ttCode)
        lOtherCol = m_COL_AREZZO
    Case m_COL_AREZZO
        If gclmGuideline.IsValidExpression(txtUniqueCode.Text) Then
            txtUniqueCode.BackColor = vbWindowBackground
        Else
            txtUniqueCode.BackColor = g_INVALID_BACKCOLOUR
        End If
        lOtherCol = m_COL_CODE
    End Select
    With flxUnique
        If (.Row = .Rows - 1) _
                And .TextMatrix(.Row, lOtherCol) = "" Then
            'last row empty so not invalid
           txtUniqueCode.BackColor = vbWindowBackground
        End If
    End With
    'if we are not loading then there are changes
    moRegistration.IsDirty = Not mbIsLoading
    mbChanged = Not mbIsLoading

End Sub

Private Sub flxUnique_LeaveCell()
'----------------------------------------------------------------------------------------
' validation on uniqueness expr (not in class as there is no class)
'----------------------------------------------------------------------------------------
        
    If frmMenu.StudyAccessMode < sdReadWrite Then Exit Sub
    
    With flxUnique
        Select Case .Col
        Case m_COL_CODE
            If ValidText(.Text, ttCode) Then
                .CellBackColor = vbWindowBackground
            Else
                .CellBackColor = g_INVALID_BACKCOLOUR
            End If
        Case m_COL_AREZZO
            If gclmGuideline.IsValidExpression(.Text) Then
                .CellBackColor = vbWindowBackground
            Else
                .CellBackColor = g_INVALID_BACKCOLOUR
            End If
        End Select
        If .Col = m_COL_CODE Or .Col = m_COL_AREZZO Then
            If (.Row = .Rows - 1) _
                    And .TextMatrix(.Row, m_COL_CODE) = "" _
                    And .TextMatrix(.Row, m_COL_AREZZO) = "" Then
                'last row empty so not invalid
               .CellBackColor = vbWindowBackground
            End If
        End If
    End With
    'if we are not loading then there are changes
    moRegistration.IsDirty = Not mbIsLoading
    mbChanged = Not mbIsLoading

End Sub


'----------------------------------------------------------------------------------------
Private Sub txtNumberWidth_Change()
'----------------------------------------------------------------------------------------

    moRegistration.NumberWidth = txtNumberWidth.Text
    mbChanged = Not mbIsLoading

End Sub

'----------------------------------------------------------------------------------------
Private Sub txtStartNo_Change()
'----------------------------------------------------------------------------------------

    moRegistration.StartNumber = txtStartNo.Text
    mbChanged = Not mbIsLoading

End Sub

'----------------------------------------------------------------------------------------
Private Function SaveRegistration()
'----------------------------------------------------------------------------------------
'TA 17/11/2000: save registration details (after getting new tables from grid
'----------------------------------------------------------------------------------------

    Set moRegistration.UniquenessChecks = moUniqueInputGrid.Table
    Set moRegistration.RegistrationConditions = moRegInputGrid.Table
    moRegistration.Save
    
End Function

'----------------------------------------------------------------------------------------
Private Function RegGridsAreValid() As Boolean
'----------------------------------------------------------------------------------------
Dim i As Long
Dim bValid As Boolean
Dim sCode As String
Dim sArezzo As String
Dim sCodeList As String

    RegGridsAreValid = True
    
    sCodeList = "|"
    With flxReg
        For i = 1 To .Rows - 2
            sCode = .TextMatrix(i, m_COL_CODE)
            sArezzo = .TextMatrix(i, m_COL_AREZZO)
            
            If i = .Row Then
                'use text control beacsue grid might not yet have been updated
                Select Case .Col
                Case m_COL_CODE: sCode = txtRegCode.Text
                Case m_COL_AREZZO: sArezzo = txtRegCode.Text
                End Select
            End If
            
            RegGridsAreValid = RegGridsAreValid And ValidText(sCode, ttCode) _
                        And (gclmGuideline.IsValidCondition(sArezzo))
                        
            If InStr(sCodeList, "|" & sCode & "|") Then
                .ToolTipText = "The registration condition code '" & sCode & "' exists more than once"
                txtRegCode.ToolTipText = "The registration condition code '" & sCode & "' exists more than once"
'                DialogError "The uniqueness expression code " & sCode & " exists more than once"
                RegGridsAreValid = False
            Else
                .ToolTipText = ""
                txtRegCode.ToolTipText = ""
            End If
            
            If Not RegGridsAreValid Then Exit Function
            
            If Trim(sCode) <> "" Then
                sCodeList = sCodeList & sCode & "|"
            End If
        Next
    End With

    sCodeList = "|"
    With flxUnique
        For i = 1 To .Rows - 2
            sCode = .TextMatrix(i, m_COL_CODE)
            sArezzo = .TextMatrix(i, m_COL_AREZZO)
            
            If i = .Row Then
                'use text control beacsue grid might not yet have been updated
                Select Case .Col
                Case m_COL_CODE: sCode = txtUniqueCode.Text
                Case m_COL_AREZZO: sArezzo = txtUniqueCode.Text
                End Select
            End If
            
            RegGridsAreValid = RegGridsAreValid And ValidText(sCode, ttCode) _
                        And (gclmGuideline.IsValidExpression(sArezzo))
    
            If InStr(sCodeList, "|" & sCode & "|") Then
                .ToolTipText = "The uniqueness expression code '" & sCode & "' exists more than once"
                txtUniqueCode.ToolTipText = "The uniqueness expression code '" & sCode & "' exists more than once"
'                DialogError "The uniqueness check code " & sCode & " exists more than once"
                RegGridsAreValid = False
            Else
                .ToolTipText = ""
                txtUniqueCode.ToolTipText = ""
            End If
            
            If Not RegGridsAreValid Then Exit Function
            
            If Trim(sCode) <> "" Then
                'if not blank then add to checklist
                sCodeList = sCodeList & sCode & "|"
            End If
        Next
    End With

End Function

'----------------------------------------------------------------------------------------
Private Sub cmdRRServerSettings_Click()
'----------------------------------------------------------------------------------------
' show RR Server Connection Settings
'----------------------------------------------------------------------------------------

    Call frmRRConnection.Display(ClinicalTrialId, VersionId)
    

End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtUniqueCode_LostFocus()
'----------------------------------------------------------------------------------------'

    Call EnableOK(True)
    
End Sub

'----------------------------------------------------------------------------------------'
Private Function GetGridRows(oReasons As clsReasonsForChange) As Integer
'----------------------------------------------------------------------------------------'
'set the number of rows for RFC/RFO grids
'----------------------------------------------------------------------------------------'
Dim oReason As clsReasonForChange
Dim nRows As Integer
    
    nRows = 0
    
    For Each oReason In oReasons
        If Not oReason.IsDeleted Then
            nRows = nRows + 1
        End If
    Next
    
    GetGridRows = nRows

End Function

